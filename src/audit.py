#!/usr/bin/env python3
"""
Audit a tree of .pptx files without opening PowerPoint.

Reads the OOXML directly and reports, per deck: slide count, slide dimensions,
fonts (and which of them will be substituted at export time), animation build
counts, embedded playable media, speaker notes, and likely duplicates.

    python3 audit.py <root> [--json out.json]

Everything here is read-only and needs nothing but the Python standard library.
"""
import json, math, os, re, sys, zipfile
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import buildstates            # single definition of what counts as a build state

# A 333 MB deck opened and then sat at 0% CPU indefinitely, unresponsive even to
# Accessibility queries; PowerPoint never recovered on its own. Decks this size
# are worth pulling out and handling on their own rather than discovering them
# an hour into a run, so audit flags them.
OVERSIZED_BYTES = 200 * 1000 * 1000

VIDEO = {'.mp4', '.mov', '.m4v', '.avi', '.wmv', '.mpg', '.mpeg', '.mkv', '.webm', '.asf', '.3gp'}
AUDIO = {'.mp3', '.wav', '.m4a', '.aac', '.wma', '.aiff', '.au', '.mid', '.midi'}
SKIP_FONTS = {'Wingdings', 'Wingdings 2', 'Wingdings 3', 'Webdings', 'Symbol', 'Arial Unicode MS'}
EMU_PER_PX = 9525

_SHAPE = re.compile(rb'<p:sp\b.*?</p:sp>', re.S)


# --------------------------------------------------------------------------
# fonts available to PowerPoint: system faces plus the ones Office ships
# --------------------------------------------------------------------------
def _norm(s):
    s = re.sub(r'\.(ttf|otf|ttc|dfont)$', '', s, flags=re.I)
    s = re.sub(r'[-_ ]?(Regular|Bold|Italic|BoldItalic|Oblique|Light|Medium|SemiBold|'
               r'Semibold|DemiBold|Black|Heavy|Thin|ExtraLight|ExtraBold|Condensed|'
               r'Roman|Book|MT|LT|Std|Pro)$', '', s)
    return re.sub(r'[^a-z0-9]', '', s.lower())


def available_fonts():
    names = set()
    dirs = ['/System/Library/Fonts', '/Library/Fonts', os.path.expanduser('~/Library/Fonts')]
    for app in ('PowerPoint', 'Word', 'Excel'):
        dirs.append('/Applications/Microsoft %s.app/Contents/Resources/DFonts' % app)
    # Walk, don't listdir: macOS keeps Arial, Verdana, Trebuchet MS, Georgia,
    # Gill Sans and ~290 other faces in /System/Library/Fonts/Supplemental/.
    # A non-recursive scan misses all of them and reports installed fonts as
    # about to be substituted - the opposite of what this command is for.
    for d in dirs:
        for _root, _subdirs, files in os.walk(d):
            for f in files:
                if not f.lower().endswith(('.ttf', '.otf', '.ttc', '.dfont')):
                    continue
                names.add(_norm(f))
                names.add(re.sub(r'[^a-z0-9]', '', os.path.splitext(f)[0].lower()))
    return names


# --------------------------------------------------------------------------
# per-deck inspection
# --------------------------------------------------------------------------
def gif_frames(data):
    if not data.startswith((b'GIF87a', b'GIF89a')):
        return None
    i, flags, n = 13, data[10], 0
    if flags & 0x80:
        i += 3 * (2 ** ((flags & 7) + 1))
    while i < len(data):
        b = data[i]
        if b == 0x3B:
            break
        if b == 0x2C:
            n += 1
            i += 10
            lf = data[i - 1]
            if lf & 0x80:
                i += 3 * (2 ** ((lf & 7) + 1))
            i += 1
            while i < len(data) and data[i]:
                i += data[i] + 1
            i += 1
        elif b == 0x21:
            i += 2
            while i < len(data) and data[i]:
                i += data[i] + 1
            i += 1
        else:
            i += 1
    return n


def notes_text(xml):
    for sp in _SHAPE.findall(xml):
        if not re.search(rb'<p:ph\b[^>]*type="body"', sp):
            continue
        paras = []
        for p in re.findall(rb'<a:p\b.*?</a:p>', sp, re.S):
            line = "".join(t.decode('utf8', 'ignore')
                           for t in re.findall(rb'<a:t>(.*?)</a:t>', p, re.S)).strip()
            if line:
                paras.append(line)
        text = "\n".join(paras).strip()
        if text and not re.fullmatch(r'[\d\s]*', text):
            return text
    return ""


def ratio_label(r):
    if r is None:
        return None
    if abs(r - 16 / 9) < 0.01:
        return "16:9"
    if abs(r - 4 / 3) < 0.01:
        return "4:3"
    if abs(r - math.sqrt(2)) < 0.02:
        return "A4"
    return "%.2f:1" % r


def inspect(path, avail):
    rec = {"path": path, "name": os.path.basename(path),
           "bytes": os.path.getsize(path), "slides": 0,
           "width_px": None, "height_px": None, "aspect": None,
           "fonts": [], "fonts_embedded": [], "fonts_missing": [],
           "animated_slides": 0, "build_states": 0,
           "gifs": 0, "videos": 0, "audio": 0, "external_media": [],
           "media_bytes": 0, "noted_slides": 0, "notes_chars": 0,
           "error": None}
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            slides = [n for n in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)]
            rec["slides"] = len(slides)

            if 'ppt/presentation.xml' in names:
                pres = z.read('ppt/presentation.xml').decode('utf8', 'ignore')
                m = re.search(r'<p:sldSz\b[^>]*>', pres)
                if m:
                    cx = re.search(r'cx="(\d+)"', m.group(0))
                    cy = re.search(r'cy="(\d+)"', m.group(0))
                    if cx and cy:
                        cxv, cyv = int(cx.group(1)), int(cy.group(1))
                        rec["width_px"] = round(cxv / EMU_PER_PX)
                        rec["height_px"] = round(cyv / EMU_PER_PX)
                        rec["aspect"] = ratio_label(cxv / cyv)
                for fm in re.finditer(r'<p:embeddedFont>\s*<p:font[^>]*typeface="([^"]+)"', pres):
                    rec["fonts_embedded"].append(fm.group(1))

            faces = set()
            for n in names:
                if (n.startswith('ppt/slides/slide') or n.startswith('ppt/slideLayouts/')
                        or n.startswith('ppt/slideMasters/')):
                    xml = z.read(n).decode('utf8', 'ignore')
                    for fm in re.finditer(r'<a:latin[^>]*typeface="([^"]+)"', xml):
                        t = fm.group(1)
                        if t and not t.startswith('+') and t not in SKIP_FONTS:
                            faces.add(t)
                elif n.startswith('ppt/theme/'):
                    xml = z.read(n).decode('utf8', 'ignore')
                    for block in re.findall(r'<a:(?:major|minor)Font>.*?</a:(?:major|minor)Font>',
                                            xml, re.S):
                        for fm in re.finditer(r'<a:latin[^>]*typeface="([^"]+)"', block):
                            t = fm.group(1)
                            if t and not t.startswith('+') and t not in SKIP_FONTS:
                                faces.add(t)
            rec["fonts"] = sorted(faces)
            emb = {_norm(e) for e in rec["fonts_embedded"]}
            emb |= {re.sub(r'[^a-z0-9]', '', e.lower()) for e in rec["fonts_embedded"]}
            rec["fonts_embedded"] = sorted(set(rec["fonts_embedded"]))
            missing = []
            for f in rec["fonts"]:
                a, b = _norm(f), re.sub(r'[^a-z0-9]', '', f.lower())
                if a in emb or b in emb or a in avail or b in avail:
                    continue
                missing.append(f)
            rec["fonts_missing"] = missing

            # Animation builds. Count what `export --states` will actually
            # produce, not every click step: a step animates a build only if it
            # has an entrance or exit targeting a shape. Motion-path and
            # emphasis effects are click steps too, but they have no meaningful
            # still frame, so they yield no extra page. Counting raw
            # clickEffect nodes promised pages the export never delivered
            # (one deck here forecast 42 and produced none).
            for n in slides:
                xml = z.read(n)
                if b'<p:timing' not in xml:
                    continue
                clicks = len(buildstates.click_steps(xml))
                if clicks:
                    rec["animated_slides"] += 1
                    rec["build_states"] += clicks

            # media
            for n in names:
                if not n.startswith('ppt/media/'):
                    continue
                ext = os.path.splitext(n)[1].lower()
                if ext in VIDEO:
                    rec["videos"] += 1; rec["media_bytes"] += z.getinfo(n).file_size
                elif ext in AUDIO:
                    rec["audio"] += 1; rec["media_bytes"] += z.getinfo(n).file_size
                elif ext == '.gif':
                    if (gif_frames(z.read(n)) or 0) > 1:
                        rec["gifs"] += 1; rec["media_bytes"] += z.getinfo(n).file_size

            # externally linked media (YouTube etc.) - recorded, never downloaded
            for n in names:
                m = re.fullmatch(r'ppt/slides/_rels/slide(\d+)\.xml\.rels', n)
                if not m:
                    continue
                x = z.read(n).decode('utf8', 'ignore')
                for tag in re.findall(r'<Relationship\b[^>]*TargetMode="External"[^>]*/>', x):
                    t = re.search(r'Target="([^"]+)"', tag)
                    if not t:
                        continue
                    url = t.group(1).replace('&amp;', '&')
                    if (os.path.splitext(url)[1].lower() in VIDEO | AUDIO
                            or re.search(r'youtube|youtu\.be|vimeo|wistia|loom', url, re.I)):
                        rec["external_media"].append({"slide": int(m.group(1)), "url": url})

            # speaker notes
            for n in names:
                m = re.fullmatch(r'ppt/slides/_rels/slide(\d+)\.xml\.rels', n)
                if not m:
                    continue
                rels = z.read(n).decode('utf8', 'ignore')
                t = re.search(r'Target="\.\./(notesSlides/notesSlide\d+\.xml)"', rels)
                if not t or 'ppt/' + t.group(1) not in names:
                    continue
                text = notes_text(z.read('ppt/' + t.group(1)))
                if text:
                    rec["noted_slides"] += 1
                    rec["notes_chars"] += len(text)
    except Exception as e:
        rec["error"] = "%s: %s" % (type(e).__name__, e)
    return rec


def find_decks(root):
    out = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith('.')]
        for fn in filenames:
            if fn.startswith(('.', '~$')):
                continue
            if fn.lower().endswith('.pptx'):
                out.append(os.path.join(dirpath, fn))
    return sorted(out)


def audit(root):
    avail = available_fonts()
    decks = [inspect(p, avail) for p in find_decks(root)]

    # duplicate detection: same byte size and slide count
    groups = defaultdict(list)
    for d in decks:
        groups[(d["bytes"], d["slides"])].append(d["path"])
    for d in decks:
        peers = groups[(d["bytes"], d["slides"])]
        d["duplicate_of"] = peers[0] if len(peers) > 1 and peers[0] != d["path"] else None
        d["duplicate_group_size"] = len(peers) if len(peers) > 1 else 1

    risky = [d for d in decks if d["fonts_missing"]]
    face_counts = Counter()
    for d in risky:
        face_counts.update(d["fonts_missing"])
    redundant = sum(len(v) - 1 for v in groups.values() if len(v) > 1)

    return {
        "root": os.path.abspath(root),
        "decks": decks,
        "summary": {
            "decks": len(decks),
            "slides": sum(d["slides"] for d in decks),
            "bytes": sum(d["bytes"] for d in decks),
            "unreadable": sum(1 for d in decks if d["error"]),
            "font_risk_decks": len(risky),
            "missing_faces": face_counts.most_common(),
            "decks_embedding_fonts": sum(1 for d in decks if d["fonts_embedded"]),
            "animated_decks": sum(1 for d in decks if d["animated_slides"]),
            "animated_slides": sum(d["animated_slides"] for d in decks),
            "build_states": sum(d["build_states"] for d in decks),
            "decks_with_media": sum(1 for d in decks
                                    if d["gifs"] or d["videos"] or d["audio"] or d["external_media"]),
            "animated_gifs": sum(d["gifs"] for d in decks),
            "videos": sum(d["videos"] for d in decks),
            "external_media": sum(len(d["external_media"]) for d in decks),
            "noted_slides": sum(d["noted_slides"] for d in decks),
            "notes_chars": sum(d["notes_chars"] for d in decks),
            "redundant_copies": redundant,
            "oversized_decks": sum(1 for d in decks if d["bytes"] >= OVERSIZED_BYTES),
            "oversized_bytes_threshold": OVERSIZED_BYTES,
        },
    }


def main(argv):
    if not argv or argv[0] in ('-h', '--help'):
        print(__doc__.strip()); return 0
    root = argv[0]
    if not os.path.isdir(root):
        sys.stderr.write("not a directory: %s\n" % root); return 2
    out = None
    if '--json' in argv:
        i = argv.index('--json')
        out = argv[i + 1] if i + 1 < len(argv) else None
    result = audit(root)
    if out:
        with open(out, 'w') as f:
            json.dump(result, f, indent=1)
    else:
        json.dump(result, sys.stdout, indent=1)
        print()
    s = result["summary"]
    sys.stderr.write(
        "%d decks, %d slides, %.2f GB\n"
        "  %d decks will substitute fonts | %d embed theirs\n"
        "  %d animated slides -> %d extra build pages\n"
        "  %d animated GIFs, %d videos, %d linked | %d slides have notes\n"
        "  %d redundant duplicate copies\n"
        % (s["decks"], s["slides"], s["bytes"] / 1e9,
           s["font_risk_decks"], s["decks_embedding_fonts"],
           s["animated_slides"], s["build_states"],
           s["animated_gifs"], s["videos"], s["external_media"], s["noted_slides"],
           s["redundant_copies"]))

    # Things that will cost time or fidelity later, named now while it is cheap
    # to act on them.
    big = sorted((d for d in result["decks"] if d["bytes"] >= OVERSIZED_BYTES),
                 key=lambda d: -d["bytes"])
    if big:
        sys.stderr.write(
            "\n  %d deck%s over %d MB - large decks can hang PowerPoint; export "
            "these on their own:\n" % (len(big), "" if len(big) == 1 else "s",
                                       OVERSIZED_BYTES // 1000000))
        for d in big[:5]:
            sys.stderr.write("    %6.0f MB  %s\n" % (d["bytes"] / 1e6, d["name"]))
        if len(big) > 5:
            sys.stderr.write("    ... and %d more (see the JSON)\n" % (len(big) - 5))
    bad = [d for d in result["decks"] if d["error"]]
    if bad:
        sys.stderr.write("\n  %d file%s could not be read at all:\n"
                         % (len(bad), "" if len(bad) == 1 else "s"))
        for d in bad[:5]:
            sys.stderr.write("    %s  (%s)\n" % (d["name"], d["error"]))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
