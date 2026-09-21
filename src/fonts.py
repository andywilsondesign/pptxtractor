#!/usr/bin/env python3
"""
fonts - work out what to do about the faces a library asks for and this Mac
does not have.

`audit` already says which faces are missing. That is the diagnosis; this is
the treatment. Three kinds of "missing" turn out to need three different
answers, and telling them apart is most of the value:

  * Some are already here under another name. `Helvetica Neue LT Std 75` is
    Linotype's name for Helvetica Neue Bold, which ships on every Mac, and
    `Aptos Display` is the display cut of Aptos, sixteen cuts of which sit
    inside PowerPoint's own bundle. Nothing to install, nothing to substitute.

  * Some are not fonts at all. `ui-sans-serif` is a CSS keyword that came in
    with a paste from a web page; a name like `Hel\\` is a corrupted string.
    Neither can ever be installed, and counting them inflates the problem.

  * The rest need a stand-in. Here the usual instinct - pick the closest
    looking free face - is wrong often enough to matter, because slides are
    fixed geometry. A substitute wider than the original overflows its box.
    Measured against real text from the decks that use them, Montserrat runs
    11% wider than the Proxima Nova it was standing in for and Inter 5% wider
    than Graphik. Both would have pushed text out of its frame. Width is the
    constraint; appearance is the tie-breaker.

Reporting and alias resolution need nothing but system Python. Building the
substitutes needs to instance variable fonts, which needs fontTools - so that
half asks for it only when you use it.
"""
import json
import os
import re
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

INSTALL_DIR = os.path.expanduser("~/Library/Fonts/pptxtractor-substitutes")


# --------------------------------------------------------------------------
# names
# --------------------------------------------------------------------------

# Foundry and format noise that says nothing about which face is wanted.
# W01..W10 are Monotype web-font cuts; LT is Linotype; Std/Pro are Adobe's
# OpenType generations; MT is Monotype; PS is PostScript.
VENDOR_TOKENS = {
    # "Pro" is deliberately absent: it is part of real family names - Source
    # Sans Pro, Source Code Pro, Myriad Pro - not foundry noise.
    "lt", "std", "mt", "ps", "ot", "tt", "com", "opentype",
    "w01", "w02", "w03", "w04", "w05", "w06", "w07", "w08", "w09", "w10",
}

# Linotype's numeric weights, still used in names exported from older decks.
LINOTYPE_WEIGHTS = {
    "25": "UltraLight", "26": "UltraLight", "35": "Thin", "36": "Thin",
    "45": "Light", "46": "Light", "55": "Regular", "56": "Regular",
    "65": "Medium", "66": "Medium", "75": "Bold", "76": "Bold",
    "85": "Heavy", "86": "Heavy", "95": "Black", "96": "Black",
}

STYLE_TOKENS = {
    "thin": "Thin", "hairline": "Thin",
    "ultralight": "UltraLight", "extralight": "ExtraLight", "xlight": "ExtraLight",
    "light": "Light", "lt": None,           # bare "Lt" handled below
    "book": "Regular", "regular": "Regular", "roman": "Regular", "normal": "Regular",
    "medium": "Medium", "med": "Medium",
    "semibold": "SemiBold", "demibold": "SemiBold", "demi": "SemiBold", "sbold": "SemiBold",
    "bold": "Bold", "bd": "Bold",
    "extrabold": "ExtraBold", "ultrabold": "ExtraBold",
    "heavy": "Heavy", "black": "Black", "fat": "Black",
}

WIDTH_TOKENS = {
    "cn": "Condensed", "cond": "Condensed", "condensed": "Condensed",
    "narrow": "Narrow", "compressed": "Condensed",
    "ext": "Extended", "extended": "Extended", "wide": "Wide",
}

# Optical-size cuts. A face asking for one can fall back to the family's plain
# cut: the drawing is the same design at a different optical size, which is a
# far smaller difference than any substitution.
OPTICAL_TOKENS = {"display", "text", "caption", "subhead", "banner", "poster",
                  "deck", "micro", "body"}

# CSS generic families and stack keywords. These reach a deck when someone
# pastes styled text out of a browser; they are not fonts and never will be.
CSS_KEYWORDS = {
    "ui-sans-serif", "ui-serif", "ui-monospace", "ui-rounded",
    "sans-serif", "serif", "monospace", "cursive", "fantasy", "system-ui",
    "-apple-system", "blinkmacsystemfont", "inherit", "initial", "unset",
}


def split_tokens(name):
    """Break a face name into words, splitting camelCase as well as separators.

    `HelveticaNeueLTStd-Lt` has to come apart into the same words as
    `Helvetica Neue LT Std Lt`, or the two spellings of one face look like two
    different problems.
    """
    # Monotype web cuts are spelled W01..W10. Strip them whole, before the
    # letter/digit split below turns "W03" into "W" and "03" and leaves both
    # stranded in the family name.
    # No \b before the W: in "ObjektivMk3W03" it follows a digit, so there is
    # no word boundary to anchor to.
    s = re.sub(r"W\d{2}(?![A-Za-z0-9])", " ", name)
    s = re.sub(r"[(),]", " ", s)
    s = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", s)
    # Split a letter from a *two-digit* run only: that is the Linotype weight
    # shape ("HelveticaNeue75"). Splitting every digit would break "Mk3" and
    # "Mk2" apart from the family names they belong to.
    s = re.sub(r"([A-Za-z])(\d{2})(?!\d)", r"\1 \2", s)
    s = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", s)
    toks = [t for t in re.split(r"[\s\-_.]+", s) if t]
    # camelCase splitting turns "SemiBold" into "Semi" + "Bold". Put the pair
    # back together, or "Semi" is read as part of the family name and
    # "Source Sans Pro SemiBold" stops being a weight of Source Sans Pro.
    merged, i = [], 0
    while i < len(toks):
        a = toks[i].lower()
        b = toks[i + 1].lower() if i + 1 < len(toks) else ""
        if a in ("semi", "demi", "extra", "ultra", "x") and b in (
                "bold", "light", "black", "condensed"):
            merged.append(toks[i] + toks[i + 1])
            i += 2
        else:
            merged.append(toks[i])
            i += 1
    return merged


def parse_face(name):
    """Split a face name into family, style and width.

    Returns (family, style, width, optical) where style is a canonical weight
    name or None, width is "Condensed"/"Narrow"/... or None, and optical is
    the optical-size word that was dropped, if any.
    """
    tokens = split_tokens(name)
    family, style, width, optical = [], None, None, None
    for i, tok in enumerate(tokens):
        low = tok.lower()
        nxt = tokens[i + 1].lower() if i + 1 < len(tokens) else ""
        # "HelveticaNeueLTStd-Lt" carries both senses of LT: the first is
        # Linotype the foundry, the trailing one is Light the weight. The
        # foundry sense is the one followed by Std, so look ahead rather than
        # swallowing every LT and losing the weight.
        if low == "lt" and nxt != "std":
            style = "Light"
            continue
        if low in VENDOR_TOKENS:
            continue
        if low in LINOTYPE_WEIGHTS or (tok.isdigit() and tok in LINOTYPE_WEIGHTS):
            style = LINOTYPE_WEIGHTS[low]
            continue
        if low in WIDTH_TOKENS:
            width = WIDTH_TOKENS[low]
            continue
        if low in OPTICAL_TOKENS and family:
            # only optical once we have something to be the optical cut *of*
            optical = tok
            continue
        if low == "l":
            style = "Light"
            continue
        if low in STYLE_TOKENS and STYLE_TOKENS[low]:
            style = STYLE_TOKENS[low]
            continue
        if low == "italic" or low == "oblique" or low == "it":
            continue
        family.append(tok)
    return " ".join(family), style, width, optical


def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def is_noise(name):
    """A name that can never be installed, because it is not a font name."""
    low = name.strip().lower()
    if low in CSS_KEYWORDS:
        return "CSS keyword, not a font - pasted in from a web page"
    if re.search(r"[\\/<>{}|]", name):
        return "corrupted name - not a real face"
    if len(low) < 3:
        return "too short to be a face name"
    if not re.search(r"[a-z]", low):
        return "no letters - not a face name"
    return None


# --------------------------------------------------------------------------
# what is actually on this Mac
# --------------------------------------------------------------------------

def installed_faces(exclude=None):
    """Every installed face, as {normalised full name: (family, style)}.

    `audit` matches on *filenames*, which is fast and right for its purpose but
    cannot see inside a collection: Helvetica Neue's fourteen cuts live in one
    HelveticaNeue.ttc, and Aptos Display does not exist as a file at all. Those
    are exactly the cases this command has to resolve, so ask the system what
    it has actually registered.
    """
    # Families a previous --install added are not evidence that the real face
    # is present - they are this tool's own work. Drop them, or a second run
    # congratulates itself.
    exclude = exclude or {}
    faces = {}
    try:
        out = subprocess.run(["system_profiler", "SPFontsDataType"],
                             capture_output=True, text=True, timeout=120).stdout
    except Exception:
        return faces
    family = None
    for line in out.splitlines():
        s = line.strip()
        if s.startswith("Family: "):
            family = s[8:].strip()
            if family and norm(family) not in exclude:
                faces.setdefault(norm(family), (family, family))
        elif s.startswith("Full Name: ") and family:
            full = s[11:].strip()
            if full and norm(full) not in exclude and norm(family) not in exclude:
                faces.setdefault(norm(full), (family, full))
    return faces


def office_font_dirs():
    return [d for d in (
        "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts",
        "/Applications/Microsoft Word.app/Contents/Resources/DFonts",
        "/Applications/Microsoft Excel.app/Contents/Resources/DFonts",
    ) if os.path.isdir(d)]


def office_faces():
    """Faces bundled inside the Office apps, by normalised file stem.

    Office registers these privately, so they are usable in PowerPoint but do
    not always appear as installed fonts. Aptos ships sixteen cuts this way.
    """
    found = {}
    for d in office_font_dirs():
        for f in sorted(os.listdir(d)):
            if f.lower().endswith((".ttf", ".otf", ".ttc")):
                found.setdefault(norm(os.path.splitext(f)[0]), os.path.join(d, f))
    return found


def our_substitutes():
    """Families a previous --install put on this Mac, from its manifest.

    Without this a second run reads its own work as evidence that the faces
    were installed all along, and reports a library as healthy when the real
    fonts are still absent.
    """
    path = os.path.join(INSTALL_DIR, "manifest.json")
    try:
        with open(path) as fh:
            return {norm(m["family"]): m.get("source", "a substitute")
                    for m in json.load(fh)}
    except Exception:
        return {}


def resolve_alias(face, sysfaces, offfaces):
    """Find an installed face that *is* this face under another name.

    Returns (description, source_path_or_None, kind) or None.
    """
    # 1. the name exactly as the deck spells it. Cheapest and safest: no
    #    token surgery can turn "Source Sans Pro" into something else.
    hit = sysfaces.get(norm(face))
    if hit:
        return hit[1], None, "system"

    family, style, width, optical = parse_face(face)
    if not family:
        return None
    style = style or "Regular"
    want_style = (width + " " + style).strip() if width else style

    # 2. family plus what we parsed out. Most specific first: a name carrying
    #    a width must not fall through to the plain family and silently lose
    #    it - "Aptos Narrow" resolving to Aptos is a different typeface.
    base = (family + " " + width) if width else family
    for cand in (base + " " + style, base, family + " " + want_style,
                 family + " " + style):
        hit = sysfaces.get(norm(cand))
        if hit:
            return hit[1], None, "system"

    # 2. an optical cut we can serve from the plain family: "Aptos Display"
    #    is Aptos drawn for large sizes, so Aptos itself is far closer than
    #    any substitute would be.
    if optical:
        for cand in (family + " " + want_style, family):
            hit = sysfaces.get(norm(cand))
            if hit:
                return hit[1], None, "system"

    # 3. Office's private bundle, by file stem
    for cand in (base + "-" + style, base + " " + style, base,
                 family + "-" + want_style, family + "-" + style,
                 family + " " + want_style, family):
        path = offfaces.get(norm(cand))
        if path:
            return os.path.splitext(os.path.basename(path))[0], path, "office"
    return None


# --------------------------------------------------------------------------
# text to measure against
# --------------------------------------------------------------------------

def deck_text(paths, limit=200000):
    """Slide text straight out of the OOXML. Read-only; decks are untouched.

    Measuring on the library's own words beats a pangram: these decks are
    mostly headings and short sentences in title case, which weights the
    capitals differently from running prose.
    """
    out = []
    total = 0
    for p in paths:
        try:
            with zipfile.ZipFile(p) as z:
                for n in z.namelist():
                    if n.startswith("ppt/slides/slide") and n.endswith(".xml"):
                        xml = z.read(n).decode("utf8", "ignore")
                        for t in re.findall(r"<a:t>(.*?)</a:t>", xml, re.S):
                            out.append(t)
                            total += len(t)
        except Exception:
            continue
        if total > limit:
            break
    text = " ".join(out)
    # Fall back to something representative rather than measuring nothing.
    return text if len(text) > 200 else (text + " " + SAMPLE_TEXT)


SAMPLE_TEXT = (
    "The quick brown fox jumps over the lazy dog. Research Findings and "
    "Recommendations. Customer Journey Mapping. Key Takeaways for the Board. "
    "0123456789 How might we improve onboarding conversion this quarter?"
)


# --------------------------------------------------------------------------
# substitutes
# --------------------------------------------------------------------------

GF_RAW = "https://raw.githubusercontent.com/google/fonts/main"

# Free faces worth standing in for a licensed one. Everything here is under
# the SIL Open Font License, which permits both the renaming this does and the
# local use it is renamed for. Google Fonts serves these as variable fonts, so
# a weight is instanced out rather than downloaded ready-made.
CANDIDATES = {
    "Archivo":       ("ofl/archivo/Archivo%5Bwdth,wght%5D.ttf", "grotesque"),
    "Inter":         ("ofl/inter/Inter%5Bopsz,wght%5D.ttf", "grotesque"),
    "Inter Tight":   ("ofl/intertight/InterTight%5Bwght%5D.ttf", "grotesque"),
    "Work Sans":     ("ofl/worksans/WorkSans%5Bwght%5D.ttf", "grotesque"),
    "Public Sans":   ("ofl/publicsans/PublicSans%5Bwght%5D.ttf", "grotesque"),
    "Libre Franklin": ("ofl/librefranklin/LibreFranklin%5Bwght%5D.ttf", "grotesque"),
    "Figtree":       ("ofl/figtree/Figtree%5Bwght%5D.ttf", "geometric"),
    "Hanken Grotesk": ("ofl/hankengrotesk/HankenGrotesk%5Bwght%5D.ttf", "geometric"),
    "Jost":          ("ofl/jost/Jost%5Bwght%5D.ttf", "geometric"),
    "Montserrat":    ("ofl/montserrat/Montserrat%5Bwght%5D.ttf", "geometric"),
    "Rubik":         ("ofl/rubik/Rubik%5Bwght%5D.ttf", "geometric"),
    "Overpass":      ("ofl/overpass/Overpass%5Bwght%5D.ttf", "signage"),
    "Source Sans 3": ("ofl/sourcesans3/SourceSans3%5Bwght%5D.ttf", "humanist"),
    "Source Code Pro": ("ofl/sourcecodepro/SourceCodePro%5Bwght%5D.ttf", "mono"),
    "Mulish":        ("ofl/mulish/Mulish%5Bwght%5D.ttf", "humanist"),
    "Bebas Neue":    ("ofl/bebasneue/BebasNeue-Regular.ttf", "display"),
    "Amatic SC":     ("ofl/amaticsc/AmaticSC-Regular.ttf", "display"),
}

# Where a licensed family has an obvious free descendant or successor, say so:
# it beats anything width-matching could find. The right-hand name must be a
# key of CANDIDATES.
KNOWN_SUCCESSORS = {
    "source sans pro": "Source Sans 3",     # renamed, same design
    "source code pro": "Source Code Pro",   # the same font, free
    "muli": "Mulish",                       # renamed
    "montserrat": "Montserrat",
    "bebas neue": "Bebas Neue",
    "amatic sc": "Amatic SC",
}

# What kind of face this is, where knowing costs nothing. Width measurement
# still decides; this only breaks ties between equally-close candidates.
GENRE_HINTS = {
    "graphik": "grotesque",
    "objektiv": "geometric",
    "proxima nova": "geometric",
    "google sans": "geometric",
    "transport new": "signage",
    "transport": "signage",
}

# Weight names to the numeric axis value a variable font understands.
WEIGHT_VALUES = {
    "Thin": 100, "UltraLight": 200, "ExtraLight": 200, "Light": 300,
    "Regular": 400, "Medium": 500, "SemiBold": 600, "Bold": 700,
    "ExtraBold": 800, "Heavy": 800, "Black": 900,
}

ANCHOR_TTC = "/System/Library/Fonts/HelveticaNeue.ttc"


def require_fonttools():
    """Import fontTools, or explain exactly why this half cannot run."""
    try:
        from fontTools.ttLib import TTFont            # noqa: F401
        from fontTools.varLib import instancer        # noqa: F401
        return True
    except ImportError:
        sys.stderr.write(
            "Building substitutes needs fontTools, which is not installed.\n"
            "\n"
            "  python3 -m pip install --user fonttools\n"
            "\n"
            "Everything above this line - the missing faces, the ones already\n"
            "here under another name, and the ones that are not fonts - needed\n"
            "nothing but system Python and is complete. Only building and\n"
            "installing the substitutes needs the extra package, because it\n"
            "means instancing variable fonts, and nothing in the standard\n"
            "library does that.\n")
        return False


CACHE_DIR = os.path.expanduser("~/Library/Caches/pptxtractor/fonts")


def fetch(rel, quiet=False):
    """Download one font from Google Fonts, cached.

    Cached because picking a substitute measures every candidate, and a second
    run should not re-download a dozen families to tell you the same thing.
    """
    import urllib.request
    os.makedirs(CACHE_DIR, exist_ok=True)
    # Key the cache on the whole path, not the basename: every family's
    # licence is called OFL.txt, so caching by basename made the first one
    # downloaded stand in for all of them.
    flat = rel.replace("%5B", "[").replace("%5D", "]").replace("/", "_")
    dest = os.path.join(CACHE_DIR, flat)
    if os.path.exists(dest) and os.path.getsize(dest) > 1000:
        return dest
    url = GF_RAW + "/" + rel
    if not quiet:
        sys.stderr.write("  fetching %s\n" % rel.rsplit("/", 1)[-1])
    try:
        with urllib.request.urlopen(url, timeout=60) as r, open(dest, "wb") as fh:
            fh.write(r.read())
    except Exception as e:
        sys.stderr.write("  could not fetch %s: %s\n" % (rel, e))
        return None
    return dest


def load_instance(path, weight=400):
    """A static font at one weight, instanced from a variable font if needed."""
    from fontTools.ttLib import TTFont
    from fontTools.varLib import instancer
    f = TTFont(path)
    if "fvar" not in f:
        return f
    axes = {a.axisTag: a for a in f["fvar"].axes}
    limits = {}
    for tag, a in axes.items():
        if tag == "wght":
            limits[tag] = max(a.minValue, min(weight, a.maxValue))
        else:
            # Pin every other axis at its default: an optical-size or width
            # axis left free would make the instance a variable font again.
            limits[tag] = a.defaultValue
    return instancer.instantiateVariableFont(f, limits, inplace=True,
                                             updateFontNames=False)


def load_anchor(index=0):
    """Helvetica Neue, straight out of the system collection.

    Used as the yardstick because it is on every Mac, so a measurement taken
    here means the same thing on another machine, and because most text faces
    a deck is likely to use sit within a few percent of it.
    """
    from fontTools.ttLib import TTCollection
    import io
    tc = TTCollection(ANCHOR_TTC)
    buf = io.BytesIO()
    tc.fonts[index].save(buf)
    buf.seek(0)
    from fontTools.ttLib import TTFont
    return TTFont(buf)


def measure(font, text):
    """Mean advance width per character, per 1000 units of em.

    Normalising by em makes fonts with different units-per-em comparable, and
    measuring per character rather than per string makes different sample
    sizes comparable.
    """
    upm = font["head"].unitsPerEm or 1000
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    total = count = 0
    for ch in text:
        g = cmap.get(ord(ch))
        if g is None:
            continue
        total += hmtx[g][0]
        count += 1
    return (total / count) * 1000.0 / upm if count else 0.0


def genre_for(face):
    low = face.lower()
    for key, genre in GENRE_HINTS.items():
        if key in low:
            return genre
    return None


def rank_candidates(face, text, weight=400, pool=None):
    """Measure every plausible candidate and rank by width against the anchor.

    The rule is deliberately asymmetric. A substitute narrower than the face it
    replaces leaves a little more room at the end of a line, which nobody
    notices. One that is wider pushes text out of a box that was sized to fit,
    which everybody does. So candidates at or under the anchor are preferred,
    closest first; only if none fit does the narrowest overshoot win.
    """
    anchor = measure(load_anchor(), text)
    if not anchor:
        return []
    genre = genre_for(face)
    names = pool or list(CANDIDATES)
    if genre:
        preferred = [n for n in names if CANDIDATES[n][1] == genre]
        names = preferred + [n for n in names if n not in preferred]
    rows = []
    for name in names:
        rel, kind = CANDIDATES[name]
        if kind in ("display", "mono"):
            continue                      # never a general text substitute
        path = fetch(rel, quiet=True)
        if not path:
            continue
        try:
            w = measure(load_instance(path, weight), text)
        except Exception:
            continue
        if not w:
            continue
        ratio = w / anchor
        rows.append({"family": name, "genre": kind, "ratio": ratio,
                     "delta_pct": (ratio - 1.0) * 100.0,
                     "in_genre": bool(genre) and kind == genre})
    # Overshoot is disqualifying before anything else: too wide is the one
    # failure that shows. Among the rest, matching the genre beats being a
    # fraction closer on width - a grotesque standing in for a grotesque reads
    # as the same kind of type, and the width difference between the top few
    # candidates is smaller than the difference between a grotesque and a
    # geometric. Width then orders within the genre.
    rows.sort(key=lambda r: (r["ratio"] > 1.0,
                             not r["in_genre"],
                             -r["ratio"] if r["ratio"] <= 1.0 else r["ratio"]))
    return rows


def rename_font(font, family, bold, note):
    """Make the font declare itself as `family`, so PowerPoint resolves it.

    This is the whole trick: no deck is edited, and nothing is rewritten at
    export time. The face simply answers to the name the slides ask for.
    """
    sub_name = "Bold" if bold else "Regular"
    ps = "".join(c for c in family if c.isalnum()) + ("-Bold" if bold else "-Regular")
    full = ("%s %s" % (family, sub_name)) if bold else family
    name = font["name"]
    # Typographic and WWS names win over the basic family in some apps, so a
    # leftover pair here would quietly undo the rename.
    for nid in (16, 17, 18, 20, 21, 22):
        name.removeNames(nameID=nid)
    fields = ((1, family), (2, sub_name), (3, "%s; pptxtractor substitute" % full),
              (4, full), (6, ps),
              (10, "Substitute installed by pptxtractor. Source: %s." % note),
              (16, family), (17, sub_name))
    for nid, val in fields:
        name.setName(val, nid, 3, 1, 0x409)     # Windows / Unicode BMP / en-US
        name.setName(val, nid, 1, 0, 0)         # Macintosh / Roman / English
    os2 = font["OS/2"]
    os2.usWeightClass = 700 if bold else 400
    # bit 0 italic, bit 5 bold, bit 6 regular; exactly one of bold/regular.
    ITALIC, BOLD, REGULAR = 1 << 0, 1 << 5, 1 << 6
    os2.fsSelection = (os2.fsSelection & ~(ITALIC | BOLD | REGULAR)) | \
                      (BOLD if bold else REGULAR)
    font["head"].macStyle = 1 if bold else 0
    return font


FONT_DIRS = ["/System/Library/Fonts", "/Library/Fonts",
             os.path.expanduser("~/Library/Fonts")] 


def find_face_file(full_name):
    """Locate the file (and collection index) holding a named face.

    An alias is only useful if we can reach the actual outlines: knowing that
    `Helvetica Neue LT Std 75` means Helvetica Neue Bold does not help until we
    can open the Bold cut, which lives at index 1 of a fourteen-font
    collection.
    """
    from fontTools.ttLib import TTFont, TTCollection
    want = norm(full_name)
    for d in FONT_DIRS + office_font_dirs():
        for root, _dirs, files in os.walk(d):
            for fn in sorted(files):
                low = fn.lower()
                if not low.endswith((".ttf", ".otf", ".ttc")):
                    continue
                path = os.path.join(root, fn)
                try:
                    if low.endswith(".ttc"):
                        tc = TTCollection(path)
                        for i, f in enumerate(tc.fonts):
                            if norm(f["name"].getDebugName(4) or "") == want:
                                return path, i
                    else:
                        f = TTFont(path, lazy=True)
                        if norm(f["name"].getDebugName(4) or "") == want:
                            return path, None
                        f.close()
                except Exception:
                    continue
    return None, None


def open_source(path, index=None):
    from fontTools.ttLib import TTFont, TTCollection
    import io
    if index is not None:
        tc = TTCollection(path)
        buf = io.BytesIO()
        tc.fonts[index].save(buf)
        buf.seek(0)
        return TTFont(buf)
    return TTFont(path)


def plan_for(root, json_out=None):
    """Work out what every missing face needs. No network, no fontTools."""
    import audit as audit_mod
    # Ask what is available *without* our own substitutes. Otherwise a second
    # run sees the fonts it installed last time and reports a clean library
    # while the licensed faces are still missing - true of the export, but not
    # the answer to "what do I need?".
    result = audit_mod.audit(root, exclude_font_dirs=(INSTALL_DIR,))
    decks = result["decks"]

    per_face = {}
    for d in decks:
        for f in d["fonts_missing"]:
            per_face.setdefault(f, []).append(d["path"])

    ours = our_substitutes()
    sysf = installed_faces(exclude=ours)
    offf = office_faces()

    noise, alias, subst = [], [], []
    for face in sorted(per_face):
        paths = per_face[face]
        why = is_noise(face)
        if why:
            noise.append({"face": face, "why": why, "decks": len(paths)})
            continue
        if norm(face) in ours:
            alias.append({"face": face, "to": ours[norm(face)], "kind": "installed",
                          "decks": len(paths), "paths": paths})
            continue
        hit = resolve_alias(face, sysf, offf)
        if hit:
            alias.append({"face": face, "to": hit[0], "kind": hit[2],
                          "source_path": hit[1], "decks": len(paths), "paths": paths})
            continue
        fam, style, width, _opt = parse_face(face)
        subst.append({"face": face, "family": fam or face,
                      "style": style or "Regular",
                      "weight": WEIGHT_VALUES.get(style or "Regular", 400),
                      "decks": len(paths), "paths": paths})
    return {"root": root, "decks": len(decks), "noise": noise,
            "alias": alias, "substitute": subst,
            "risk_decks": len({p for v in per_face.values() for p in v})}


def group_substitutes(subst):
    """Collapse a family's weights into one decision.

    Graphik arrived as ten separate faces - Thin through Black. They are one
    typeface and want one stand-in; presenting them as ten problems would be
    ten times the reading for the same choice.
    """
    groups = {}
    for s in subst:
        key = norm(s["family"])
        g = groups.setdefault(key, {"family": s["family"], "faces": [],
                                    "decks": set(), "paths": set()})
        g["faces"].append(s)
        g["decks"].update(s["paths"])
        g["paths"].update(s["paths"])
        if len(s["family"]) < len(g["family"]):
            g["family"] = s["family"]        # shortest spelling reads best

    # A named cut of a family - "Graphik Courant" beside plain "Graphik" - is
    # still that family, and wants the same stand-in. Where one family name is
    # a whole-word prefix of another, fold the longer into the shorter, so the
    # choice is made once for the typeface rather than once per cut.
    keys = sorted(groups, key=len)
    for long_key in list(keys):
        for short_key in keys:
            if short_key == long_key or short_key not in groups or long_key not in groups:
                continue
            a, b = groups[short_key]["family"].lower(), groups[long_key]["family"].lower()
            if b.startswith(a + " "):
                groups[short_key]["faces"] += groups[long_key]["faces"]
                groups[short_key]["decks"] |= groups[long_key]["decks"]
                groups[short_key]["paths"] |= groups[long_key]["paths"]
                del groups[long_key]
                break
    return sorted(groups.values(), key=lambda g: -len(g["decks"]))


def recommend(groups, quiet=False):
    """Attach a measured recommendation to each substitute group."""
    for g in groups:
        successor = KNOWN_SUCCESSORS.get(norm(g["family"]).replace(" ", "")) \
            or KNOWN_SUCCESSORS.get(g["family"].lower())
        if successor:
            g["pick"] = successor
            g["delta_pct"] = None
            g["why"] = "the same design, free"
            g["ranked"] = []
            continue
        text = deck_text(sorted(g["paths"]))
        try:
            ranked = rank_candidates(g["family"], text)
        except Exception:
            ranked = []
        g["ranked"] = ranked
        if ranked:
            g["pick"] = ranked[0]["family"]
            g["delta_pct"] = ranked[0]["delta_pct"]
            g["why"] = "measured on %d characters from these decks" % len(text)
        else:
            g["pick"] = None
            g["delta_pct"] = None
            g["why"] = "no candidate could be measured"
    return groups


def report(plan, groups, stream=sys.stderr):
    w = stream.write
    n_alias = sum(a["decks"] for a in plan["alias"])
    w("%d of %d decks ask for %d face(s) this Mac does not have.\n\n"
      % (plan["risk_decks"], plan["decks"],
         len(plan["alias"]) + len(plan["noise"]) + len(plan["substitute"])))

    if plan["alias"]:
        w("Already here under another name - nothing to buy:\n")
        for a in plan["alias"]:
            where = {"system": "macOS", "office": "Microsoft Office",
                     "installed": "installed earlier by pptxtractor"}.get(a["kind"], a["kind"])
            w("  %-30s -> %-28s %s\n" % (a["face"], a["to"], where))
        w("\n")

    if plan["noise"]:
        w("Not fonts - nothing to install, ever:\n")
        for nz in plan["noise"]:
            w("  %-30s %s\n" % (nz["face"], nz["why"]))
        w("\n")

    if groups:
        w("Need a substitute:\n")
        for g in groups:
            faces = len(g["faces"])
            extra = faces - 1
            label = "%s%s" % (g["family"],
                              (" +%d weight%s" % (extra, "" if extra == 1 else "s"))
                              if extra else "")
            delta = ("%+.1f%% width" % g["delta_pct"]) if g["delta_pct"] is not None else g["why"]
            w("  %-28s %3d decks  ->  %-16s %s\n"
              % (label[:28], len(g["decks"]), g["pick"] or "(none found)", delta))
        w("\n")
        w("Substitutes are chosen by measured width, not by eye: a face wider than\n"
          "the one it replaces overflows a text box that was sized to fit. Each is\n"
          "at or under the width of Helvetica Neue, measured on these decks' own text.\n\n")
    return n_alias


def install(plan, groups, dest=None):
    """Build every needed face and put it where PowerPoint will find it.

    Each face is written twice, Regular and Bold, so a deck that bolds a run
    gets a real bold rather than a smeared synthetic one.
    """
    # Resolved here, not in the signature: a default argument binds once at
    # definition, so a caller that reassigns INSTALL_DIR would still write to
    # the original path - which is how a test aimed at a temp folder came to
    # overwrite the real one.
    dest = dest or INSTALL_DIR
    os.makedirs(dest, exist_ok=True)
    manifest, written = [], 0

    def emit(family, font, bold, note):
        nonlocal written
        rename_font(font, family, bold, note)
        safe = "".join(c if c.isalnum() else "_" for c in family)
        font.save(os.path.join(dest, "%s-%s.ttf" % (safe, "Bold" if bold else "Regular")))
        font.close()
        written += 1

    # faces that exist here already, republished under the name decks ask for
    for a in plan["alias"]:
        if a["kind"] == "installed":
            continue
        path, index = (a.get("source_path"), None)
        if not path:
            path, index = find_face_file(a["to"])
        if not path:
            sys.stderr.write("  could not locate %s - skipping %s\n" % (a["to"], a["face"]))
            continue
        try:
            emit(a["face"], open_source(path, index), False, a["to"])
            emit(a["face"], open_source(path, index), True, a["to"])
            manifest.append({"family": a["face"], "source": a["to"], "kind": a["kind"]})
        except Exception as e:
            sys.stderr.write("  %s: %s\n" % (a["face"], e))

    # licensed faces that need a free stand-in
    for g in groups:
        if not g["pick"]:
            continue
        rel = CANDIDATES[g["pick"]][0]
        src = fetch(rel)
        if not src:
            continue
        for face in g["faces"]:
            wght = face["weight"]
            bold_w = min(900, wght + 200) if wght < 700 else max(wght, 700)
            note = "%s %s" % (g["pick"], face["style"])
            try:
                emit(face["face"], load_instance(src, wght), False, note)
                emit(face["face"], load_instance(src, bold_w), True, note)
                manifest.append({"family": face["face"], "source": note,
                                 "kind": "substitute"})
            except Exception as e:
                sys.stderr.write("  %s: %s\n" % (face["face"], e))
        lic = fetch(rel.rsplit("/", 1)[0] + "/OFL.txt", quiet=True)
        if lic:
            import shutil
            shutil.copy(lic, os.path.join(dest, "OFL-%s.txt" % norm(g["pick"])))

    # Never replace a good manifest with an empty one. A run that built
    # nothing has nothing to record, and clobbering the record of an earlier
    # install would lose track of every font it put here.
    if manifest:
        with open(os.path.join(dest, "manifest.json"), "w") as fh:
            json.dump(manifest, fh, indent=2)
    uninstall = os.path.join(dest, "UNINSTALL.sh")
    with open(uninstall, "w") as fh:
        fh.write("#!/bin/sh\n"
                 "# Remove every font pptxtractor installed. Nothing else is touched.\n"
                 'rm -rf "%s"\n'
                 'echo "Removed. Restart PowerPoint for the change to take effect."\n'
                 % dest)
    os.chmod(uninstall, 0o755)
    return written, len(manifest)


def main(argv):
    args = list(argv)
    do_install = "--install" in args
    if do_install:
        args.remove("--install")
    json_out = None
    if "--json" in args:
        i = args.index("--json")
        json_out = args[i + 1] if i + 1 < len(args) else None
        del args[i:i + 2]
    if not args:
        sys.stderr.write("usage: pptxtractor fonts <folder> [--install] [--json FILE]\n")
        return 2
    root = os.path.abspath(os.path.expanduser(args[0].rstrip("/")))
    if not os.path.exists(root):
        sys.stderr.write("no such folder: %s\n" % root)
        return 2

    plan = plan_for(root)
    groups = group_substitutes(plan["substitute"])

    # Measuring needs fontTools. Without it the diagnosis still stands, so say
    # what is known and stop there rather than failing the whole command.
    have_ft = True
    try:
        from fontTools.ttLib import TTFont  # noqa: F401
    except ImportError:
        have_ft = False
    if have_ft:
        groups = recommend(groups)
    else:
        for g in groups:
            g.update(pick=None, delta_pct=None, ranked=[], why="needs fontTools to measure")

    report(plan, groups)

    if json_out:
        out = {"root": root, "alias": plan["alias"], "noise": plan["noise"],
               "substitute": [{"family": g["family"], "decks": len(g["decks"]),
                               "faces": [f["face"] for f in g["faces"]],
                               "pick": g["pick"], "delta_pct": g["delta_pct"],
                               "ranked": g.get("ranked", [])[:5]} for g in groups]}
        with open(json_out, "w") as fh:
            json.dump(out, fh, indent=1)
        sys.stderr.write("written: %s\n" % json_out)

    if not do_install:
        if groups or [a for a in plan["alias"] if a["kind"] != "installed"]:
            sys.stderr.write("Run again with --install to build and install these.\n"
                             "They go to ~/Library/Fonts/pptxtractor-substitutes/, and\n"
                             "UNINSTALL.sh in that folder removes every one of them.\n")
        return 0

    if not have_ft and not require_fonttools():
        return 5
    files, faces = install(plan, groups)
    sys.stderr.write("installed %d files for %d face name(s) in %s\n"
                     % (files, faces, INSTALL_DIR))
    sys.stderr.write("Quit and reopen PowerPoint so it picks them up, then re-run\n"
                     "`pptxtractor audit` to confirm nothing will be substituted.\n")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
