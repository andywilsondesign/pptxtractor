#!/usr/bin/env python3
"""
Pull speaker notes out of a .pptx and map them to PDF page numbers.

    python3 extract-notes.py <deck.pptx> <notes.json> [slides.json]

`slides.json` is the page map written when animation builds were expanded; pass
it so notes land on every page belonging to their slide. Without it, page number
equals slide number.
"""
import html, json, os, re, sys, zipfile

SHAPE = re.compile(r'<p:sp\b.*?</p:sp>', re.S)


def notes_for(xml):
    """Text of the body placeholder only - not the slide-number placeholder."""
    for sp in SHAPE.findall(xml):
        if not re.search(r'<p:ph\b[^>]*type="body"', sp):
            continue
        paras = []
        for p in re.findall(r'<a:p\b.*?</a:p>', sp, re.S):
            runs = [html.unescape(t) for t in re.findall(r'<a:t>(.*?)</a:t>', p, re.S)]
            line = "".join(runs).strip()
            if line:
                paras.append(line)
        text = "\n".join(paras).strip()
        if text and not re.fullmatch(r'[\d\s]*', text):
            return text
    return ""


def slide_notes(src):
    """slide number -> notes text."""
    out = {}
    with zipfile.ZipFile(src) as z:
        names = set(z.namelist())
        for n in names:
            m = re.fullmatch(r'ppt/slides/_rels/slide(\d+)\.xml\.rels', n)
            if not m:
                continue
            idx = int(m.group(1))
            rels = z.read(n).decode('utf8', 'ignore')
            t = re.search(r'Target="\.\./(notesSlides/notesSlide\d+\.xml)"', rels)
            if not t:
                continue
            part = 'ppt/' + t.group(1)
            if part not in names:
                continue
            text = notes_for(z.read(part).decode('utf8', 'ignore'))
            if text:
                out[idx] = text
    return out


def to_pages(notes, slides_json):
    """Map slide-keyed notes onto page numbers."""
    if not slides_json or not os.path.exists(slides_json):
        return {str(k): v for k, v in sorted(notes.items())}
    pages = json.load(open(slides_json)).get("pages", [])
    out = {}
    for p in pages:
        text = notes.get(p["slide"])
        if not text:
            continue
        if p.get("state") is not None:
            text = "[build %d of %d]\n%s" % (p["state"], p["of"] - 1, text)
        out[str(p["page"])] = text
    return out


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    smap = sys.argv[3] if len(sys.argv) > 3 else None
    notes = slide_notes(src)
    pages = to_pages(notes, smap)
    if pages:
        json.dump(pages, open(dst, "w"), indent=1)
        chars = sum(len(v) for v in pages.values())
        print("  notes: %d slide%s, %d characters"
              % (len(notes), "" if len(notes) == 1 else "s", chars))
