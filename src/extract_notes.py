#!/usr/bin/env python3
"""
Pull speaker notes out of a .pptx and map them to PDF page numbers.

    python3 extract-notes.py <deck.pptx> <notes.json> [slides.json]

`slides.json` is the page map written when animation builds were expanded; pass
it so notes land on every page belonging to their slide. Without it there is one
page per slide - but hidden slides are skipped either way, because PowerPoint
leaves them out of the PDF.
"""
import html, json, os, re, sys, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import buildstates          # shared slide ordering and hidden-slide test

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
    """slide number -> notes text, keyed by *presentation position*.

    Keyed by position in sldIdLst rather than by the slideN.xml file name:
    the page map is built in presentation order, and a deck whose file
    numbering has drifted from its running order would otherwise pair notes
    with the wrong slide.
    """
    out, page_of, page = {}, {}, 0
    with zipfile.ZipFile(src) as z:
        names = set(z.namelist())
        blobs = {n: z.read(n) for n in names
                 if n in ('ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels')
                 or re.fullmatch(r'ppt/slides/slide\d+\.xml', n)}
        order = buildstates.slide_order(blobs)
        for pos, (part, _tag) in enumerate(order, 1):
            if buildstates.is_hidden(blobs[part]):
                continue                     # not exported, so it occupies no page
            page += 1
            page_of[pos] = page
            rels_part = part.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
            if rels_part not in names:
                continue
            rels = z.read(rels_part).decode('utf8', 'ignore')
            t = re.search(r'Target="\.\./(notesSlides/notesSlide\d+\.xml)"', rels)
            if not t:
                continue
            npart = 'ppt/' + t.group(1)
            if npart not in names:
                continue
            text = notes_for(z.read(npart).decode('utf8', 'ignore'))
            if text:
                out[pos] = text
    return out, page_of


def to_pages(notes, slides_json, page_of=None):
    """Map slide-keyed notes onto page numbers."""
    if not slides_json or not os.path.exists(slides_json):
        # No build expansion: one page per slide - but PowerPoint drops hidden
        # slides from the PDF, so page number only equals slide number until
        # the first hidden slide. page_of already accounts for that.
        if page_of is None:
            return {str(k): v for k, v in sorted(notes.items())}
        return {str(page_of[s]): t for s, t in sorted(notes.items()) if s in page_of}
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
    notes, page_of = slide_notes(src)
    pages = to_pages(notes, smap, page_of)
    if pages:
        json.dump(pages, open(dst, "w"), indent=1)
        chars = sum(len(v) for v in pages.values())
        print("  notes: %d slide%s, %d characters"
              % (len(notes), "" if len(notes) == 1 else "s", chars))
