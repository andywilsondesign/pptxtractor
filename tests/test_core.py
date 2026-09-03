#!/usr/bin/env python3
"""
Tests for everything that does not need PowerPoint.

The PowerPoint-driven export cannot be covered here - CI runners have no Office
licence - so this exercises the parts that do the actual thinking: the audit,
the animation expansion, notes extraction and the PDF incremental update.
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

import make_fixture                                  # noqa: E402
import audit as audit_mod                            # noqa: E402
import buildstates                                   # noqa: E402
import extract_notes                                 # noqa: E402
import extract_media                                 # noqa: E402
import add_notes                                     # noqa: E402
import name_images                               # noqa: E402

FAILURES = []


def check(label, got, want):
    ok = got == want
    print(("  ok   " if ok else "  FAIL ") + label + ("" if ok else "  got %r want %r" % (got, want)))
    if not ok:
        FAILURES.append(label)


def truthy(label, got):
    ok = bool(got)
    print(("  ok   " if ok else "  FAIL ") + label + ("" if ok else "  got %r" % (got,)))
    if not ok:
        FAILURES.append(label)


def main():
    tmp = tempfile.mkdtemp(prefix="pptxtractor-test-")
    try:
        deck = make_fixture.build_pptx(os.path.join(tmp, "fixture.pptx"))
        pdf = make_fixture.build_pdf(os.path.join(tmp, "fixture.pdf"), pages=4)

        print("audit")
        result = audit_mod.audit(tmp)
        d = result["decks"][0]
        check("finds the deck", result["summary"]["decks"], 1)
        check("counts slides", d["slides"], 2)
        check("reads slide size", (d["width_px"], d["height_px"]), (1280, 720))
        check("labels aspect", d["aspect"], "16:9")
        check("counts animated slides", d["animated_slides"], 1)
        check("counts build states", d["build_states"], 2)
        check("counts animated gifs", d["gifs"], 1)
        check("counts noted slides", d["noted_slides"], 1)
        check("records external media", len(d["external_media"]), 1)
        truthy("flags the missing font", "NotARealFontFace" in d["fonts_missing"])
        truthy("does not flag a real font", "Helvetica" not in d["fonts_missing"])
        check("no read errors", d["error"], None)

        print("animation expansion")
        steps = buildstates.click_steps(zipfile.ZipFile(deck).read("ppt/slides/slide1.xml"))
        check("two click steps", len(steps), 2)
        out = os.path.join(tmp, "expanded.pptx")
        expanded = buildstates.expand_deck(deck, out)
        check("one slide expanded", len(expanded), 1)
        check("into three pages", expanded[0][1], 3)
        with zipfile.ZipFile(out) as z:
            slides = [n for n in z.namelist() if n.startswith("ppt/slides/slide")
                      and n.endswith(".xml")]
            check("slide parts added", len(slides), 4)
            pres = z.read("ppt/presentation.xml").decode()
            import re
            ids = re.findall(r'<p:sldId\b[^>]*id="(\d+)"', pres)
            check("sldIdLst grew", len(ids), 4)
            check("slide ids unique", len(set(ids)), 4)
            s0 = z.read("ppt/slides/slide1.xml").decode()
            truthy("state 0 drops later shapes", 'name="One"' not in s0 and 'name="Title"' in s0)
            truthy("state 0 drops timing", "<p:timing" not in s0)
        pmap = buildstates.page_map(deck, expanded)
        check("page map length", len(pmap), 4)
        check("page 2 is a build state", (pmap[1]["slide"], pmap[1]["state"]), (1, 1))
        check("page 4 is the plain slide", (pmap[3]["slide"], pmap[3]["state"]), (2, None))

        print("speaker notes")
        notes, page_of = extract_notes.slide_notes(deck)
        truthy("reads the body placeholder", notes.get(1, "").startswith("Fixture speaker notes."))
        truthy("ignores the slide-number placeholder", "7" not in notes.get(1, "").split("\n"))
        smap = os.path.join(tmp, "slides.json")
        json.dump({"pages": pmap}, open(smap, "w"))
        paged = extract_notes.to_pages(notes, smap)
        check("notes land on every build page", sorted(paged), ["1", "2", "3"])
        truthy("build pages are labelled", paged["2"].startswith("[build 1 of 2]"))

        # PowerPoint omits hidden slides from a PDF export. If we count them the
        # page map runs ahead of the real PDF and every later note lands on the
        # wrong page - silently, because the PDF still looks fine.
        print("hidden slides")
        hidden_deck = os.path.join(tmp, "hidden.pptx")
        with zipfile.ZipFile(deck) as zin, \
                zipfile.ZipFile(hidden_deck, "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                data = zin.read(it.filename)
                if it.filename == "ppt/slides/slide1.xml":     # hide the animated one
                    data = data.replace(b"<p:sld ", b'<p:sld show="0" ', 1)
                zout.writestr(it, data)
        truthy("detects show=\"0\"", buildstates.is_hidden(
            zipfile.ZipFile(hidden_deck).read("ppt/slides/slide1.xml")))
        truthy("visible slide is not flagged", not buildstates.is_hidden(
            zipfile.ZipFile(hidden_deck).read("ppt/slides/slide2.xml")))
        hexp = buildstates.expand_deck(hidden_deck, os.path.join(tmp, "hidden-exp.pptx"))
        check("hidden slide is not expanded", hexp, [])
        hmap = buildstates.page_map(hidden_deck, hexp)
        check("hidden slide takes no page", len(hmap), 1)
        check("survivor is slide 2 on page 1", (hmap[0]["page"], hmap[0]["slide"]), (1, 2))
        hnotes, hpage_of = extract_notes.slide_notes(hidden_deck)
        check("hidden slide has no page number", 1 in hpage_of, False)
        check("slide 2 shifts down to page 1", hpage_of.get(2), 1)

        # A single-page render used to be named "<deck>.<ext>" with no page
        # number, so rendering page 7 wrote over page 3 and the file did not
        # sort with the rest of the set.
        renderer = os.path.join(ROOT, "bin", "pdfrender")
        if os.path.exists(renderer):
            print("renderer page naming")
            rdir = os.path.join(tmp, "img")
            for page in ("2", "3"):
                subprocess.run([renderer, pdf, rdir, "800", "deck", "png", page],
                               check=True, capture_output=True)
            got = sorted(os.listdir(rdir))
            check("one file per requested page", len(got), 2)
            check("page number is kept", got, ["deck-02.png", "deck-03.png"])

        # A page number stops meaning anything once builds are expanded, so the
        # images are named for the slide they show, not the page they were.
        print("image naming")
        pmap = [{"page": 1, "slide": 1, "state": None, "of": None},
                {"page": 2, "slide": 2, "state": 0, "of": 10},
                {"page": 3, "slide": 2, "state": 9, "of": 10},
                {"page": 4, "slide": 3, "state": None, "of": None}]
        lab = name_images.labels_for(pmap)
        check("plain slide keeps its number", lab[1], "01")
        check("first build is .01", lab[2], "02.01")
        check("tenth build pads to .10", lab[3], "02.10")
        check("builds sort before the next slide", sorted(lab.values()),
              ["01", "02.01", "02.10", "03"])
        json.dump({"pages": pmap}, open(os.path.join(tmp, "pm.json"), "w"))
        lpath = os.path.join(tmp, "labels.txt")
        n = name_images.write_labels(os.path.join(tmp, "pm.json"), lpath)
        check("one label per page", n, 4)
        check("labels are in page order", open(lpath).read().split("\n"),
              ["01", "02.01", "02.10", "03"])
        if os.path.exists(renderer):
            ldir = os.path.join(tmp, "labelled")
            subprocess.run([renderer, pdf, ldir, "400", "deck", "png", "", lpath],
                           check=True, capture_output=True)
            check("renderer uses the labels", sorted(os.listdir(ldir)),
                  ["deck-01.png", "deck-02.01.png", "deck-02.10.png", "deck-03.png"])
            # rendering again must overwrite, not accumulate or shuffle
            subprocess.run([renderer, pdf, ldir, "400", "deck", "png", "", lpath],
                           check=True, capture_output=True)
            check("re-rendering is idempotent", len(os.listdir(ldir)), 4)

        print("media extraction")
        mdir = os.path.join(tmp, "media")
        items, links = extract_media.extract(deck, mdir)
        check("extracts the animated gif", len(items), 1)
        check("names it by slide", items[0]["file"].count("slide001"), 1)
        check("records the youtube link", len(links), 1)
        truthy("link keeps its slide number", links[0]["slide"] == 1)

        print("pdf notes, incremental")
        before = open(pdf, "rb").read()
        noted = os.path.join(tmp, "noted.pdf")
        n = add_notes.add_notes(pdf, noted, {"1": "First page note", "3": "Third page note"})
        check("annotates two pages", n, 2)
        after = open(noted, "rb").read()
        truthy("original bytes untouched", after[:len(before)] == before)
        truthy("only appended", len(after) > len(before))
        truthy("growth is small", len(after) - len(before) < 4096)
        truthy("trailer chains to the old xref", b"/Prev" in after[len(before):])
        truthy("annotation objects written", after.count(b"/Subtype /Text") == 2)

        print()
        if FAILURES:
            print("%d FAILED: %s" % (len(FAILURES), ", ".join(FAILURES)))
            return 1
        print("all checks passed")
        return 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
