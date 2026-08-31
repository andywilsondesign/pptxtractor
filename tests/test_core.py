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
        notes = extract_notes.slide_notes(deck)
        truthy("reads the body placeholder", notes.get(1, "").startswith("Fixture speaker notes."))
        truthy("ignores the slide-number placeholder", "7" not in notes.get(1, "").split("\n"))
        smap = os.path.join(tmp, "slides.json")
        json.dump({"pages": pmap}, open(smap, "w"))
        paged = extract_notes.to_pages(notes, smap)
        check("notes land on every build page", sorted(paged), ["1", "2", "3"])
        truthy("build pages are labelled", paged["2"].startswith("[build 1 of 2]"))

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
