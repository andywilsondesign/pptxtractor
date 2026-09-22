#!/usr/bin/env python3
"""
Tests for everything that does not need PowerPoint.

The PowerPoint-driven export cannot be covered here - CI runners have no Office
licence - so this exercises the parts that do the actual thinking: the audit,
the animation expansion, notes extraction and the PDF incremental update.
"""
import json, os, re, shutil, subprocess, sys, tempfile, zipfile

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
import strip_fonts                               # noqa: E402
import fonts                                      # noqa: E402

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
        check("no hidden slides in the plain fixture", d["hidden_slides"], 0)

        # PowerPoint never exports a hidden slide, so the PDF comes out shorter
        # than the slide count. Unreported, that gap looks like dropped work.
        print("hidden slides")
        hid_dir = os.path.join(tmp, "hidden")
        os.makedirs(hid_dir)
        make_fixture.build_hidden_pptx(os.path.join(hid_dir, "hidden.pptx"))
        hres = audit_mod.audit(hid_dir)
        hd = hres["decks"][0]
        check("counts every slide", hd["slides"], 3)
        check("counts the hidden one", hd["hidden_slides"], 1)
        check("totals hidden slides", hres["summary"]["hidden_slides"], 1)
        check("counts decks holding them", hres["summary"]["decks_with_hidden"], 1)

        print("a deck with no slides")
        ns_dir = os.path.join(tmp, "noslides")
        os.makedirs(ns_dir)
        make_fixture.build_noslides_pptx(os.path.join(ns_dir, "Template.pptx"))
        nres = audit_mod.audit(ns_dir)
        nd = nres["decks"][0]
        check("reads it without error", nd["error"], None)
        check("reports zero slides", nd["slides"], 0)

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

        # PowerPoint's default build for a bulleted list animates one text
        # box paragraph by paragraph. Read as whole-shape steps it produces a
        # run of identical frames and then one where everything appears.
        print("paragraph-level builds")
        lines = ["First point", "Second point", "Third point", "Fourth point"]
        para_slide = make_fixture.slide(
            [make_fixture.shape(2, "Title", "Heading"),
             make_fixture.bullets(3, "Body", lines)],
            make_fixture.timing_paragraphs(3, 4)).encode("utf8")
        steps = buildstates.click_steps(para_slide)
        check("four click steps", len(steps), 4)
        check("no whole-shape targets", sum(len(a) + len(b) for a, b, _c, _d in steps), 0)
        check("four paragraph targets",
              sum(len(c) + len(d) for _a, _b, c, d in steps), 4)

        def para_count(x):
            return len(re.findall(rb"<a:p(?=[\s/>])", x))

        counts = [para_count(buildstates.state_xml(para_slide, k, steps))
                  for k in range(len(steps) + 1)]
        # State 0 is the title plus an empty placeholder paragraph: the body
        # has had all its bullets hidden, and a <p:txBody> may not be left
        # with none. From the first click on, that placeholder is replaced by
        # real content and each click adds one more bullet.
        check("paragraphs grow one per click", counts, [2, 2, 3, 4, 5])
        truthy("the shape itself always survives",
               all(b"Body" in buildstates.state_xml(para_slide, k, steps)
                   for k in range(len(steps) + 1)))
        truthy("the last bullet appears only at the end",
               b"Fourth point" in buildstates.state_xml(para_slide, 4, steps)
               and b"Fourth point" not in buildstates.state_xml(para_slide, 3, steps))
        truthy("the first bullet appears at the first click",
               b"First point" in buildstates.state_xml(para_slide, 1, steps)
               and b"First point" not in buildstates.state_xml(para_slide, 0, steps))
        check("no two states are the same",
              len(buildstates.distinct_states(para_slide, steps)), 5)

        # A <p:txBody> must hold at least one <a:p>; the schema cannot express
        # "no text". Hiding every paragraph of a box leaves it empty, and
        # PowerPoint then offers to repair the file instead of opening it -
        # which, driven by AppleScript, is a prompt nobody answers and reports
        # as a timeout. The earlier checks here counted paragraphs across the
        # whole slide and so never noticed the body had been emptied.
        def empty_bodies(x):
            return sum(1 for m in re.finditer(rb"<p:txBody>(.*?)</p:txBody>", x, re.S)
                       if not re.search(rb"<a:p(?=[\s/>])", m.group(1)))

        check("no text body is ever left empty",
              [empty_bodies(buildstates.state_xml(para_slide, k, steps))
               for k in range(len(steps) + 1)], [0, 0, 0, 0, 0])

        # the same slide with *every* paragraph animated is the case that broke
        only_bullets = make_fixture.slide(
            [make_fixture.bullets(3, "Body", ["One", "Two", "Three"])],
            make_fixture.timing_paragraphs(3, 3)).encode("utf8")
        obsteps = buildstates.click_steps(only_bullets)
        check("every paragraph animated, body still legal",
              [empty_bodies(buildstates.state_xml(only_bullets, k, obsteps))
               for k in range(len(obsteps) + 1)], [0, 0, 0, 0])
        truthy("state 0 shows none of the bullets",
               b"One" not in buildstates.state_xml(only_bullets, 0, obsteps))
        truthy("the final state shows them all",
               all(w in buildstates.state_xml(only_bullets, 3, obsteps)
                   for w in (b"One", b"Two", b"Three")))

        # A step that changes nothing visible still costs a page. Comparing the
        # generated XML catches that; comparing rendered pages would not,
        # because each page carries its own slide number.
        print("duplicate states")
        dead = make_fixture.slide(
            [make_fixture.shape(2, "Only", "Nothing animates here")],
            make_fixture.timing_paragraphs(99, 3)).encode("utf8")
        dsteps = buildstates.click_steps(dead)
        check("steps are still read", len(dsteps), 3)
        check("but collapse to one state",
              len(buildstates.distinct_states(dead, dsteps)), 1)

        # Checking our own output before PowerPoint has to. Handed a deck it
        # considers damaged, PowerPoint puts up a repair dialog and waits -
        # which, driven by AppleScript, costs a whole deadline and is then
        # reported as a timeout, the one explanation that is certainly wrong.
        print("validating a generated deck")
        good = os.path.join(tmp, "good.pptx")
        buildstates.expand_deck(deck, good)
        check("a deck we built passes", buildstates.validate_deck(good), [])

        def rebuild(path, mangle):
            src = zipfile.ZipFile(good)
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as w:
                for i in src.infolist():
                    w.writestr(i.filename, mangle(i.filename, src.read(i.filename)))
            src.close()
            return path

        # the fault that reached a real library: every paragraph hidden
        empty = rebuild(os.path.join(tmp, "empty.pptx"), lambda n, d:
                        re.sub(rb"<a:p>.*?</a:p>", b"", d, flags=re.S)
                        if n == "ppt/slides/slide1.xml" else d)
        truthy("an emptied text body is caught",
               any("no paragraphs" in f for f in buildstates.validate_deck(empty)))

        broken = rebuild(os.path.join(tmp, "broken.pptx"), lambda n, d:
                         d.replace(b"</p:sld>", b"") if n == "ppt/slides/slide1.xml" else d)
        truthy("malformed XML is caught",
               any("well-formed" in f for f in buildstates.validate_deck(broken)))

        nolayout = rebuild(os.path.join(tmp, "nolayout.pptx"), lambda n, d:
                           b'<Relationships xmlns="http://schemas.openxmlformats.org'
                           b'/package/2006/relationships"/>'
                           if n == "ppt/slides/_rels/slide1.xml.rels" else d)
        truthy("a slide with no layout is caught",
               any("slide layout" in f for f in buildstates.validate_deck(nolayout)))

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

        # A deck can embed fonts licensed for preview-and-print only. PowerPoint
        # will not open one for editing without a modal prompt, which nobody is
        # there to answer, so the export hangs until the watchdog kills it.
        # Removing the embedded fonts removes the question.
        print("embedded fonts")
        fdeck = os.path.join(tmp, "fonts.pptx")
        with zipfile.ZipFile(deck) as zin, \
                zipfile.ZipFile(fdeck, "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                data = zin.read(it.filename)
                if it.filename == "ppt/presentation.xml":
                    data = data.replace(b"</p:presentation>",
                                        b'<p:embeddedFontLst><p:embeddedFont>'
                                        b'<p:font typeface="Graphik"/>'
                                        b'</p:embeddedFont></p:embeddedFontLst>'
                                        b'</p:presentation>')
                elif it.filename == "ppt/_rels/presentation.xml.rels":
                    data = data.replace(b"</Relationships>",
                                        b'<Relationship Id="rIdFont9" Target="fonts/font1.fntdata" '
                                        b'Type="http://schemas.openxmlformats.org/officeDocument/'
                                        b'2006/relationships/font"/></Relationships>')
                zout.writestr(it, data)
            zout.writestr("ppt/fonts/font1.fntdata", b"\x00" * 64)
        truthy("detects embedded fonts", strip_fonts.has_embedded_fonts(fdeck))
        truthy("plain deck has none", not strip_fonts.has_embedded_fonts(deck))
        clean = os.path.join(tmp, "nofonts.pptx")
        check("reports what it removed", strip_fonts.strip(fdeck, clean), 1)
        with zipfile.ZipFile(clean) as z:
            names = z.namelist()
            pres = z.read("ppt/presentation.xml").decode()
            rels = z.read("ppt/_rels/presentation.xml.rels").decode()
            truthy("font part gone", not any(n.startswith("ppt/fonts/") for n in names))
            truthy("embeddedFontLst gone", "<p:embeddedFontLst" not in pres)
            truthy("no dangling font relationship", "fonts/font1.fntdata" not in rels)
            check("slides untouched", len([n for n in names
                  if n.startswith("ppt/slides/slide") and n.endswith(".xml")]), 2)
        check("nothing to strip is reported", strip_fonts.strip(deck, os.path.join(tmp, "x.pptx")), 0)

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

        # Everything here runs on system Python. Choosing and building the
        # substitutes needs fontTools, but naming - which is where the subtle
        # mistakes live - does not, so it is covered on any machine.
        print("font names")
        for spelling, want in [
                # Linotype's numeric weights, and LT meaning two different
                # things in one name: the foundry, then Light.
                ("Helvetica Neue LT Std 75", ("Helvetica Neue", "Bold")),
                ("HelveticaNeueLTStd-Lt", ("Helvetica Neue", "Light")),
                ("HelveticaNeueLT Std Bold Cn", ("Helvetica Neue", "Bold")),
                # camelCase must not strand "Semi" in the family name
                ("Source Sans Pro SemiBold", ("Source Sans Pro", "SemiBold")),
                # a Monotype web cut, wedged between digits
                ("ObjektivMk3W03-Bold", ("Objektiv Mk3", "Bold")),
                # "Pro" is part of the name, not foundry noise
                ("Source Code Pro", ("Source Code Pro", None))]:
            fam, style, _w, _o = fonts.parse_face(spelling)
            check("reads %s" % spelling, (fam, style), want)

        check("keeps a width apart from the family",
              fonts.parse_face("Aptos Narrow (Body)")[:1], ("Aptos",))
        truthy("spots a CSS keyword", fonts.is_noise("ui-sans-serif"))
        truthy("spots a corrupted name", fonts.is_noise("Hel\\"))
        check("passes a real face", fonts.is_noise("Graphik Semibold"), None)

        print("grouping")
        subst = [{"face": "Graphik", "family": "Graphik", "style": "Regular",
                  "weight": 400, "decks": 1, "paths": ["a.pptx"]},
                 {"face": "Graphik Thin", "family": "Graphik", "style": "Thin",
                  "weight": 100, "decks": 1, "paths": ["a.pptx"]},
                 # a named cut belongs with its family, not beside it
                 {"face": "Graphik Courant", "family": "Graphik Courant",
                  "style": "Regular", "weight": 400, "decks": 1, "paths": ["b.pptx"]},
                 {"face": "Proxima Nova", "family": "Proxima Nova",
                  "style": "Regular", "weight": 400, "decks": 1, "paths": ["c.pptx"]}]
        groups = fonts.group_substitutes(subst)
        check("one group per typeface", len(groups), 2)
        biggest = groups[0]
        check("the family absorbs its cuts", biggest["family"], "Graphik")
        check("all three faces land in it", len(biggest["faces"]), 3)
        check("and both decks", len(biggest["decks"]), 2)

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
