#!/usr/bin/env python3
"""
Write the per-page image labels the renderer should use.

    python3 name_images.py <slides.json> <labels.txt>

The renderer numbers by PDF page, which stops meaning anything as soon as
animation builds are expanded: page 6 of a deck can be slide 2's fourth build
state, and hidden slides shift the count even when nothing was expanded. So the
page map is used to name each file after the slide it actually shows.

    <prefix>-04.png        slide 4, no builds
    <prefix>-02.1.png      slide 2, first build state
    <prefix>-02.5.png      slide 2, fifth build state

Slide numbers are the running order of the deck, so the images sort into the
order you would present them.

The renderer applies these itself rather than anything renaming files after the
fact: once an image is called "deck-03.png" there is no way to tell whether the
03 was a page or a slide, so a second pass would remap files that were already
correct.
"""
import json, os, re, sys


def labels_for(pages):
    """page number -> label, from a slides.json 'pages' list."""
    if not pages:
        return {}
    width = max(2, len(str(max(p["slide"] for p in pages))))
    # Pad the build number too, or a slide with ten states sorts 1, 10, 2, 3 -
    # which puts the last frame second in every file browser.
    most = max((p.get("of") or 1) for p in pages)
    swidth = max(2, len(str(most)))
    out = {}
    for p in pages:
        base = str(p["slide"]).zfill(width)
        # states are 0-based in the map; people count builds from one
        out[p["page"]] = (base if p.get("state") is None
                          else "%s.%s" % (base, str(p["state"] + 1).zfill(swidth)))
    return out


def write_labels(slides_json, out_path):
    """One label per line, in page order. Empty file if there is no map."""
    pages = []
    if slides_json and os.path.exists(slides_json):
        try:
            pages = json.load(open(slides_json)).get("pages", [])
        except (ValueError, OSError):
            pages = []
    labels = labels_for(pages)
    if not labels:
        return 0
    ordered = [labels[p] for p in sorted(labels)]
    with open(out_path, "w") as f:
        f.write("\n".join(ordered))
    return len(ordered)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__.strip())
        sys.exit(2)
    sys.exit(0 if write_labels(sys.argv[1], sys.argv[2]) else 1)
