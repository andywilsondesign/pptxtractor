---
name: pptxtractor
description: Bulk-extract faithful representations from PowerPoint decks - vector PDFs, animation build states, speaker notes, and embedded GIF/video - across a whole folder tree. Use when someone needs to get slides out of many .pptx files at once, wants portfolio imagery from old presentations, asks why an export lost fonts or animations or media, or wants to know what is inside a library of decks before exporting. Requires macOS; export requires Microsoft PowerPoint.
---

# pptxtractor

Extract slides from PowerPoint libraries without the losses generic converters
introduce: substituted fonts, animation builds flattened into one frame, dropped
media, discarded speaker notes.

## Always audit first

`audit` needs neither PowerPoint nor write access. Run it before anything else
and read the JSON — it tells you what the export will and will not manage.

```
pptxtractor audit <folder> --json audit.json
```

Each deck record carries `slides`, `aspect`, `fonts_missing`,
`animated_slides`, `build_states`, `gifs`, `videos`, `external_media`,
`noted_slides`, `duplicate_of`. The `summary` object totals them.

Act on it:

- **`fonts_missing` is non-empty** → those decks will export with substituted
  fonts. Say so plainly and name the faces. Installing them is the only fix;
  no flag works around it. Offer to export the clean decks now and the risky
  ones after the fonts are installed.
- **`duplicate_of` is set** → the same deck exists elsewhere in the tree.
  Exporting both wastes time. Offer to skip.
- **`build_states` > 0** → recommend `--states`, and say how many extra pages
  it will add.
- **`external_media`** → linked YouTube/Vimeo, nothing to extract. The URLs are
  recorded in the output; mention it rather than promising video files.

## Then export

```
pptxtractor export --root <folder> --out <folder> --states
```

Add `--dry-run` first if the user is nervous. `--only TEXT` filters by path.

Default output is a vector PDF per deck. **Recommend keeping it that way.**
Images can be rendered from the PDFs at any time and any size; rendering
everything up front costs roughly ten times the disk for pixels the user has
not chosen yet. Only pass `--format png|jpeg` when they have said they want
images immediately.

## Then render, if asked

```
pptxtractor render <archive> --format png --edge 3840
```

PowerPoint is not involved. Safe to re-run at a different size.

## Operational rules that matter

**If exports hang, suspect the sandbox before anything else.** PowerPoint can
only open files the user picked in a dialog or files inside its own container.
Given any other path it shows a modal "Grant File Access" prompt and waits
silently, sometimes with no visible window. The tool avoids this by working
inside the PowerPoint container — but if you see hangs at 0% CPU, check for that
dialog rather than assuming the deck is at fault.

**Never kill a running export.** Interrupting it can wedge PowerPoint: no
window, one core pinned, every open failing, clearable only with `pkill -9` and
a relaunch. The tool resets and retries on its own and stops after three
consecutive failures. If you interrupt it, you cause the failure you then have
to diagnose.

**A deck that fails is usually not the deck's fault.** It is almost always a
wedged PowerPoint. Let the tool reset and retry before concluding anything
about the file.

**The export closes open PowerPoint documents.** It refuses to start if the
user has files open. Tell them to save and close rather than forcing it.

**Long runs belong in the background.** A large library takes tens of minutes.
Start it in the background and report progress rather than blocking.

## Reporting back

Lead with what the user cares about: how many decks, how many slides, which
ones will lose fonts, and where the output is. The font finding is the one most
likely to change what they do next — surface it early, not in a footnote.
