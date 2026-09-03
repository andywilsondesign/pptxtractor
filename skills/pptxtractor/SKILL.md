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
pptxtractor export <folder|deck.pptx> --out <folder> --states
```

Add `--dry-run` first if the user is nervous — it needs no PowerPoint, so it
also works for planning on a machine that has not got it. `--only TEXT` filters
by path.

**Always pass `--out` when driving this yourself.** Without it the archive goes
to `~/Documents/pptxtractor/<today>/`, which is fine for a person typing the
command and reading the output, but you should be putting it somewhere the user
asked for and telling them where that is.

Re-running over the same source is safe: decks already exported are skipped.
Without a terminal that is the fixed default, so a second run resumes rather
than redoing work. Pass `--on-conflict overwrite` or `new` only when the user
has asked to redo something.

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

**These live in [AGENTS.md](../../AGENTS.md), which is the single source of
truth for how the tool behaves — the sandbox, hidden slides, oversized decks,
what not to interrupt, and what each exit code means. Read it before driving a
run.** It is not duplicated here so the two can never drift apart.

The short version: exit status carries the outcome (1 decks failed, 3 no
PowerPoint, 4 automation denied, 5 unresponsive, 6 locked); codes 3-6 need a
person, so report them rather than retrying. Never kill an export in flight —
the tool resets and retries on its own, and interrupting it causes the wedge you
would then have to diagnose.

## Reporting back

Lead with what the user cares about: how many decks, how many slides, which
ones will lose fonts, and where the output is. The font finding is the one most
likely to change what they do next — surface it early, not in a footnote.
