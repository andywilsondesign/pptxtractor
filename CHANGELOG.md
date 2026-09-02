# Changelog

## Unreleased

Fixes found by running the tool over a 384-deck library.

- **Hidden slides no longer shift the page map.** PowerPoint omits `show="0"`
  slides from a PDF export, but the expansion counted them, so `slides.json`
  ran ahead of the real PDF and speaker notes attached to the wrong pages —
  silently, because the PDF still looked correct. A deck with one hidden
  5-state slide mid-deck produced a 27-page map for a 22-page PDF. Hidden
  slides are now skipped by both the page map and the notes mapping.
- **`audit` no longer reports installed fonts as missing.** The font scan used
  a non-recursive `listdir`, so it never saw
  `/System/Library/Fonts/Supplemental/` — where macOS keeps Arial, Verdana,
  Trebuchet MS, Georgia, Gill Sans and ~290 other faces. It flagged 340 of 384
  decks for substitution; the true figure was 129.
- **`audit`'s build-state forecast now matches what `export` produces.** It
  counted every `clickEffect`, including motion-path and emphasis effects that
  have no still frame and yield no page. One deck forecast 42 extra pages and
  produced none.
- **Notes are keyed by presentation order**, not by `slideN.xml` file number,
  which can differ from the running order.
- **The export log no longer lands in the source tree.** It records deck paths
  and file names; it now sits beside the archive being built (`--out`), or in
  the temporary workspace. Override with `PPTXTRACTOR_LOG`.
- **`render --page N` no longer overwrites its own output.** A single-page
  render was named `<deck>.<ext>` with no page number, so rendering page 7 wrote
  over page 3. Pages are always numbered now.
- **`render --format jpeg` no longer skips an archive already rendered to PNG.**
  The "already has images" test looked for any file rather than the format
  asked for.
- **`render`'s page count is the number of pages rendered**, not the number of
  files in the folder - it reported 47 for a 23-page deck holding both formats.
- **`export` checks a deck is a readable package before opening PowerPoint.**
  A truncated .pptx — a zip whose central directory never arrived — satisfies
  `file` and opens as nothing; handed to PowerPoint it hangs, costing two full
  deadlines. Three such files turned up in a 384-deck library. `--dry-run`
  reports them too.
- **`--deadline` now defaults to 240s, down from 600s.** The slowest successful
  export measured was 34s (95 slides, 116 pages, notes and media); decks that
  exceed that do not run slowly, they hang. A wedged deck now costs 8 minutes
  including the retry instead of 20.
- `.gitignore` now refuses decks, PDFs and audit output, so private material
  cannot be committed to a public clone by accident.

## 1.0.0

First release.

- `audit` — read a tree of .pptx files without opening PowerPoint and report
  slides, aspect ratios, font-substitution risk, animation builds, embedded
  media, speaker notes and duplicates. Emits JSON.
- `export` — drive PowerPoint to produce a vector PDF per deck, with animation
  build states expanded into extra pages, speaker notes attached as annotations,
  and playable media extracted alongside.
- `render` — produce PNG or JPEG at any size from the exported PDFs, without
  PowerPoint.
