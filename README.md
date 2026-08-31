# pptxtractor

Bulk-extract faithful representations of PowerPoint slides — vector PDFs,
animation build states, speaker notes, and the GIFs and video buried inside the
files — in one pass over a whole library.

## Why this exists

Generic .pptx converters are lossy in ways that matter if the slides are your
work. They substitute fonts you licensed, flatten every animation build into a
single muddled frame, silently drop the screen recordings and video embedded in
the deck, and throw away the speaker notes that explain what any of it meant.

pptxtractor drives **real PowerPoint** to do the rendering, so the layout,
gradients, shadows, masters and embedded fonts come out exactly as they look on
screen. Everything else — reading the animation timing tree, pulling media,
attaching notes, rasterising — is done directly against the file and the PDF.

The practical shape of this: **you need PowerPoint once, not forever.**
Microsoft offers a one-month Microsoft 365 trial. That is enough to point this
at a decade of decks, walk away, and come back to an archive you can mine
indefinitely without a licence. Built for designers, researchers and consultants
sitting on hundreds of old presentations that hold portfolio material.

## Requirements

- macOS
- Microsoft PowerPoint — for `export` only. `audit` and `render` need neither.
- Xcode Command Line Tools, to build the renderer (`xcode-select --install`)
- Python 3 (system Python is fine — no third-party packages)

Windows is not supported. The export path is AppleScript driving the Mac
PowerPoint app; a Windows port would use PowerPoint's COM automation and
`ExportAsFixedFormat`, which is a fair amount of work but a clean fork.

## Install

```
git clone https://github.com/andywilsondesign/pptxtractor.git
cd pptxtractor
make
```

Optionally put it on your PATH:

```
ln -s "$PWD/bin/pptxtractor" /usr/local/bin/pptxtractor
```

## Use

**Look before you leap.** No PowerPoint involved, nothing written:

```
pptxtractor audit ~/Decks --json audit.json
```

Reports, per deck: slides, aspect ratio, fonts that *will* be substituted,
animation build counts, embedded media, speaker notes, and likely duplicates.

**Export the archive:**

```
pptxtractor export --root ~/Decks --out /Volumes/Archive --states
```

**Render images later, from the archive:**

```
pptxtractor render /Volumes/Archive --format png  --edge 3840
pptxtractor render /Volumes/Archive --format jpeg --edge 2560 --only Acme
```

## What you get

```
Acme Q3 Review/
├── Acme Q3 Review.pdf     vector slides; builds inline; notes attached
├── slides.json            page -> slide/state map (when builds were expanded)
├── media/
│   ├── ...-slide014-loop.gif
│   └── media.json         manifest, including externally linked video
└── images/                only if you asked for PNG or JPEG
```

### PDF is the default on purpose

A PDF stays vector, so it is resolution-independent — pull a 4K or an 8K still
out of it years from now. It is roughly a tenth the size of rendered images,
keeps each deck as one browsable object, and its text stays searchable. Render
PNG or JPEG from it whenever you actually need pixels; PowerPoint is not
involved in that step, so it is fast and repeatable.

Use `--format png|jpeg` on either command if you want images immediately.
PNG for anything with flat colour, type or transparency. JPEG for
photograph-heavy slides where file size matters more than edge fidelity.

### Animation builds become pages

A slide with three click-triggered reveals becomes four consecutive PDF pages,
each a complete static frame. This is done by rewriting the deck so the slide is
followed by one copy per state, with the not-yet-revealed shapes removed — so
every state is a real vector page, not a screen grab. `slides.json` maps page
numbers back to slide and state.

Entrance and exit builds are exact. Motion paths and emphasis effects have no
meaningful still frame and are captured in their final position.

### Speaker notes ride along, invisibly

Notes attach to the PDF as sticky-note annotations. A PDF viewer lists them in
its notes sidebar; the image renderer never draws them, because annotations are
not page content. The same file gives you clean imagery and the full narrative.

They are written as an *incremental update* — the original PDF bytes are copied
verbatim and about 1 KB per note is appended. Nothing is re-encoded.

## Known limits

- **Fonts that are neither installed nor embedded in the deck are substituted.**
  PowerPoint cannot conjure them and no pipeline change fixes it. Run `audit`
  first, install what it flags, then export.
- **PowerPoint is sandboxed, and that shapes the whole design.** It can only
  open files you picked in a dialog or files inside its own container, so
  handing it a scratch directory makes it raise a modal "Grant File Access"
  prompt and wait — with no error, at 0% CPU, sometimes with no visible window.
  Because a temp path is different every run, the prompt comes back forever.
  `export` therefore works inside
  `~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/`, where no grant is
  needed. If you fork this, do not move that workspace outside the container.
- **Do not interrupt an export.** Force-quitting PowerPoint, or killing the
  `osascript` driving it, can wedge the app: it relaunches with no window, a
  pinned core, and every subsequent open failing. Only `pkill -9` and a
  relaunch clears it. The exporter resets and retries on its own, and gives up
  after three consecutive failures.
- **The export closes documents in PowerPoint as it works**, so it refuses to
  start while you have your own files open.
- **Notes annotation needs a classic cross-reference table.** PowerPoint for Mac
  produces these; a PDF from another tool using cross-reference streams is
  rejected with a clear error rather than corrupted.
- **Linked online video is recorded, not downloaded.** YouTube and Vimeo
  references have no file to extract, so their URLs and slide numbers go into
  `media/media.json`.

## For agents

`audit` emits JSON on stdout and a human summary on stderr, so it composes.
`export` and `render` are idempotent — decks that already have output are
skipped, so a run can be stopped and resumed freely.

If you are driving this from Claude Code or a similar agent, see
[`skills/pptxtractor/SKILL.md`](skills/pptxtractor/SKILL.md) and
[`AGENTS.md`](AGENTS.md).

## Tests

```
make test
```

Builds a .pptx and a PDF from scratch and checks the audit, the animation
expansion, notes extraction and the PDF incremental update. The
PowerPoint-driven export is not covered — CI runners have no Office licence, so
that path is only exercised by running it for real.

## Licence

MIT — see [LICENSE](LICENSE). Copyright Andy Wilson Design,
[andywilsondesign.com](https://andywilsondesign.com).

**Use at your own risk.** This drives another application and writes files. It
is read-only with respect to your source decks — every deck is copied to a temp
workspace before PowerPoint opens it, so no lock files or timestamp changes
reach your originals — but no warranty is given, as set out in the licence.
Test on a copy of anything you cannot replace.
