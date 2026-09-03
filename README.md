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

The practical shape of this: **you need PowerPoint once, not forever.** Point it
at a decade of decks, walk away, and come back to an archive you can mine
indefinitely — the PDFs stay useful whether or not you still have a licence.
Built for designers, researchers and consultants sitting on hundreds of old
presentations that hold portfolio material.

## Requirements

- macOS 11 Big Sur or later
- **Microsoft PowerPoint for Mac, 2016 or newer** — for `export` only.
  `audit` and `render` never touch it, and `export --dry-run` works without it.
- Xcode Command Line Tools, to build the renderer (`xcode-select --install`)
- Python 3.9 or later (the system Python is fine — no third-party packages)

### About the PowerPoint requirement

There is no way around this one. `export` works by driving the real PowerPoint
app, which is exactly what keeps your licensed fonts, gradients, masters and
layouts intact — the thing every generic .pptx converter gets wrong. Nothing
else on the machine can stand in for it.

Any licensed copy of PowerPoint for Mac 2016 or newer works, including the one
in a [Microsoft 365 subscription](https://www.microsoft.com/en-us/microsoft-365/powerpoint).
Microsoft usually offers a free trial, which is enough to run this over a whole
library in an afternoon.

**You need it once, not forever.** Point this at a decade of decks, walk away,
and the archive it leaves behind — vector PDFs you can mine at any resolution —
outlives the subscription. `render` keeps working long after the licence lapses.

Office 2011 is not supported: its AppleScript dictionary predates the
`save as PDF` verb this relies on, and it is long out of support. `export`
checks the version at startup and tells you if it is too old.

### Does the PowerPoint version change the output?

Less than you would expect, because **PowerPoint does not write the PDF — macOS
does.** Every file this produces reports `macOS Version …` as its producer:
PowerPoint draws the slides and hands them to the system's Quartz PDF engine.
So the PDF structure — version, cross-reference format, how text and vectors are
encoded — follows your macOS, not your Office release. That is also why
attaching notes is safe across Office versions: it needs a classic
cross-reference table, and Quartz always writes one. (If that ever changed, the
notes step refuses with a clear error rather than producing a broken file.)

What *does* vary with the PowerPoint version:

- **Which fonts are installed.** Office ships its own — recent releases include
  Aptos, older ones Calibri. A deck can therefore substitute on one machine and
  not on another. `audit` reports what your machine will actually do.
- **Newer slide features.** A deck built in a current PowerPoint and exported by
  an older one may lose or approximate things that version never supported. If
  you have the choice, export with a PowerPoint at least as new as the decks.

Within 2016 → Microsoft 365 the export path itself is unchanged: same
AppleScript verb, same sandbox model, same container layout.

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

It also names the two things worth dealing with before you start: decks big
enough to hang PowerPoint, and files it could not read at all.

**Export the archive.** A whole tree, or a single deck:

```
pptxtractor export ~/Decks --out /Volumes/Archive --states
pptxtractor export "~/Decks/Acme Q3.pptx" --out /Volumes/Archive
```

Without `--out`, exports go to `~/Documents/pptxtractor/<today>/`, and the run
tells you so before it starts and again when it finishes.

**Where each export lands** is up to you:

| `--layout` | result |
|---|---|
| `mirror` (default) | recreates your source folders under `--out` |
| `beside` | writes next to each deck, in its own source folder |
| `flat` | every export in one folder |

**Running it again** skips decks that already have an export, so a stopped run
resumes for free. When something *has* been exported before, `--on-conflict`
decides: `ask` (the default in a terminal), `skip`, `overwrite`, or `new` to
keep both. Answer a prompt with `!` — `o!`, `s!` — to apply it to every
remaining deck. Outside a terminal the default is always `skip`, so scripts and
agents behave the same way every time and never block on a prompt.

**Render images later, from the archive:**

```
pptxtractor render /Volumes/Archive --format png  --edge 3840
pptxtractor render /Volumes/Archive --format jpeg --edge 2560 --only Acme
```

## What you get

```
Acme Q3 Review (export)/
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

## Running over a large library

`audit` first, always. It is instant, needs no PowerPoint, and tells you what
you are in for — including the two things worth fixing before you start:
duplicates, and fonts you should install.

Then, for anything longer than a few minutes:

```
caffeinate -i pptxtractor export --root ~/Decks --out /Volumes/Archive --states
```

`caffeinate -i` stops the machine idle-sleeping mid-run. An export that sleeps
looks exactly like an export that hung, and burns a `--deadline` before it
recovers.

**Export to local disk, not to a synced folder.** Writing straight into Dropbox,
iCloud Drive, Proton Drive or a network share means the client tries to upload
part-written PDFs while PowerPoint is still producing them, which is slow at
best. Export locally, check the result, then copy the finished archive up.

The run is resumable — decks that already have a PDF are skipped — so stopping
it and restarting later costs nothing.

## Known limits

- **The first run asks macOS for permission to control PowerPoint.** Approve it
  and it never asks again. If you miss the prompt or decline it, `export` stops
  with error `-1743` and tells you where to fix it: System Settings → Privacy &
  Security → Automation.
- **One export at a time.** The exporter closes open presentations before each
  deck, so two runs would sabotage each other. A second run is refused while the
  first holds the lock.
- **Hidden slides are not exported.** PowerPoint leaves `show="0"` slides out
  of a PDF, so the archive does too, and the page map skips them. If you want a
  hidden slide in the archive, unhide it in PowerPoint before exporting.
- **Very large decks can wedge PowerPoint.** A 333 MB deck sat at 0% CPU
  indefinitely with the file open and no dialog; the per-deck `--deadline`
  watchdog is what stops one such deck stalling a whole run. Check `audit` for
  outliers before a big export.
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
