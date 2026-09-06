# Changelog

## Unreleased

- **Workspaces left by killed runs are swept up.** The EXIT trap covers a normal
  finish and a Ctrl-C, but not a `kill -9` or a crash, and the leftovers sit
  inside PowerPoint's container holding copies of decks. A run now removes any
  whose process is gone, leaving live ones alone. Not done on `--dry-run`, which
  still writes and deletes nothing.
- The name-clash notice no longer prints a blank line when the clash was
  predicted from the source tree rather than found on disk — there is no
  previous deck to name in that case, so it says so.

## 1.0.1

Found by a fresh pair of eyes cloning the repo cold, and by pointing it at a
real 156-deck library.

- **A trailing slash on the source path scattered the output.** `export decks/`
  put the entire absolute source path underneath `--out`, instead of one folder
  per deck. Tab-completing a directory adds that slash, so this was the normal
  way to type it. Paths are normalised now, and `tests/test_paths.sh` covers all
  four slash combinations.
- **Flattening no longer loses decks.** With `--layout flat`, two different
  decks sharing a name both claimed one folder, and the second was reported as
  "already exported" and skipped — a real library here had ten distinct decks
  called `Personas Template.pptx`, one per cohort, so nine would have vanished
  while the run reported success. An export folder now records which deck wrote
  it, so "I already did this one" is told apart from "a different deck wants
  this name", and the latter is named after the folder that distinguishes it:
  `Personas Template (Kiwi) (export)`. Names the tree uses more than once get
  that treatment from the first deck onwards, so a set reads consistently.
- **`--dry-run` predicts clashes** instead of only revealing them mid-run.
- **`audit` accepts a single .pptx**, matching `export`.
- The pre-run notice says that recovering from a stuck deck force-quits
  PowerPoint, so anything you have open goes with it.

## Unreleased

Fixes found by running the tool over a 384-deck library, plus the checks a
public release needs.

### A tractor

- **`export` now shows a field being ploughed** — furrows for decks done, crop
  for what is left — and a yield when it finishes. A long run is mostly waiting,
  and this makes the waiting legible.
- **`pptxtractor plough`** plays the whole thing on demand, so you can see what
  a run looks like without committing a library to it.
- It stays out of the way: stderr only, never stdout, and switched off entirely
  when the output is not a terminal. Scripts, agents and CI see plain text.
  `NO_COLOR=1` or `PPTXTRACTOR_PLAIN=1` turn it off for people too. CI checks
  that it cannot leak into `audit`'s JSON.
- **Failures explain whose problem they are.** An unreadable deck is a damaged
  file that PowerPoint cannot open either; a deck PowerPoint gives up on is
  usually just too large. Neither is something the user did wrong, and the
  messages now say so.

### Where things go, and who decides

- **Images are named for the slide they show, not the PDF page they were.**
  `deck-04.png` is slide 4; `deck-02.01.png` … `deck-02.05.png` are slide 2's
  build states in order. Page numbers stopped meaning anything as soon as builds
  were expanded — page 6 of one deck here was slide 2's fourth build — and
  hidden slides shift the count even with no expansion. Build numbers are padded
  so a slide with ten states does not sort 1, 10, 2.
- **The page map is written for every export**, not only when `--states`
  expanded something, so the mapping from page to slide is always on disk.

- **`export` takes a path, and it can be a single deck.**
  `pptxtractor export ~/Decks` or `pptxtractor export "Q3 Review.pptx"`.
  `--root` still works.
- **Each deck becomes `<name> (export)/`** — one folder holding its PDF, media,
  images and page map, so a deck stays one object however much you ask for and
  reads as output when it sits beside the original.
- **`--layout mirror|beside|flat`.** Mirror (the default) recreates the source
  folders under `--out`; beside writes next to each deck; flat puts everything
  in one folder.
- **Without `--out`, exports go to `~/Documents/pptxtractor/<today>/`** — said
  before the run starts and again when it finishes, with the command to open it.
  Writing somewhere the user did not pick is only acceptable if they are told.
- **`--on-conflict ask|skip|overwrite|new`** decides what happens to a deck that
  was already exported. In a terminal it asks per deck, and `!` on the answer
  (`o!`, `s!`) applies it to everything remaining. Without a terminal it is
  always `skip`, so nothing can block waiting for input.
- **A run explains itself before it starts**: leave PowerPoint alone, keep the
  machine awake, and Ctrl-C between decks is safe.
- **`audit` flags what will cost you later** — decks over 200 MB, which can hang
  PowerPoint, and files that cannot be read at all.
- `--dry-run` no longer creates anything, not even the folder it would have
  logged into.

### Ready for someone else's machine

- **`export` says so plainly when PowerPoint is missing**, with a link to get
  it, instead of failing somewhere further in. It also refuses versions older
  than 2016, whose AppleScript dictionary has no `save as PDF`.
- **The Automation permission prompt is handled.** macOS asks once, per calling
  program, before one app may control another; until it is granted every command
  returns `-1743` and the run looks wedged. That is now detected and explained.
- **A PowerPoint sitting on a dialog no longer hangs the start-up probe.** A
  fresh install can stop on sign-in, activation or "What's New"; every probe is
  now bounded and says what to clear.
- **Two exports can no longer fight over one PowerPoint.** The exporter closes
  open presentations before each deck, so a concurrent run would sabotage the
  first. A lock refuses the second, and clears itself if a run died.
- **The workspace is derived from PowerPoint's real bundle id**, not a hardcoded
  path, and falling back to a temp folder now warns instead of silently taking
  the slow, prompt-ridden route.
- **`export` exits non-zero when decks fail** (1), and has distinct codes for a
  missing PowerPoint (3), denied automation (4), an unresponsive app (5) and a
  held lock (6). Scripts and agents can act on the outcome without parsing
  stdout.
- CI now proves the no-PowerPoint paths on a runner that genuinely has none.
- Documented how to run over a large library: `audit` first, `caffeinate -i` so
  an idle-sleeping machine is not mistaken for a hung one, and export to local
  disk rather than into a synced folder.
- Documented what the PowerPoint version does and does not change. The PDF is
  written by macOS's Quartz engine, not by PowerPoint, so its structure follows
  the OS; what varies by Office release is which fonts ship and how newer slide
  features survive an older renderer.
- Requirements are stated: macOS 11+, PowerPoint 2016+, Python 3.9+.
- Agent guidance is single-sourced. `AGENTS.md` is the one place behaviour is
  described; `CLAUDE.md` is a symlink to it and `SKILL.md` points at it instead
  of restating it, so the two cannot drift apart.

### Correctness

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
