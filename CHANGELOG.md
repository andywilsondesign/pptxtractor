# Changelog

## Unreleased

Found by running a 156-deck library end to end. Every item below cost that run
something real.

- **A bulleted list built one line at a time no longer exports as a run of
  identical pages.** PowerPoint's default build for a list animates a single
  text box paragraph by paragraph, and that reaches the XML as several click
  steps that all name the same shape, distinguished only by a
  `<p:txEl><p:pRg>`. `click_steps` read the spid and ignored the range, so
  every step looked like "hide shape 8" - and the expansion produced six
  copies of the same header-only slide followed by one where the whole list
  appeared at once. Steps are now read as paragraph ranges where they carry
  one, and the expander removes individual `<a:p>` elements instead of the
  shape that holds them, so the build comes out as the reveal it actually is.
  On the library that fixed 8 slides across 4 decks; one slide went from six
  identical frames to a six-step build.

- **The expansion checks its own output before PowerPoint sees it.** Handed a
  deck it considers damaged, PowerPoint does not fail - it raises a "found a
  problem with content" dialog and waits. Driven by AppleScript nobody answers
  that, so the run burns a whole deadline and the deck is then reported as a
  timeout, which is the one explanation certainly wrong. Two separate faults
  reached a real library this way, each costing about forty minutes per deck
  to learn nothing. `expand` now validates the deck it just wrote - XML
  well-formedness, text bodies with no paragraphs, slides with no layout or
  content type, duplicate slide ids, relationships resolving to nothing - and
  exits non-zero if any of it fails, so the caller falls back to the untouched
  deck. The cost is the build pages; the alternative was a modal and a dead
  deadline.

- **"Timed out" no longer covers for "PowerPoint is asking you something".**
  A deck that is slow and a deck waiting on a dialog look identical from the
  outside: the script simply does not return. They are not the same problem,
  and reporting the second as a timeout sends the reader after deck size when
  the answer is a prompt on screen. Converting burns CPU and a dialog does
  not, so the watchdog now watches that: if PowerPoint has done nothing for
  the last minute or more, it says a dialog is the likely cause, how long it
  has been idle, and that a longer `--deadline` will not help.

- **Hiding every paragraph of a text box no longer produces a file PowerPoint
  offers to repair.** A `<p:txBody>` must hold at least one `<a:p>`; the schema
  has no way to say "no text". The paragraph-level build above hid whole boxes
  when every one of their paragraphs was animated, and PowerPoint then asked to
  repair the deck rather than opening it - a prompt AppleScript can never
  answer, so the deck reported as a timeout with nothing in the log to explain
  it. An empty paragraph is left behind instead: it shows nothing and keeps the
  file legal.

- **Each part keeps the compression it arrived with.** Re-packing a deck is not
  this tool's business. One 349 MB deck here stores 341 of its 858 parts
  uncompressed, including an 86 MB TIFF; deflating all of it turned a package
  PowerPoint could read straight through into 350 MB it had to inflate before
  showing a slide, and shrank the file by 58% for no reason. Only the slides
  that were rewritten differ now; everything else goes back byte for byte.

- **A state that changes nothing no longer costs a page.** `distinct_states`
  drops a build state whose generated XML matches the one before it. This is
  deliberately done on the markup rather than on rendered images: two such
  pages are *not* pixel-identical, because each carries its own slide number,
  so comparing renders finds a difference and keeps the duplicate. Comparing
  the XML is exact, needs no rendering, and has no tolerance to tune. It
  removed 8 further redundant pages across 2 decks that the paragraph fix
  did not touch - different cause, same symptom, which is why the guard is
  worth having on its own.

- **New: `pptxtractor fonts`.** `audit` could say which faces were missing but
  not what to do about them, and on that library it was the largest piece of
  preparation by a distance — 30 faces across 45 decks, resolved by hand.

  The command separates three problems that all look like "missing font".
  Faces **already present under another name** are matched to what is here:
  `Helvetica Neue LT Std 75` is Linotype's spelling of Helvetica Neue Bold,
  `HelveticaNeueLTStd-Lt` is Light, and `Aptos Display` is a cut of the Aptos
  that ships inside PowerPoint — five faces covering 23 of those 45 decks, none
  of which needed substituting at all. Things that **are not fonts** are named
  as such rather than counted: `ui-sans-serif` is a CSS keyword that came in
  with a paste from a browser, and `Hel\` is a corrupted string.

  What is left gets a stand-in chosen by **measured width**. Slides are fixed
  geometry, so a substitute wider than the face it replaces pushes text out of
  its box; candidates are measured against Helvetica Neue on real text from the
  decks that want the missing face, and overshoot is disqualifying before genre
  or looks are weighed. This matters more than it sounds: the obvious
  aesthetic choice was wrong in three of four cases here — Montserrat measured
  **+11.1%** against the Proxima Nova it was standing in for and Inter
  **+5.5%** against Graphik, while Archivo, Figtree and Overpass all came in
  just under the anchor.

  `--install` builds each face at the weight the decks ask for and installs it
  with the OFL licence of every family used and an `UNINSTALL.sh`. **No deck is
  edited** — a substitute simply answers to the name the slides ask for. The
  faces end up embedded in the exported PDFs, so output stays right after the
  substitutes are removed.

  Reporting and alias resolution run on system Python. `--install` needs
  fontTools, and says so, because building a weight out of a variable font
  means instancing it and nothing in the standard library does that.

- **A truncated run no longer reports success.** A deck whose `media/` held
  only `media.json` — linked video, nothing extractable — made
  `grep -cv`'s own `0` collide with a `|| echo 0` fallback, so the arithmetic
  received `0\n0` and failed. That ended the loop at deck 21 of 156, and the
  run then printed `done: 21 exported, 0 skipped, 0 failed` and exited 0. The
  arithmetic is fixed, and separately the run now counts the decks it actually
  reached against the total, says `INCOMPLETE: n of m deck(s) were never
  reached`, and exits non-zero. The false success was the dangerous half: a
  script or an agent believes an exit code.

- **A deck that will not export with `--states` is retried without it.** The
  build-state expansion rewrites the deck, and that rewrite is itself something
  PowerPoint can refuse to open. A 349 MB, 125-slide deck failed four times as
  the expanded copy — twice at 240s, twice at 2400s — then exported from the
  untouched original in under a minute. More time never helped because nothing
  was happening. The cost of the fallback is the build pages; the alternative
  was no PDF at all. `slides.json` is rewritten as a plain slide map when this
  happens, so page numbers stay right.

- **A .pptx with no slides is skipped, not failed.** A template saved with the
  wrong extension is a valid package holding layouts, masters and themes and
  nothing else. It used to be attempted, fail, and be blamed on deck size — the
  file in question was 239 KB. It is now recognised up front and reported as a
  skip, so a library containing one can still report a clean run.

- **Failure messages say what was tried instead of guessing why.** The old text
  named very large decks as "the usual cause" unconditionally. In this library
  the 753 MB and 299 MB decks both exported fine and the 349 MB one did not, so
  size was actively misleading. The message now lists the recovery steps that
  ran — reset and retry, embedded fonts removed, `--states` dropped — and says
  that a deck failing this way sits at 0% CPU, so a longer `--deadline` will
  not help.

- **`audit` reports hidden slides.** PowerPoint never exports them, so the PDF
  comes out shorter than the slide count and the gap looks like dropped work.
  Thirty decks in this library held 186 hidden slides between them, and every
  page shortfall in the finished archive turned out to be exactly that — but
  only after counting them by hand to prove nothing had been lost.

- **`FONT-SUBSTITUTIONS.txt` is written in `beside` layout too.** It was gated
  on there being an `--out`, so the one layout with no central output folder
  was also the one with no durable record — precisely the run that needs it.
  It now lands at the source root.

- **The run log survives `--layout beside`.** With no `--out` the log fell back
  to the workspace, which is deleted on exit, taking the record of a multi-hour
  run with it. It now defaults to `~/Library/Logs/pptxtractor/<date>.log` —
  durable, and still outside the source tree, which the log must stay out of
  because it records deck paths.

- **`caffeinate` is applied, not advised.** The banner told the user to prefix
  the command with `caffeinate -i` and then did not do it; a run long enough to
  need that advice is long enough that forgetting it loses the run. It now
  re-execs itself under `caffeinate` once. `--no-caffeinate` opts out, and a
  dry run does not bother.

- **`total_decks` no longer counts `._` resource forks**, which the loop skips.
  On a synced volume that made the denominator one too high.

- **A run that had to strip fonts now says so in the archive, not just in the
  scrollback.** `FONT-SUBSTITUTIONS.txt` lands next to the exports, naming the
  decks affected and both ways to get them at full fidelity. By the time a long
  run finishes, the lines explaining a deck from an hour ago are thousands of
  lines back.

- **Decks with restricted embedded fonts no longer hang the export.** A deck can
  embed fonts licensed for preview-and-print only — most licensed families are,
  Graphik among them — and PowerPoint refuses to open one for editing without a
  modal prompt: *"This presentation cannot be edited because it contains one or
  more read-only embedded (restricted) fonts."* Driven by AppleScript with
  nobody at the keyboard, that dialog is never answered, so the export sat there
  until the watchdog killed it and the deck was reported as a timeout. No
  deadline is long enough, because nothing is happening. After the normal
  attempts fail, the export now retries with the embedded fonts removed and says
  so — those faces fall back to whatever is installed, which is worse than a
  perfect export and much better than none. A 64-slide deck that had failed
  twice completed in 10 seconds this way.

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

- **A run that had to strip fonts now says so in the archive, not just in the
  scrollback.** `FONT-SUBSTITUTIONS.txt` lands next to the exports, naming the
  decks affected and both ways to get them at full fidelity. By the time a long
  run finishes, the lines explaining a deck from an hour ago are thousands of
  lines back.

- **Decks with restricted embedded fonts no longer hang the export.** A deck can
  embed fonts licensed for preview-and-print only — most licensed families are,
  Graphik among them — and PowerPoint refuses to open one for editing without a
  modal prompt: *"This presentation cannot be edited because it contains one or
  more read-only embedded (restricted) fonts."* Driven by AppleScript with
  nobody at the keyboard, that dialog is never answered, so the export sat there
  until the watchdog killed it and the deck was reported as a timeout. No
  deadline is long enough, because nothing is happening. After the normal
  attempts fail, the export now retries with the embedded fonts removed and says
  so — those faces fall back to whatever is installed, which is worse than a
  perfect export and much better than none. A 64-slide deck that had failed
  twice completed in 10 seconds this way.

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
