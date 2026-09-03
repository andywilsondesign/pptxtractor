# Notes for agents

`pptxtractor` is a macOS CLI. Three commands: `audit`, `export`, `render`.
Full usage in [README.md](README.md); the workflow an agent should follow is in
[skills/pptxtractor/SKILL.md](skills/pptxtractor/SKILL.md).

**This file is the single source of truth for how the tool behaves.** `CLAUDE.md`
is a symlink to it, and `SKILL.md` points here rather than restating any of it,
so guidance cannot drift between agents. If you are adding a rule about the
tool's behaviour, add it here — not in a copy.

## Contract

Exit status is meaningful — check it rather than parsing stdout:

| code | meaning |
|-----:|---------|
| 0 | everything asked for was produced |
| 1 | one or more decks failed; the rest are on disk and the run resumes |
| 2 | bad arguments |
| 3 | PowerPoint missing, or older than 2016 |
| 4 | not allowed to control PowerPoint (`-1743`) — needs a human, once |
| 5 | PowerPoint did not answer; it is sitting on a dialog |
| 6 | another export already holds the lock |

Codes 3-6 all need a person at the keyboard. Do not retry them in a loop —
report them and say what the user has to do.

- `audit <folder> [--json FILE]` — JSON on stdout (or to FILE), human summary on
  stderr. Read-only, no PowerPoint. Exit 0 on success, 2 on a bad argument.
- `export <folder|deck.pptx> [--out <folder>]` — takes a whole tree or one
  deck. Idempotent: a deck that already has an export is skipped, so a stopped
  run resumes cleanly. Per-deck progress on stdout.
  - Each deck becomes one `<name> (export)/` folder holding its PDF, media,
    images and page map.
  - `--layout mirror|beside|flat` — mirror (default) recreates the source
    folders under `--out`; beside writes next to each deck; flat uses one folder.
  - `--on-conflict ask|skip|overwrite|new` — what to do when a deck already has
    an export. **Without a terminal this is always `skip`**, never `ask`, so a
    non-interactive run can never block waiting for input. Pass it explicitly
    when the user has asked to redo work.
  - **Pass `--out`.** Without it the archive lands in
    `~/Documents/pptxtractor/<today>/`. That is a sensible default for a person,
    but you should choose the destination deliberately and tell the user where
    it went.
- `render <archive>` — idempotent; skips decks that already have images unless
  `--force`.

## The sandbox, which explains most failures

PowerPoint on macOS is sandboxed: `codesign -d --entitlements` shows
`com.apple.security.app-sandbox` and `files.user-selected.read-write`. It can
only open files the **user** picked in a dialog, or files inside its own
container.

So if you hand it a path it has not been granted, it raises a modal **"Grant
File Access"** dialog and waits — indefinitely, at 0% CPU, with no error and
often with no visible window if it is behind something. A scratch directory from
`mktemp` is a new path every run, so a previous grant never applies and the
dialog returns for every deck.

`export` therefore does its work inside
`~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/`, where no grant is
needed. Decks are copied in and the PDF copied back out by the shell, which is
not sandboxed. **Do not change the workspace to somewhere outside the
container** — that reintroduces the dialog and every export will appear to hang.

Error `-9074` from an `open` command is consistent with the same cause: the
sandbox refusing a path.

## Before you drive `export`

`export --dry-run` needs no PowerPoint, so use it to plan a run and to confirm
paths on a machine that may not have PowerPoint at all. A real run preflights
for you: PowerPoint present, new enough, scriptable, answering, and no other
export in progress. Each of those failures exits with its own code and a message
written for the user rather than for you.

## Hidden slides

PowerPoint does not put `show="0"` slides into a PDF. Anything that maps pages
back to slides — the build expansion, `slides.json`, the notes mapping — must
skip them, or the map runs ahead of the PDF and notes land on the wrong pages.
This fails silently: the PDF still looks right. `buildstates.is_hidden()` is the
single test; use it rather than writing another one.

## Very large decks

A 333 MB deck opened and then sat at 0% CPU indefinitely, with the file open, no
dialog, and PowerPoint unresponsive even to Accessibility queries. The
`--deadline` watchdog exists for this. Run `audit` first and treat multi-hundred
-MB decks as likely to need handling on their own.

## Do not kill an export in flight

If an export is interrupted — killing the driving `osascript`, or force-quitting
PowerPoint — the app can come back wedged: no window, a pinned core, and every
subsequent open failing. Only `pkill -9` and a relaunch clears that; a graceful
quit will not.

The exporter already handles failures itself: a per-deck `--deadline` watchdog,
a PowerPoint reset on any failure, one retry, and a stop after three consecutive
failures. Let it do that rather than intervening.

## Long runs

A large library takes tens of minutes. Run it in the background and report
progress; do not block a turn on it, and do not impose a timeout shorter than
the deck needs — that is how the wedge happens.
