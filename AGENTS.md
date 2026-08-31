# Notes for agents

`pptxtractor` is a macOS CLI. Three commands: `audit`, `export`, `render`.
Full usage in [README.md](README.md); the workflow an agent should follow is in
[skills/pptxtractor/SKILL.md](skills/pptxtractor/SKILL.md).

## Contract

- `audit <folder> [--json FILE]` — JSON on stdout (or to FILE), human summary on
  stderr. Read-only, no PowerPoint. Exit 0 on success, 2 on a bad argument.
- `export --root <folder> --out <folder>` — idempotent. Decks that already have
  a PDF are skipped, so a stopped run resumes cleanly. Per-deck progress on
  stdout.
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
