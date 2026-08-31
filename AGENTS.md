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

## The one rule

**Do not kill an export in flight.** Killing the `osascript` that drives
PowerPoint wedges the application: it relaunches with no window, one core
pinned, and every subsequent `open` failing — sometimes with error `-9074`
immediately, sometimes by hanging indefinitely with no error at all. Counting
open presentations still succeeds in that state, so a naive health probe reports
all clear and then every deck fails.

The exporter already handles this: a per-deck `--deadline` watchdog, an
unconditional PowerPoint reset on any failure, one retry, and a stop after three
consecutive failures. Let it do that. If you interrupt it, you manufacture the
failure you then have to diagnose.

If PowerPoint is already wedged when you start, only `pkill -9` and a relaunch
clears it — a graceful quit will not.

## Long runs

A large library takes tens of minutes. Run it in the background and report
progress; do not block a turn on it, and do not impose a timeout shorter than
the deck needs — that is how the wedge happens.
