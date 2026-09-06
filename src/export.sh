#!/bin/bash
# export - archive a PowerPoint library as vector PDFs plus extracted media.
#
# Each deck becomes a folder:
#
#   <Deck name>/
#     <Deck name>.pdf     every slide, vector; animated slides followed by one
#                         page per build state (with --states); speaker notes
#                         attached as annotations
#     slides.json         page -> slide/state map, when builds were expanded
#     media/              animated GIFs, video, audio + media.json manifest
#     images/             only with --png / --format
#
#   pptxtractor export --root ~/Decks --out /Volumes/Archive --states
#   pptxtractor export --root ~/Decks --out /Volumes/Archive --format png --edge 3840
#   pptxtractor export --root ~/Decks --out /Volumes/Archive --only "Q3" --dry-run
#
# Originals are only ever read: every deck is copied to a temp workspace before
# PowerPoint opens it, so no ~$lock files or mtime changes reach the source.

set -uo pipefail

ROOT=""
LAYOUT="mirror"        # mirror | beside | flat  - see resolve_dest()
ON_CONFLICT=""         # ask | skip | overwrite | new; resolved from the tty
DEFAULTED_OUT=0
SINGLE_FILE=0
EDGE=3840
FORMAT="png"
MODE="run"
OUTROOT=""
ONLY=""
STATES=0
PNG=0
NOMEDIA=0
SVG=0
NOTES=1
# Seconds per deck before we give up and move on. Measured on real decks with
# the sandbox fix in place: a typical deck exports in ~15s and the slowest
# success seen - 95 slides, 116 pages, notes, media - took 34s. The cases that
# run long do not run long, they hang: a 333 MB deck sat at 0% CPU past 900s,
# twice. So the useful setting is a few times the slowest real export, not a
# fraction of the longest hang. A wedged deck now costs 8 minutes including the
# retry rather than 20. Raise it with --deadline for unusually heavy libraries.
DEADLINE=240
HERE="$(cd "$(dirname "$0")" && pwd)"
RENDER="$HERE/../bin/pdfrender"
WORK=""
. "$HERE/farm.sh"

# --------------------------------------------------------------------------
# PowerPoint has to be here, and it has to be scriptable. Both are worth
# checking before we start walking a library, because both fail in ways that
# look like the tool is broken rather than the machine being unready.
# --------------------------------------------------------------------------
PPT_BUNDLE_ID=""

require_powerpoint() {
  PPT_BUNDLE_ID="$(osascript -e 'id of application "Microsoft PowerPoint"' 2>/dev/null)"
  if [ -z "$PPT_BUNDLE_ID" ]; then
    cat >&2 <<'MISSING'
Microsoft PowerPoint is not installed, and `export` cannot run without it.

This tool drives the real PowerPoint app to render your slides. That is the
whole point: it is what keeps licensed fonts, gradients, masters and layouts
exactly as they look on screen, which every generic .pptx converter gets wrong.
Nothing else on the machine can stand in for it.

  Get PowerPoint for Mac (2016 or newer, including Microsoft 365):
  https://www.microsoft.com/en-us/microsoft-365/powerpoint

You only need it for `export`. These two work right now, with no PowerPoint:

  pptxtractor audit  <folder>     inspect a library, write JSON
  pptxtractor render <archive>    images from PDFs you already exported

MISSING
    exit 3
  fi

  local app ver major
  app="$(osascript -e 'POSIX path of (path to application "Microsoft PowerPoint")' 2>/dev/null)"
  ver="$(defaults read "${app%/}/Contents/Info.plist" CFBundleShortVersionString 2>/dev/null)"
  major="${ver%%.*}"
  # 16.x is Microsoft 365 / Office 2016-2024. 14.x is Office 2011, whose
  # AppleScript dictionary predates the `save as PDF` verb used here and which
  # is not sandboxed, so none of the container handling applies either.
  if [ -n "$major" ] && [ "$major" -lt 15 ] 2>/dev/null; then
    echo "PowerPoint $ver is too old - this needs 2016 or newer (version 15+)." >&2
    echo "Office 2011 uses a different AppleScript dictionary and is unsupported." >&2
    exit 3
  fi
  [ -n "$ver" ] && echo "PowerPoint $ver"
}

# Run an AppleScript with a hard wall-clock cap. A freshly installed PowerPoint
# can sit on a sign-in, "What's New" or activation sheet, and an Apple event
# sent to it never returns - so every probe here needs a way out. Prints the
# script's output; returns 124 if it had to be killed.
osa_bounded() {
  local limit="$1" out rc i=0; shift
  out="$(mktemp)"
  osascript "$@" >"$out" 2>&1 &
  local pid=$!
  while kill -0 "$pid" 2>/dev/null; do
    [ $i -ge "$limit" ] && break
    sleep 1; i=$((i+1))
  done
  if kill -0 "$pid" 2>/dev/null; then
    kill -9 "$pid" 2>/dev/null; wait "$pid" 2>/dev/null
    rm -f "$out"; return 124
  fi
  wait "$pid" 2>/dev/null; rc=$?
  cat "$out"; rm -f "$out"; return $rc
}

# Sending Apple events to another app needs the user's consent, once, per
# calling program. Until it is granted every command fails with -1743 and the
# run looks wedged for no visible reason.
check_automation_permission() {
  local r
  r="$(osa_bounded 90 -e 'tell application "Microsoft PowerPoint" to count of presentations')"
  if [ $? -eq 124 ]; then
    cat >&2 <<'STUCK'

PowerPoint did not answer within 90 seconds.

It is usually sitting on a dialog that has to be dealt with once by hand - a
sign-in prompt, an activation or licence notice, or the "What's New" window.
Open PowerPoint, clear whatever it is showing, quit it, then run this again.

STUCK
    exit 5
  fi
  case "$r" in
    *-1743*)
      cat >&2 <<'DENIED'

Not allowed to control PowerPoint (error -1743).

macOS asks for this once, per app that sends the request. Grant it in:
  System Settings > Privacy & Security > Automation
and tick "Microsoft PowerPoint" under whichever app you are running this from
(Terminal, iTerm, your editor). Then run this again.

DENIED
      exit 4 ;;
  esac
}

# PowerPoint is sandboxed: its entitlements are app-sandbox plus
# files.user-selected.read-write, so it can only reach files the *user* picked
# in a dialog, or files inside its own container. A fresh mktemp path is neither,
# which is why a scratch dir in /tmp makes it raise "Grant File Access" on every
# single deck - the path is new every run, so a previous grant never applies.
#
# Working inside its container sidesteps the dialog completely: no grant is
# needed, and nothing has to be clicked. We copy decks in and copy the PDF back
# out with the shell, which is not sandboxed.
# Destinations claimed earlier in this run. Without this a --dry-run cannot see
# a clash at all - nothing has been written yet, so every deck looks free - and
# the user only discovers it mid-run. Tracked in a file rather than an array
# because macOS ships bash 3.2, which has no associative arrays.
claim_dest() { printf '%s\n' "$1" >> "$CLAIMED"; }
dest_claimed() { [ -s "$CLAIMED" ] && grep -Fxq -- "$1" "$CLAIMED"; }

# Provenance. An export folder records which deck made it, so a later run can
# tell "I already did this one" from "a different deck wants this name".
write_source_record() {
  local dest="$1" src="$2"
  printf '{\n  "path": %s,\n  "bytes": %s,\n  "exported": %s\n}\n' \
    "\"$(printf '%s' "$src" | sed 's/\\/\\\\/g; s/"/\\"/g')\"" \
    "$(wc -c < "$src" 2>/dev/null | tr -d ' ')" \
    "\"$(date -u +%Y-%m-%dT%H:%M:%SZ)\"" > "$dest/source.json" 2>/dev/null || true
}

source_of() {
  python3 -c 'import json,sys
try: print(json.load(open(sys.argv[1]))["path"])
except Exception: print("")' "$1/source.json" 2>/dev/null
}

# True when this export folder was made by this same deck. Folders written
# before provenance existed have no record; treat those as a match so an older
# archive still resumes rather than duplicating itself.
same_source() {
  local dest="$1" src="$2" recorded
  [ -f "$dest/source.json" ] || return 0
  recorded="$(source_of "$dest")"
  [ -z "$recorded" ] && return 0
  [ "$recorded" = "$src" ]
}

# Name a colliding deck after the folders that actually tell it apart: the
# parent, then the grandparent, then a number. "Personas Template (Kiwi)".
disambiguate() {
  local dest="$1" src="$2" folder="$3" parent base try n
  base="$(dirname "$dest")"
  parent="$(basename "$(dirname "$src")")"
  for try in "$parent" "$(basename "$(dirname "$(dirname "$src")")")/$parent"; do
    [ -z "$try" ] || [ "$try" = "/" ] && continue
    candidate="$base/${folder% (export)} ($(printf '%s' "$try" | tr '/' '-')) (export)"
    if ! dest_claimed "$candidate" \
       && { [ ! -e "$candidate" ] || same_source "$candidate" "$src"; }; then
      printf '%s' "$candidate"; return 0
    fi
  done
  n=2
  while dest_claimed "$base/${folder% (export)} ($n) (export)" \
        || { [ -e "$base/${folder% (export)} ($n) (export)" ] \
             && ! same_source "$base/${folder% (export)} ($n) (export)" "$src"; }; do
    n=$((n+1))
  done
  printf '%s' "$base/${folder% (export)} ($n) (export)"
}

# One PowerPoint, one export at a time. Two runs would close each other's
# documents mid-save - the exporter clears open presentations before each deck -
# and the failures would look random. mkdir is atomic, so it works as a lock.
LOCK=""
take_lock() {
  local lock="${TMPDIR:-/tmp}/pptxtractor.export.lock" owner=""
  if ! mkdir "$lock" 2>/dev/null; then
    [ -f "$lock/pid" ] && owner="$(cat "$lock/pid" 2>/dev/null)"
    if [ -n "$owner" ] && kill -0 "$owner" 2>/dev/null; then
      echo "Another export is already running (pid $owner)." >&2
      echo "Only one can drive PowerPoint at a time. Wait for it, or stop it first." >&2
      exit 6            # LOCK stays empty: we must not clean up someone else's
    fi
    echo "Clearing a stale lock from pid ${owner:-unknown}." >&2
    rm -rf "$lock"; mkdir "$lock" 2>/dev/null || { echo "Could not take the lock." >&2; exit 6; }
  fi
  LOCK="$lock"          # only now do we own it, and only now may the trap remove it
  echo $$ > "$LOCK/pid"
}

CLAIMED=""
setup_workspace() {
  local container=""
  [ -n "$PPT_BUNDLE_ID" ] && container="$HOME/Library/Containers/$PPT_BUNDLE_ID/Data"
  if [ -n "$container" ] && [ -d "$container" ]; then
    WORK="$container/tmp/pptxtractor.$$"
  else
    # No container yet - PowerPoint has never been launched, or this build
    # stores it elsewhere. Say so, because the fallback is the slow, prompt-
    # ridden path rather than a silent equivalent.
    if [ "$MODE" = "run" ]; then
      echo "Note: PowerPoint's container is not at ~/Library/Containers/${PPT_BUNDLE_ID:-com.microsoft.Powerpoint}/." >&2
      echo "      Falling back to a temporary folder; expect a \"Grant File Access\" prompt per deck." >&2
      echo "      Opening PowerPoint once by hand usually creates the container and avoids this." >&2
    fi
    WORK="$(mktemp -d /tmp/pptxtractor.XXXXXX)"
  fi
  # Sweep workspaces left by runs that never got to clean up after themselves.
  # The EXIT trap handles a normal finish and a Ctrl-C, but not a kill -9 or a
  # crash, and these sit inside PowerPoint's container holding copies of decks.
  # Only touch ones whose process is definitely gone - another run may be live.
  if [ "$MODE" = "run" ] && [ -n "$container" ] && [ -d "$container/tmp" ]; then
    for stale in "$container"/tmp/pptxtractor.*; do
      [ -d "$stale" ] || continue
      stale_pid="${stale##*.}"
      case "$stale_pid" in ''|*[!0-9]*) continue ;; esac
      kill -0 "$stale_pid" 2>/dev/null && continue     # still running
      rm -rf "$stale" 2>/dev/null
    done
  fi
  mkdir -p "$WORK"
  CLAIMED="$WORK/claimed.txt"; : > "$CLAIMED"
}

while [ $# -gt 0 ]; do
  case "$1" in
    --run) MODE="run"; shift ;;          # accepted, and the default
    --dry-run) MODE="dry"; shift ;;
    --edge) EDGE="$2"; shift 2 ;;
    --out) OUTROOT="$2"; shift 2 ;;
    --layout) LAYOUT="$2"; shift 2 ;;
    --on-conflict) ON_CONFLICT="$2"; shift 2 ;;
    --only) ONLY="$2"; shift 2 ;;
    --root) ROOT="$2"; shift 2 ;;
    --states) STATES=1; shift ;;
    --png) PNG=1; shift ;;
    --format) FORMAT="$2"; PNG=1; shift 2 ;;
    --no-media) NOMEDIA=1; shift ;;
    --no-notes) NOTES=0; shift ;;
    --svg) SVG=1; shift ;;
    --deadline) DEADLINE="$2"; shift 2 ;;
    # Catch-alls last: `case` takes the first match, so anything below this
    # would never be reached.
    -*) echo "unknown arg: $1"; exit 2 ;;
    # A bare path is the source, so `export ~/Decks` and `export deck.pptx` work.
    *) if [ -z "$ROOT" ]; then ROOT="$1"; shift; else echo "unexpected: $1"; exit 2; fi ;;
  esac
done

if [ -z "$ROOT" ]; then
  echo "No source given. Use:  pptxtractor export <folder|deck.pptx> [--out <folder>]"
  exit 2
fi
# A single deck is as valid a source as a whole tree.
if [ -f "$ROOT" ]; then
  case "$ROOT" in
    *.pptx|*.PPTX) SINGLE_FILE=1 ;;
    *) echo "Not a .pptx file: $ROOT"; exit 2 ;;
  esac
elif [ ! -d "$ROOT" ]; then
  echo "No such file or folder: $ROOT"; exit 2
fi
# Strip trailing slashes before anything uses these as prefixes. "$ROOT/" with
# a trailing slash makes the rel= prefix-strip below miss entirely, and the
# whole absolute source path ends up recreated inside --out. Tab-completing a
# directory adds that slash, so this is the normal way to type it.
while [ "$ROOT" != "/" ] && [ "${ROOT%/}" != "$ROOT" ]; do ROOT="${ROOT%/}"; done
while [ "$OUTROOT" != "/" ] && [ "${OUTROOT%/}" != "$OUTROOT" ]; do OUTROOT="${OUTROOT%/}"; done

case "$LAYOUT" in mirror|beside|flat) ;; *)
  echo "--layout must be mirror, beside or flat"; exit 2 ;; esac
case "${ON_CONFLICT:-unset}" in unset|ask|skip|overwrite|new) ;; *)
  echo "--on-conflict must be ask, skip, overwrite or new"; exit 2 ;; esac

# Where the archive lands, and saying so plainly. Writing somewhere the user did
# not choose is only acceptable if they are told - before, during and after.
if [ "$LAYOUT" = "beside" ]; then
  [ -n "$OUTROOT" ] && { echo "--layout beside writes next to each deck; drop --out."; exit 2; }
elif [ -z "$OUTROOT" ]; then
  OUTROOT="$HOME/Documents/pptxtractor/$(date +%Y-%m-%d)"
  DEFAULTED_OUT=1
fi

# Prompting is only safe when someone is there to answer. Without a terminal the
# defaults are fixed and documented, so an agent or a cron job behaves the same
# way every time.
INTERACTIVE=0
[ -t 0 ] && [ -t 1 ] && INTERACTIVE=1
if [ -z "$ON_CONFLICT" ]; then
  if [ "$INTERACTIVE" = "1" ]; then ON_CONFLICT="ask"; else ON_CONFLICT="skip"; fi
fi
case "$FORMAT" in png|jpeg|jpg) ;; *) echo "--format must be png, jpeg or jpg"; exit 2 ;; esac
[ "$PNG" = "1" ] && [ ! -x "$RENDER" ] && { echo "Renderer missing. Run: make"; exit 1; }

# --dry-run is deliberately usable without PowerPoint: it is how you plan a run
# on a machine that has not got it yet.
if [ "$MODE" = "run" ]; then
  take_lock
  require_powerpoint
  check_automation_permission
fi
setup_workspace

# The log records deck paths and file names, so it must not land in the source
# tree - this repo is public and the decks are not. Keep it beside the archive
# being built; fall back to the (auto-deleted) workspace when there is no --out.
LOG="${PPTXTRACTOR_LOG:-${OUTROOT:+$OUTROOT/pptxtractor.log}}"
LOG="${LOG:-$WORK/pptxtractor.log}"
# A dry run must not create anything, including the folder it would log into.
if [ "$MODE" = "dry" ]; then
  LOG="$WORK/pptxtractor.log"
else
  mkdir -p "$(dirname "$LOG")" 2>/dev/null
fi
trap 'rm -rf "$WORK"; [ -n "$LOCK" ] && rm -rf "$LOCK"' EXIT

# Force-quitting PowerPoint mid-export wedges it: it comes back with no window,
# a pinned core, and every open failing with -9074 until killed with -9.
ppt_probe() {
  local r rc
  r="$(osa_bounded 150 -e 'tell application "Microsoft PowerPoint"
  with timeout of 120 seconds
    try
      set n to count of presentations
      return "ok"
    on error e number num
      return "err " & num
    end try
  end timeout
end tell')"
  rc=$?
  # A wedged PowerPoint does not error, it simply never answers - the
  # AppleScript-level timeout only helps once the app is listening at all.
  [ $rc -eq 124 ] && { echo "no answer"; return 0; }
  printf '%s\n' "$r"
}

ppt_ready() {
  local r i
  r="$(ppt_probe)"
  [ "$r" = "ok" ] && return 0
  echo "  PowerPoint not responding ($r) - restarting it"
  ppt_reset
  r="$(ppt_probe)"
  [ "$r" = "ok" ] || { echo "  PowerPoint still unhealthy ($r) - open it by hand once, then re-run."; return 1; }
  echo "  PowerPoint recovered."
}

# Two timeouts, and both are needed.
#
# `with timeout` inside the AppleScript stops a cold PowerPoint from tripping the
# 120s default AppleEvent limit on the first deck and losing the export silently.
#
# DEADLINE is the outer watchdog. A wedged PowerPoint does not always fail loudly:
# instead of -9074 it can sit at 0% CPU partway through an export, with no dialog
# and no error, and never return. Without this, one such deck stalls the whole
# run. On timeout we kill the script, reset PowerPoint and move on - and the deck
# that hung will usually export in seconds on a later pass.
to_pdf() {
  local src="$1" out="$2" i=0
  osascript >>"$LOG" 2>&1 <<OSA &
tell application "Microsoft PowerPoint"
  with timeout of 3600 seconds
    repeat while (count of presentations) > 0
      close presentation 1 saving no
    end repeat
    open POSIX file "$src"
    set p to active presentation
    save p in POSIX file "$out" as save as PDF
    close p saving no
  end timeout
end tell
OSA
  local osa=$!
  while kill -0 "$osa" 2>/dev/null; do
    [ $i -ge "$DEADLINE" ] && break
    sleep 2; i=$((i+2))
  done
  if kill -0 "$osa" 2>/dev/null; then
    echo "  timed out after ${DEADLINE}s - skipping this deck"
    kill -9 "$osa" 2>/dev/null; wait "$osa" 2>/dev/null
    ppt_reset
    return 1
  fi
  wait "$osa" 2>/dev/null
  [ -s "$out" ]
}

ppt_reset() {
  local i=0
  osa_bounded 20 -e 'tell application "Microsoft PowerPoint" to quit saving no' >/dev/null 2>&1
  while pgrep -x "Microsoft PowerPoint" >/dev/null && [ $i -lt 10 ]; do sleep 1; i=$((i+1)); done
  pkill -9 -x "Microsoft PowerPoint" 2>/dev/null
  i=0; while pgrep -x "Microsoft PowerPoint" >/dev/null && [ $i -lt 10 ]; do sleep 1; i=$((i+1)); done
  find "$WORK" -name '~$*' -delete 2>/dev/null
  for try in 1 2 3; do open -a "Microsoft PowerPoint" 2>/dev/null && break; sleep 5; done
  i=0; while [ $i -lt 25 ]; do sleep 1; i=$((i+1)); done
}

# The export closes whatever is open in PowerPoint, so refuse to start while the
# user has their own documents up.
if [ "$MODE" = "run" ]; then
  if pgrep -x "Microsoft PowerPoint" >/dev/null; then
    n="$(osa_bounded 75 -e 'tell application "Microsoft PowerPoint" to with timeout of 60 seconds
return (count of presentations)
end timeout')"
    # No answer means it is already wedged; ppt_reset below clears that, so do
    # not treat silence as "the user has documents open".
    [ $? -eq 124 ] && n=""
    if [ -n "$n" ] && [ "$n" != "0" ]; then
      echo "PowerPoint has $n document(s) open, and this script closes documents as it works."
      echo "Save and close them, then re-run."
      exit 1
    fi
  fi
  cat <<'BANNER'

This drives PowerPoint itself. While it runs:

  * leave PowerPoint alone - it opens and closes a deck per file, and clicking
    into it, opening your own file or quitting it will break the export
  * do not open your own presentations while this runs. Recovering from a stuck
    deck force-quits PowerPoint, and anything you had open goes with it
  * do not let the machine sleep (prefix the command with `caffeinate -i`)

You can keep using the rest of the Mac. Stopping the run is safe - press Ctrl-C
between decks and re-run later; finished decks are skipped.

BANNER
  farm_banner
  echo "Starting from a clean PowerPoint..."
  ppt_reset
  ppt_probe >/dev/null
fi

case "$LAYOUT" in
  beside) echo "Writing each export beside its deck, under $ROOT" ;;
  flat)   echo "Writing all exports into $OUTROOT" ;;
  mirror) echo "Writing exports to $OUTROOT, mirroring the source folders" ;;
esac
if [ "$DEFAULTED_OUT" = "1" ]; then
  echo "  (no --out given, so this is the default location - pass --out to choose)"
fi
echo

# Count the field before working it, so progress has a denominator. Same
# filters as the loop below, so the number matches what actually gets done.
if [ "$SINGLE_FILE" = "1" ]; then
  total_decks=1
else
  total_decks=$(find "$ROOT" -type f -iname "*.pptx" 2>/dev/null \
    | grep -v '/~\$' \
    | { if [ -n "$ONLY" ]; then grep -F -- "$ONLY"; else cat; fi; } \
    | wc -l | tr -d ' ')
fi

# With --layout flat every deck lands in one folder, so any name used twice in
# the source tree is a clash waiting to happen. Work out which names those are
# before starting, so the *first* deck gets its context in the name too rather
# than only the ones that follow it - otherwise one folder in a set of ten is
# the odd one out and you cannot tell which cohort it came from.
DUPNAMES="$WORK/dupnames.txt"; : > "$DUPNAMES"
if [ "$LAYOUT" = "flat" ] && [ "$SINGLE_FILE" != "1" ]; then
  find "$ROOT" -type f -iname "*.pptx" -not -name '~$*' -not -name '._*' 2>/dev/null \
    | { if [ -n "$ONLY" ]; then grep -F -- "$ONLY"; else cat; fi; } \
    | sed 's|.*/||; s|\.[Pp][Pp][Tt][Xx]$||' \
    | sort | uniq -d > "$DUPNAMES" 2>/dev/null || true
fi
name_repeats() { [ -s "$DUPNAMES" ] && grep -Fxq -- "$1" "$DUPNAMES"; }

ok=0; skip=0; fail=0; consecutive=0; CONFLICT_ALL=""; explained_unreadable=0
seen=0; yield_pages=0; yield_images=0; yield_media=0
while IFS= read -r -d '' src; do
  base="$(basename "$src")"
  case "$base" in ~\$*|._*) continue ;; esac
  [ -n "$ONLY" ] && case "$src" in *"$ONLY"*) ;; *) continue ;; esac

  farm_field_end
  stem="${base%.*}"
  # "(export)" so the folder reads as output at a glance, sitting next to decks
  # in a shared parent. It holds everything for that deck - PDF, images, media,
  # page map - so one deck stays one object however much you ask for.
  folder="$stem (export)"
  case "$LAYOUT" in
    beside) dest="$(dirname "$src")/$folder" ;;
    flat)   dest="$OUTROOT/$folder" ;;
    mirror)
      if [ "$SINGLE_FILE" = "1" ]; then
        dest="$OUTROOT/$folder"
      else
        rel="${src#$ROOT/}"; sub="$(dirname "$rel")"
        [ "$sub" = "." ] && dest="$OUTROOT/$folder" || dest="$OUTROOT/$sub/$folder"
      fi ;;
  esac

  # Two very different things look identical here: the same deck exported
  # before, and a *different* deck that happens to share a name. Flattening a
  # tree makes the second one common - one library here had ten distinct decks
  # all called "Personas Template.pptx", one per cohort. Treating that as
  # "already exported" and skipping would have silently dropped nine of them and
  # called the run a success. So check who wrote the folder before deciding.
  # Claimed earlier in this run means a different deck, full stop - we just
  # processed it. On disk from an earlier run is only a clash if the provenance
  # says a different deck wrote it; a folder with no record predates provenance
  # and is assumed to be this deck resuming.
  clash=0
  # A name the tree uses more than once always gets its context, first one
  # included, so a set of same-named decks reads consistently.
  if name_repeats "$stem"; then
    clash=1
  elif dest_claimed "$dest"; then
    clash=1
  elif [ -e "$dest" ] && [ -n "$(ls -A "$dest" 2>/dev/null)" ] && ! same_source "$dest" "$src"; then
    clash=1
  fi
  if [ "$clash" = "1" ]; then
    # A different deck wants this name. Never a conflict to resolve - just give
    # it a name of its own, borrowed from the folders that distinguish it.
    newdest="$(disambiguate "$dest" "$src" "$folder")"
    if [ -n "$newdest" ]; then
      # Two ways to get here: another deck already wrote this folder, or the
      # source tree simply uses the name more than once and we spotted it up
      # front. Only the first has a previous deck to name.
      previous="$(source_of "$dest")"
      echo "NAME CLASH  $stem"
      if [ -n "$previous" ]; then
        echo "       already exported under this name from"
        echo "         $previous"
      else
        echo "       this name is used by more than one deck in the source"
      fi
      echo "       this one is from $(dirname "$src")"
      echo "       exporting it as: $(basename "$newdest")"
      dest="$newdest"
    fi
  fi

  claim_dest "$dest"

  # An existing export is a decision, not an error. Resolved once per deck, and
  # "!" answers for every deck after it so a long run needs one answer, not 400.
  if [ -e "$dest" ] && [ -n "$(ls -A "$dest" 2>/dev/null)" ]; then
    action="$ON_CONFLICT"
    if [ "$action" = "ask" ]; then
      if [ -n "$CONFLICT_ALL" ]; then
        action="$CONFLICT_ALL"
      else
        echo
        echo "Already exported:  $dest"
        printf '  [s]kip  [o]verwrite  [k]eep both  [q]uit   (add ! for all remaining) '
        read -r reply </dev/tty || reply="s"
        # A terminal can hand back a carriage return, which would stop "s!"
        # matching the all-remaining pattern and re-ask on every deck.
        reply="$(printf '%s' "$reply" | tr -d '[:space:]')"
        case "$reply" in *!) CONFLICT_ALL="${reply%!}" ;; esac
        case "${reply%!}" in
          o|O) action="overwrite" ;;
          k|K) action="new" ;;
          q|Q) echo "Stopped. Nothing further was written."; break ;;
          *)   action="skip" ;;
        esac
        case "$CONFLICT_ALL" in
          o|O) CONFLICT_ALL="overwrite" ;; k|K) CONFLICT_ALL="new" ;;
          s|S) CONFLICT_ALL="skip" ;; *) [ -n "$CONFLICT_ALL" ] && CONFLICT_ALL="skip" ;;
        esac
      fi
    fi
    case "$action" in
      skip)      echo "SKIP (already exported)  $dest"; skip=$((skip+1))
                 seen=$((seen+1)); farm_field "$seen" "$total_decks" "$stem"; continue ;;
      overwrite) echo "REPLACING  $dest"; rm -rf "$dest" ;;
      new)       n=2
                 while [ -e "$dest ($n)" ]; do n=$((n+1)); done
                 dest="$dest ($n)"
                 echo "KEEPING BOTH -> $(basename "$dest")" ;;
    esac
  fi
  # Check the deck is a readable package before PowerPoint ever sees it. A
  # truncated .pptx - a zip whose central directory never arrived - looks fine
  # to `file` and opens as nothing. Handed to PowerPoint it hangs, costing two
  # full deadlines before the run moves on. Reading the directory is instant.
  if ! python3 -c 'import sys,zipfile
try:
    z = zipfile.ZipFile(sys.argv[1])
    sys.exit(0 if "ppt/presentation.xml" in z.namelist() else 3)
except Exception:
    sys.exit(3)' "$src" 2>/dev/null; then
    echo "SKIP (not a readable .pptx)  $src"
    if [ "$explained_unreadable" = "0" ]; then
      echo "       The file itself is damaged - usually a download or sync that"
      echo "       stopped early, leaving a .pptx with no index. PowerPoint cannot"
      echo "       open it either, so there is nothing this tool can do with it."
      echo "       Check for another copy, or your backups."
      explained_unreadable=1
    fi
    fail=$((fail+1))
    seen=$((seen+1)); farm_field "$seen" "$total_decks" "$stem"
    continue
  fi

  if [ "$MODE" = "dry" ]; then
    echo "WOULD EXPORT  $src"; echo "           ->  $dest/"; ok=$((ok+1)); continue
  fi

  echo "EXPORT  $src"
  mkdir -p "$dest"
  write_source_record "$dest" "$src"
  tmp="$WORK/deck.pptx"; feed="$tmp"; pdf="$WORK/deck.pdf"
  rm -f "$WORK"/*.pptx "$pdf"
  cp "$src" "$tmp" || { echo "  copy failed"; fail=$((fail+1)); continue; }

  # animation builds become extra pages, in place, in one pass
  if [ "$STATES" = "1" ]; then
    if python3 "$HERE/buildstates.py" expand "$tmp" "$WORK/expanded.pptx" "$dest/slides.json" 2>>"$LOG"; then
      [ -s "$WORK/expanded.pptx" ] && feed="$WORK/expanded.pptx"
    fi
  fi
  # The page map is written either way. Without it a hidden slide silently
  # shifts every page number after it, and the images would be named for pages
  # rather than for the slides they actually show.
  if [ ! -s "$dest/slides.json" ]; then
    python3 "$HERE/buildstates.py" map "$src" "$dest/slides.json" 2>>"$LOG" || true
  fi

  # A -9074 means PowerPoint is wedged, and it stays wedged for every deck after.
  # Counting presentations still succeeds in that state, so the only reliable
  # response is to reset unconditionally and try the deck once more.
  if ! to_pdf "$feed" "$pdf"; then
    echo "  export failed - resetting PowerPoint and retrying once"
    ppt_reset
    if ! to_pdf "$feed" "$pdf"; then
      echo "  PDF export failed (see $LOG)"
      echo "       PowerPoint could not produce a PDF for this one. Very large"
      echo "       decks are the usual cause - PowerPoint stops responding and"
      echo "       there is no way to make it finish. The deck is fine; try it"
      echo "       on its own, or open it and save a lighter copy."
      rm -f "$dest/slides.json"; rmdir "$dest" 2>/dev/null
      fail=$((fail+1)); consecutive=$((consecutive+1))
      if [ "$consecutive" -ge 3 ]; then
        echo "  three decks failed in a row - stopping. Open PowerPoint by hand, then re-run."
        break
      fi
      continue
    fi
  fi
  consecutive=0
  # Speaker notes ride along as PDF sticky-note annotations. Annotations are not
  # page content, so a viewer lists them in its notes sidebar while the PNG
  # renderer never draws them. Written as an incremental update - the original
  # bytes are untouched and only the annotations are appended.
  if [ "$NOTES" = "1" ]; then
    rm -f "$WORK/notes.json" "$WORK/noted.pdf"
    python3 "$HERE/extract_notes.py" "$src" "$WORK/notes.json" "$dest/slides.json" 2>>"$LOG"
    if [ -s "$WORK/notes.json" ]; then
      if python3 "$HERE/add_notes.py" "$pdf" "$WORK/noted.pdf" "$WORK/notes.json" 2>>"$LOG"; then
        mv "$WORK/noted.pdf" "$pdf"
      else
        echo "  (notes could not be attached - PDF left clean)"
      fi
    fi
  fi

  cp "$pdf" "$dest/$stem.pdf"
  echo "  -> $stem.pdf  ($(du -h "$pdf" | cut -f1))"

  # playable media: animated GIFs, video, audio. Static art is already in the PDF.
  if [ "$NOMEDIA" = "0" ]; then
    python3 "$HERE/extract_media.py" "$src" "$dest/media" $([ "$SVG" = "1" ] && echo --svg) 2>>"$LOG"
    rmdir "$dest/media" 2>/dev/null
  fi

  if [ "$PNG" = "1" ]; then
    mkdir -p "$dest/images"
    labels="$WORK/labels.txt"; rm -f "$labels"
    python3 "$HERE/name_images.py" "$dest/slides.json" "$labels" 2>>"$LOG" || true
    if "$RENDER" "$pdf" "$dest/images" "$EDGE" "$stem" "$FORMAT" "" "$labels" >>"$LOG" 2>&1; then
      echo "  -> $(ls -1 "$dest/images"/* 2>/dev/null | wc -l | tr -d ' ') $FORMAT @ ${EDGE}px"
    fi
  fi
  ok=$((ok+1))
  [ -s "$dest/slides.json" ] && yield_pages=$((yield_pages + $(python3 -c \
    'import json,sys;print(len(json.load(open(sys.argv[1]))["pages"]))' \
    "$dest/slides.json" 2>/dev/null || echo 0)))
  [ -d "$dest/images" ] && yield_images=$((yield_images + $(ls -1 "$dest/images" 2>/dev/null | wc -l | tr -d ' ')))
  [ -d "$dest/media" ] && yield_media=$((yield_media + $(ls -1 "$dest/media" 2>/dev/null | grep -cv '^media.json$' || echo 0)))
  seen=$((seen+1)); farm_field "$seen" "$total_decks" "$stem"
done < <(if [ "$SINGLE_FILE" = "1" ]; then printf '%s\0' "$ROOT"
         else find "$ROOT" -type f -iname "*.pptx" -print0 | sort -z; fi)

farm_field_end
echo
echo "done: $ok exported, $skip skipped, $fail failed   (mode=$MODE)"
farm_yield "$ok" "$yield_pages" "$yield_images" "$yield_media"
# Say where it went once more. Someone who scrolled past the banner, or walked
# away for an hour, should not have to hunt for their own archive.
if [ "$ok" -gt 0 ] || [ "$skip" -gt 0 ]; then
  if [ "$LAYOUT" = "beside" ]; then
    echo "exports are beside each deck, under $ROOT"
  else
    echo "exports are in: $OUTROOT"
    [ "$DEFAULTED_OUT" = "1" ] && echo "  (the default location - use --out next time to put them elsewhere)"
    command -v open >/dev/null && echo "  open it with:  open \"$OUTROOT\""
  fi
fi

# Exit non-zero when decks failed. The run is still resumable and the decks that
# worked are on disk - this reports the outcome rather than changing it - but a
# script or agent driving this needs to see failure without parsing stdout.
# 0 = everything asked for was produced, 1 = some deck failed,
# 2 = bad arguments, 3 = PowerPoint missing or too old, 4 = automation denied.
[ "$fail" -gt 0 ] && exit 1
exit 0
