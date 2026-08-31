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
DEADLINE=600      # seconds per deck before we give up and move on
HERE="$(cd "$(dirname "$0")" && pwd)"
RENDER="$HERE/../bin/pdfrender"
WORK="$(mktemp -d /tmp/slideexport.XXXXXX)"
LOG="$HERE/export.log"

while [ $# -gt 0 ]; do
  case "$1" in
    --run) MODE="run"; shift ;;          # accepted, and the default
    --dry-run) MODE="dry"; shift ;;
    --edge) EDGE="$2"; shift 2 ;;
    --out) OUTROOT="$2"; shift 2 ;;
    --only) ONLY="$2"; shift 2 ;;
    --root) ROOT="$2"; shift 2 ;;
    --states) STATES=1; shift ;;
    --png) PNG=1; shift ;;
    --format) FORMAT="$2"; PNG=1; shift 2 ;;
    --no-media) NOMEDIA=1; shift ;;
    --no-notes) NOTES=0; shift ;;
    --svg) SVG=1; shift ;;
    --deadline) DEADLINE="$2"; shift 2 ;;
    *) echo "unknown arg: $1"; exit 2 ;;
  esac
done

if [ -z "$ROOT" ]; then
  echo "No source folder given. Use:  pptxtractor export --root <folder> --out <folder>"
  exit 2
fi
[ -d "$ROOT" ] || { echo "Not a folder: $ROOT"; exit 2; }
case "$FORMAT" in png|jpeg|jpg) ;; *) echo "--format must be png, jpeg or jpg"; exit 2 ;; esac
[ "$PNG" = "1" ] && [ ! -x "$RENDER" ] && { echo "Renderer missing. Run: make"; exit 1; }
trap 'rm -rf "$WORK"' EXIT

# Force-quitting PowerPoint mid-export wedges it: it comes back with no window,
# a pinned core, and every open failing with -9074 until killed with -9.
ppt_probe() {
  osascript 2>/dev/null <<'OSA'
tell application "Microsoft PowerPoint"
  with timeout of 120 seconds
    try
      set n to count of presentations
      return "ok"
    on error e number num
      return "err " & num
    end try
  end timeout
end tell
OSA
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
  osascript -e 'tell application "Microsoft PowerPoint" to quit saving no' >/dev/null 2>&1
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
    n="$(osascript -e 'tell application "Microsoft PowerPoint" to with timeout of 60 seconds
return (count of presentations)
end timeout' 2>/dev/null)"
    if [ -n "$n" ] && [ "$n" != "0" ]; then
      echo "PowerPoint has $n document(s) open, and this script closes documents as it works."
      echo "Save and close them, then re-run."
      exit 1
    fi
  fi
  echo "Starting from a clean PowerPoint..."
  ppt_reset
  ppt_probe >/dev/null
fi

ok=0; skip=0; fail=0; consecutive=0
while IFS= read -r -d '' src; do
  base="$(basename "$src")"
  case "$base" in ~\$*|._*) continue ;; esac
  [ -n "$ONLY" ] && case "$src" in *"$ONLY"*) ;; *) continue ;; esac

  stem="${base%.*}"
  if [ -n "$OUTROOT" ]; then
    rel="${src#$ROOT/}"; sub="$(dirname "$rel")"
    [ "$sub" = "." ] && dest="$OUTROOT/$stem" || dest="$OUTROOT/$sub/$stem"
  else
    dest="$(dirname "$src")/$stem"
  fi

  if [ -s "$dest/$stem.pdf" ]; then
    echo "SKIP (already exported)  $dest"; skip=$((skip+1)); continue
  fi
  if [ "$MODE" = "dry" ]; then
    echo "WOULD EXPORT  $src"; echo "           ->  $dest/"; ok=$((ok+1)); continue
  fi

  echo "EXPORT  $src"
  mkdir -p "$dest"
  tmp="$WORK/deck.pptx"; feed="$tmp"; pdf="$WORK/deck.pdf"
  rm -f "$WORK"/*.pptx "$pdf"
  cp "$src" "$tmp" || { echo "  copy failed"; fail=$((fail+1)); continue; }

  # animation builds become extra pages, in place, in one pass
  if [ "$STATES" = "1" ]; then
    if python3 "$HERE/buildstates.py" expand "$tmp" "$WORK/expanded.pptx" "$dest/slides.json" 2>>"$LOG"; then
      [ -s "$WORK/expanded.pptx" ] && feed="$WORK/expanded.pptx"
    fi
    grep -q '"expanded_slides": \[\]' "$dest/slides.json" 2>/dev/null && rm -f "$dest/slides.json"
    [ -s "$dest/slides.json" ] || rm -f "$dest/slides.json"
  fi

  # A -9074 means PowerPoint is wedged, and it stays wedged for every deck after.
  # Counting presentations still succeeds in that state, so the only reliable
  # response is to reset unconditionally and try the deck once more.
  if ! to_pdf "$feed" "$pdf"; then
    echo "  export failed - resetting PowerPoint and retrying once"
    ppt_reset
    if ! to_pdf "$feed" "$pdf"; then
      echo "  PDF export failed (see $LOG)"
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
    if "$RENDER" "$pdf" "$dest/images" "$EDGE" "$stem" "$FORMAT" >>"$LOG" 2>&1; then
      echo "  -> $(ls -1 "$dest/images"/* 2>/dev/null | wc -l | tr -d ' ') $FORMAT @ ${EDGE}px"
    fi
  fi
  ok=$((ok+1))
done < <(find "$ROOT" -type f -iname "*.pptx" -print0 | sort -z)

echo
echo "done: $ok exported, $skip skipped, $fail failed   (mode=$MODE)"
