#!/bin/bash
# render - produce PNG or JPEG from PDFs an export already made.
#
# PowerPoint is not involved: this reads the vector PDFs, so it is fast,
# repeatable, and keeps working long after the Office trial lapses.
#
#   pptxtractor render <archive>
#   pptxtractor render <archive> --format jpeg --edge 2560
#   pptxtractor render <archive> --only "Q3 Review" --page 12 --force

set -uo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
RENDER="$HERE/../bin/pdfrender"
EDGE=3840
FORMAT="png"
ONLY=""
PAGE=""
FORCE=0

ARCHIVE="${1:-}"
[ -n "$ARCHIVE" ] || { echo "usage: render-pngs.sh <archive dir> [--edge N] [--only TEXT] [--page N] [--force]"; exit 2; }
shift
while [ $# -gt 0 ]; do
  case "$1" in
    --edge) EDGE="$2"; shift 2 ;;
    --format) FORMAT="$2"; shift 2 ;;
    --only) ONLY="$2"; shift 2 ;;
    --page) PAGE="$2"; shift 2 ;;
    --force) FORCE=1; shift ;;
    *) echo "unknown arg: $1"; exit 2 ;;
  esac
done

[ -x "$RENDER" ] || { echo "Renderer missing. Run: make"; exit 1; }
case "$FORMAT" in png|jpeg|jpg) ;; *) echo "--format must be png, jpeg or jpg"; exit 2 ;; esac

done_n=0; skip=0; pages=0
while IFS= read -r -d '' pdf; do
  [ -n "$ONLY" ] && case "$pdf" in *"$ONLY"*) ;; *) continue ;; esac
  dir="$(dirname "$pdf")"
  stem="$(basename "$pdf" .pdf)"
  out="$dir/images"
  if [ "$FORCE" = "0" ] && [ -d "$out" ] && [ -n "$(ls -A "$out" 2>/dev/null)" ]; then
    skip=$((skip+1)); continue
  fi
  mkdir -p "$out"
  if "$RENDER" "$pdf" "$out" "$EDGE" "$stem" "$FORMAT" ${PAGE:+$PAGE} >/dev/null 2>&1; then
    n=$(ls -1 "$out"/* 2>/dev/null | wc -l | tr -d ' ')
    echo "$n $FORMAT  $stem"
    done_n=$((done_n+1)); pages=$((pages+n))
  else
    echo "FAILED  $pdf"
  fi
done < <(find "$ARCHIVE" -type f -name "*.pdf" -print0 | sort -z)

echo
echo "done: $done_n decks rendered ($pages pages), $skip already had images   (edge=$EDGE, format=$FORMAT)"
