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
  # Look for the format that was asked for. Testing "is the folder non-empty"
  # meant `render --format jpeg` silently did nothing on an archive that had
  # already been rendered to PNG.
  case "$FORMAT" in jpeg|jpg) ext="jpg" ;; *) ext="png" ;; esac
  if [ "$FORCE" = "0" ] && [ -n "$(ls -A "$out"/*."$ext" 2>/dev/null)" ]; then
    skip=$((skip+1)); continue
  fi
  mkdir -p "$out"
  # Count what the renderer actually wrote - it prints one line per file.
  # Counting the whole folder reported 47 pages for a 23-page deck once both
  # PNG and JPEG lived in it.
  # Name images for the slide they show, not the page they happen to be.
  labels="$(mktemp)"
  python3 "$HERE/name_images.py" "$dir/slides.json" "$labels" 2>/dev/null || : > "$labels"
  if out_lines=$("$RENDER" "$pdf" "$out" "$EDGE" "$stem" "$FORMAT" "${PAGE:-}" "$labels" 2>/dev/null); then
    rm -f "$labels"
    n=$(printf '%s\n' "$out_lines" | grep -c .)
    echo "$n $FORMAT  $stem"
    done_n=$((done_n+1)); pages=$((pages+n))
  else
    rm -f "$labels"
    echo "FAILED  $pdf"
  fi
done < <(find "$ARCHIVE" -type f -name "*.pdf" -print0 | sort -z)

echo
echo "done: $done_n decks rendered ($pages pages), $skip already had images   (edge=$EDGE, format=$FORMAT)"
