#!/bin/bash
# Destination-picking tests. None of this needs PowerPoint: --dry-run resolves
# every path and writes nothing, so the parts most likely to scatter output
# across a disk are covered on any machine, CI included.
set -uo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(dirname "$HERE")"
CLI="$ROOT/bin/pptxtractor"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

fails=0
check() {   # check <label> <got> <want>
  if [ "$2" = "$3" ]; then
    echo "  ok   $1"
  else
    echo "  FAIL $1"; echo "         got  $2"; echo "         want $3"; fails=$((fails+1))
  fi
}

# a deck that is a real .pptx, cheaply: the fixture builder makes one
python3 -c "
import sys; sys.path.insert(0, '$HERE'); sys.path.insert(0, '$ROOT/src')
import make_fixture; make_fixture.build_pptx('$TMP/seed.pptx')" \
  || { echo 'could not build the fixture deck'; exit 1; }

dests() {   # dests <args...>  -> the destination folders, one per line, sorted
  "$CLI" export "$@" --dry-run 2>/dev/null \
    | grep -E '^ +-> ' | sed 's|^ *-> *||; s|/$||' | sort
}

echo "trailing slashes"
mkdir -p "$TMP/src/2024/Acme"
cp "$TMP/seed.pptx" "$TMP/src/2024/Acme/Deck.pptx"
want="$TMP/out/2024/Acme/Deck (export)"
# A tab-completed folder carries a trailing slash. It used to defeat the
# prefix-strip that makes paths relative, and the whole absolute source path
# was recreated underneath --out.
check "root without slash"  "$(dests "$TMP/src"  --out "$TMP/out")"  "$want"
check "root with slash"     "$(dests "$TMP/src/" --out "$TMP/out")"  "$want"
check "out with slash"      "$(dests "$TMP/src"  --out "$TMP/out/")" "$want"
check "both with slashes"   "$(dests "$TMP/src/" --out "$TMP/out/")" "$want"

echo "layouts"
check "mirror keeps the tree" \
  "$(dests "$TMP/src" --out "$TMP/out" --layout mirror)" \
  "$TMP/out/2024/Acme/Deck (export)"
check "flat collapses it" \
  "$(dests "$TMP/src" --out "$TMP/out" --layout flat)" \
  "$TMP/out/Deck (export)"
check "beside writes next to the deck" \
  "$(dests "$TMP/src" --layout beside)" \
  "$TMP/src/2024/Acme/Deck (export)"

echo "single deck"
check "a lone .pptx is a valid source" \
  "$(dests "$TMP/src/2024/Acme/Deck.pptx" --out "$TMP/out")" \
  "$TMP/out/Deck (export)"

# Flattening a tree collapses folders that were keeping same-named decks apart.
# Ten distinct decks called "Personas Template.pptx" - one per cohort - would
# otherwise all claim one folder, and "already exported" would drop nine of
# them while reporting success.
echo "name clashes when flattened"
mkdir -p "$TMP/cohorts/Kiwi" "$TMP/cohorts/Shogito" "$TMP/cohorts/BizGees"
for c in Kiwi Shogito BizGees; do cp "$TMP/seed.pptx" "$TMP/cohorts/$c/Personas.pptx"; done
got="$(dests "$TMP/cohorts" --out "$TMP/flat" --layout flat)"
check "every deck gets its own folder" "$(printf '%s\n' "$got" | wc -l | tr -d ' ')" "3"
check "named after the folder that tells them apart" "$got" \
"$TMP/flat/Personas (BizGees) (export)
$TMP/flat/Personas (Kiwi) (export)
$TMP/flat/Personas (Shogito) (export)"

echo "dry runs write nothing"
before="$(find "$TMP" | wc -l | tr -d ' ')"
dests "$TMP/src" --out "$TMP/never-created" >/dev/null
after="$(find "$TMP" | wc -l | tr -d ' ')"
check "no directories created" "$before" "$after"

echo
if [ "$fails" -eq 0 ]; then echo "all path checks passed"; else echo "$fails FAILED"; fi
exit $([ "$fails" -eq 0 ] && echo 0 || echo 1)
