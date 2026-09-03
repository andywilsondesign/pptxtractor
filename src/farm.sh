#!/bin/bash
# farm.sh - the tractor.
#
# pptxtractor spends most of its life churning slowly through a big field of
# files, which is a dull thing to watch. This is the part that makes it less
# dull: a tractor, some furrows, and a yield at the end.
#
# Two rules, so the joke never gets in the way of the tool:
#   * none of it writes to stdout. Progress goes to stderr, so piping the real
#     output somewhere useful is unaffected.
#   * none of it runs unless a person is watching. No terminal, no tractor.
#
# Source it, then call farm_banner / farm_field / farm_yield.

# A person is watching if stderr is a terminal. NO_COLOR and PPTXTRACTOR_PLAIN
# both turn the whole thing off, for anyone who would rather it did not.
farm_wants_pictures() {
  [ -n "${FARM_FORCE:-}" ] && return 0     # `pptxtractor plough` asked for it
  [ -t 2 ] || return 1
  [ -n "${NO_COLOR:-}" ] && return 1
  [ -n "${PPTXTRACTOR_PLAIN:-}" ] && return 1
  return 0
}

# Not every terminal has a tractor in its font.
farm_tractor_glyph() {
  case "${LC_ALL:-${LC_CTYPE:-${LANG:-}}}" in
    *UTF-8*|*utf8*|*UTF8*) printf '\xf0\x9f\x9a\x9c' ;;   # a tractor
    *) printf '[T]' ;;
  esac
}

farm_banner() {
  farm_wants_pictures || return 0
  cat >&2 <<'TRACTOR'

           ||
          _||_______
         |  ______  |
         | |      | |          p p t x t r a c t o r
       __|_|______|_|__
      |                 \      ploughing your slides
      |  .---.      .-.  |     for everything worth keeping
      '-(  @  )----( o )-'
         '---'      '-'
    ....................................................

TRACTOR
}

# farm_field <done> <total> [label]
# One line, redrawn in place: furrows behind the tractor, crop still standing
# in front of it.
farm_field() {
  farm_wants_pictures || return 0
  local done_n="$1" total="$2" label="${3:-}" width=34
  [ "$total" -gt 0 ] 2>/dev/null || return 0
  local cut=$(( done_n * width / total ))
  [ "$cut" -gt "$width" ] && cut="$width"

  local ploughed="" standing="" i=0
  while [ $i -lt "$cut" ]; do ploughed="$ploughed~"; i=$((i+1)); done
  i=$cut
  while [ $i -lt "$width" ]; do standing="$standing."; i=$((i+1)); done

  # \r and a clear-to-end, so consecutive updates overwrite rather than pile up
  printf '\r\033[K  %s%s%s  %d/%d' \
    "$ploughed" "$(farm_tractor_glyph)" "$standing" "$done_n" "$total" >&2
  [ -n "$label" ] && printf '  %.28s' "$label" >&2
}

farm_field_end() {
  farm_wants_pictures || return 0
  printf '\r\033[K' >&2
}

# farm_yield <decks> <pages> <images> <media>
# What came off the field. Only the non-zero parts get a mention.
farm_yield() {
  farm_wants_pictures || return 0
  local decks="$1" pages="$2" images="$3" media="$4" line=""
  [ "$decks" -gt 0 ] 2>/dev/null || return 0
  line="  yield: $decks deck"; [ "$decks" = "1" ] || line="${line}s"
  [ "${pages:-0}" -gt 0 ] 2>/dev/null && line="$line, $pages page$([ "$pages" = 1 ] || echo s)"
  [ "${images:-0}" -gt 0 ] 2>/dev/null && line="$line, $images still$([ "$images" = 1 ] || echo s)"
  [ "${media:-0}" -gt 0 ] 2>/dev/null && line="$line, $media media file$([ "$media" = 1 ] || echo s)"
  printf '%s\n' "$line" >&2
}

# The whole thing, played once, for anyone who wants to see it without
# committing a library to it. `pptxtractor plough` calls this.
farm_demo() {
  # Someone typed `plough` on purpose, so draw it whether or not this is a
  # terminal. Everywhere else the tractor stays out of the way.
  FARM_FORCE=1
  local total=24 i=0 names=("Q3 Review" "Brand Refresh" "User Research" \
    "Pitch Deck v4" "Roadmap" "Retro" "Onboarding" "All Hands")
  farm_banner
  while [ $i -le $total ]; do
    farm_field "$i" "$total" "${names[$((i % 8))]}"
    sleep 0.12
    i=$((i+1))
  done
  farm_field_end
  printf '  ploughed 24 decks\n' >&2
  farm_yield 24 612 612 37
}
