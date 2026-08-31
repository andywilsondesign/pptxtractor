# Changelog

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
