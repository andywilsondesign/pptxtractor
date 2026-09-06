#!/usr/bin/env python3
"""
Write a copy of a deck with its embedded fonts removed.

    python3 strip_fonts.py <in.pptx> <out.pptx>

Exit 0 if fonts were removed, 1 if there were none to remove.

This exists for one situation. A deck can embed fonts whose licence only allows
"preview and print" - Graphik and most other licensed families do - and
PowerPoint will not open such a deck for editing without asking first:

    This presentation cannot be edited because it contains one or more
    read-only embedded (restricted) fonts.  [Open Read-only] [Remove ...]

That is a modal dialog. Driven by AppleScript with nobody at the keyboard it
never gets answered, so the export sits there until the watchdog kills it and
the deck is reported as a timeout. Nothing about the deck is wrong, and no
deadline is long enough.

Removing the embedded fonts removes the question. The cost is that those faces
now fall back to whatever is installed, so a deck may substitute where it would
not have. That is worse than a perfect export and much better than no export at
all, which is the alternative - and it is only ever used as a retry, after the
faithful attempt has already failed.
"""
import os, re, shutil, sys, zipfile

FONT_PART = re.compile(r'^ppt/fonts/.+\.fntdata$', re.I)


def has_embedded_fonts(path):
    try:
        with zipfile.ZipFile(path) as z:
            return any(FONT_PART.match(n) for n in z.namelist())
    except (zipfile.BadZipFile, OSError):
        return False


def strip(src, dst):
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        fonts = [n for n in names if FONT_PART.match(n)]
        if not fonts:
            return 0
        # Which relationship ids point at a font part - those go too, or the
        # package refers to something that is no longer there.
        rels_name = 'ppt/_rels/presentation.xml.rels'
        drop_ids = set()
        if rels_name in names:
            rels = z.read(rels_name).decode('utf8', 'ignore')
            for m in re.finditer(r'<Relationship\b[^>]*>', rels):
                tag = m.group(0)
                t = re.search(r'Target="([^"]+)"', tag)
                i = re.search(r'Id="([^"]+)"', tag)
                if t and i and 'fonts/' in t.group(1):
                    drop_ids.add(i.group(1))

        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as w:
            for item in z.infolist():
                n = item.filename
                if FONT_PART.match(n):
                    continue
                data = z.read(n)
                if n == 'ppt/presentation.xml':
                    x = data.decode('utf8', 'ignore')
                    x = re.sub(r'<p:embeddedFontLst>.*?</p:embeddedFontLst>', '', x, flags=re.S)
                    x = re.sub(r'<p:embeddedFontLst\s*/>', '', x)
                    data = x.encode('utf8')
                elif n == rels_name and drop_ids:
                    x = data.decode('utf8', 'ignore')
                    for rid in drop_ids:
                        x = re.sub(r'<Relationship\b[^>]*Id="%s"[^>]*/>' % re.escape(rid), '', x)
                    data = x.encode('utf8')
                w.writestr(item, data)
    return len(fonts)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__.strip()); sys.exit(2)
    n = strip(sys.argv[1], sys.argv[2])
    if n:
        print("  removed %d embedded font%s" % (n, "" if n == 1 else "s"))
    sys.exit(0 if n else 1)
