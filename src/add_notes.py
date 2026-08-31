#!/usr/bin/env python3
"""
Attach speaker notes to a PDF as sticky-note annotations, without touching a
byte of the original.

    python3 add-notes.py <in.pdf> <out.pdf> <notes.json>

Annotations are not page content, so a PDF viewer lists them in its notes
sidebar while CGPDFPage rendering - what pdf2png does - never draws them. Notes
ride along with the archive; PNG exports stay clean.

This writes an *incremental update*: the original file is copied verbatim and
new objects are appended. Re-saving through a PDF library instead would
re-encode every image (PDFKit tripled a 14 MB deck to 44 MB and shifted pixels).
"""
import json, re, sys


def parse_xref_chain(data):
    """objnum -> offset, plus the newest trailer dict text."""
    m = re.search(rb'startxref\s+(\d+)\s*%%EOF\s*$', data[-2048:])
    if not m:
        raise ValueError("no startxref")
    offsets, trailer, seen = {}, None, set()
    pos = int(m.group(1))
    while pos and pos not in seen and 0 <= pos < len(data):
        seen.add(pos)
        if data[pos:pos + 4] != b'xref':
            raise ValueError("cross-reference streams are not supported")
        i = pos + 4
        while True:
            mm = re.match(rb'\s*(\d+)\s+(\d+)\s*', data[i:i + 64])
            if not mm:
                break
            start, count = int(mm.group(1)), int(mm.group(2))
            i += mm.end()
            for k in range(count):
                ent = data[i + k * 20: i + k * 20 + 20]
                if ent[17:18] == b'n':
                    offsets.setdefault(start + k, int(ent[0:10]))
            i += count * 20
        t = data.find(b'trailer', i)
        if t == -1:
            break
        tdict = data[t + 7: t + 7 + 2048]
        if trailer is None:
            trailer = tdict
        p = re.search(rb'/Prev\s+(\d+)', tdict)
        pos = int(p.group(1)) if p else 0
    return offsets, trailer


def obj_body(data, offsets, num):
    off = offsets.get(num)
    if off is None:
        return None
    m = re.match(rb'\s*\d+\s+\d+\s+obj', data[off:off + 64])
    if not m:
        return None
    start = off + m.end()
    end = data.find(b'endobj', start)
    return data[start:end]


def page_objects(data, offsets, root_num):
    """Page object numbers in document order."""
    cat = obj_body(data, offsets, root_num)
    pm = re.search(rb'/Pages\s+(\d+)\s+0\s+R', cat or b'')
    if not pm:
        return []
    out = []

    def walk(num, depth=0):
        if depth > 64:
            return
        body = obj_body(data, offsets, num)
        if body is None:
            return
        if re.search(rb'/Type\s*/Page[^s]', body):
            out.append(num)
            return
        km = re.search(rb'/Kids\s*\[(.*?)\]', body, re.S)
        if not km:
            return
        for kid in re.findall(rb'(\d+)\s+0\s+R', km.group(1)):
            walk(int(kid), depth + 1)

    walk(int(pm.group(1)))
    return out


def media_box(data, offsets, num, depth=0):
    body = obj_body(data, offsets, num)
    if body is None or depth > 32:
        return (0.0, 0.0, 612.0, 792.0)
    m = re.search(rb'/MediaBox\s*\[\s*([\d.+-]+)\s+([\d.+-]+)\s+([\d.+-]+)\s+([\d.+-]+)', body)
    if m:
        return tuple(float(x) for x in m.groups())
    p = re.search(rb'/Parent\s+(\d+)\s+0\s+R', body)
    return media_box(data, offsets, int(p.group(1)), depth + 1) if p else (0.0, 0.0, 612.0, 792.0)


def pdf_text(s):
    """UTF-16BE hex string - safe for any character, no escaping rules."""
    return b'<FEFF' + s.encode('utf-16-be').hex().upper().encode('ascii') + b'>'


def add_notes(src, dst, notes):
    data = open(src, 'rb').read()
    offsets, trailer = parse_xref_chain(data)
    rm = re.search(rb'/Root\s+(\d+)\s+0\s+R', trailer or b'')
    if not rm:
        raise ValueError("no /Root in trailer")
    pages = page_objects(data, offsets, int(rm.group(1)))
    if not pages:
        raise ValueError("no pages found")

    sm = re.search(rb'/Size\s+(\d+)', trailer)
    next_num = int(sm.group(1)) if sm else max(offsets) + 1

    out = bytearray(data)
    if not out.endswith(b'\n'):
        out += b'\n'
    new_offsets, added = {}, 0

    for key, text in sorted(notes.items(), key=lambda kv: int(kv[0])):
        idx = int(key)
        if not (1 <= idx <= len(pages)):
            continue
        pnum = pages[idx - 1]
        body = obj_body(data, offsets, pnum)
        if body is None:
            continue
        x0, y0, x1, y1 = media_box(data, offsets, pnum)
        rect = (x0 + 12, y1 - 32, x0 + 32, y1 - 12)

        anum = next_num
        next_num += 1
        ann = (b'<< /Type /Annot /Subtype /Text /Name /Note /F 4 /C [1 0.85 0.2]\n'
               b'/Rect [%.2f %.2f %.2f %.2f]\n' % rect +
               b'/T ' + pdf_text("Speaker notes") + b'\n'
               b'/Contents ' + pdf_text(text) + b'\n'
               b'/P %d 0 R >>' % pnum)
        new_offsets[anum] = len(out)
        out += b'%d 0 obj\n' % anum + ann + b'\nendobj\n'

        # re-emit the page object with the annotation attached
        existing = re.search(rb'/Annots\s*\[(.*?)\]', body, re.S)
        if existing:
            new_body = (body[:existing.start(1)] + existing.group(1)
                        + b' %d 0 R' % anum + body[existing.end(1):])
        else:
            open_at = body.find(b'<<')
            if open_at == -1:
                continue
            new_body = (body[:open_at + 2] + b' /Annots [%d 0 R]' % anum
                        + body[open_at + 2:])
        new_offsets[pnum] = len(out)
        out += b'%d 0 obj' % pnum + new_body + b'endobj\n'
        added += 1

    if not added:
        return 0

    # incremental xref: only the objects this update touched
    xref_at = len(out)
    nums = sorted(new_offsets)
    sections, run = [], [nums[0]]
    for n in nums[1:]:
        if n == run[-1] + 1:
            run.append(n)
        else:
            sections.append(run); run = [n]
    sections.append(run)

    out += b'xref\n'
    for run in sections:
        out += b'%d %d\n' % (run[0], len(run))
        for n in run:
            out += b'%010d %05d n \n' % (new_offsets[n], 0)

    prev = int(re.search(rb'startxref\s+(\d+)\s*%%EOF\s*$', data[-2048:]).group(1))
    out += (b'trailer\n<< /Size %d /Root %s 0 R /Prev %d >>\nstartxref\n%d\n%%%%EOF\n'
            % (next_num, rm.group(1), prev, xref_at))

    open(dst, 'wb').write(bytes(out))
    return added


if __name__ == "__main__":
    notes = json.load(open(sys.argv[3]))
    n = add_notes(sys.argv[1], sys.argv[2], notes)
    print("  notes: attached to %d page%s" % (n, "" if n == 1 else "s"))
