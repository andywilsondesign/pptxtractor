#!/usr/bin/env python3
"""
Pull the playable media out of a .pptx: video, audio and animated GIFs, named by
the slide that uses them. Static artwork is skipped - it is already in the PDF.

    python3 extract-media.py <deck.pptx> <outdir> [--svg] [--all-gifs]
"""
import json, os, re, sys, zipfile
from collections import defaultdict

VIDEO = {'.mp4', '.mov', '.m4v', '.avi', '.wmv', '.mpg', '.mpeg', '.mkv', '.webm', '.asf', '.3gp'}
AUDIO = {'.mp3', '.wav', '.m4a', '.aac', '.wma', '.aiff', '.au', '.mid', '.midi'}
# .wdp is a sidecar for image effects, not media anyone wants; .emf/.wmf are vector clip art
NEVER = {'.wdp', '.emf', '.wmf', '.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp', '.webp'}


def gif_frames(data):
    """Frame count of a GIF: 1 means it is a still, more means it animates."""
    if not data.startswith((b'GIF87a', b'GIF89a')):
        return None
    i, flags, n = 13, data[10], 0
    if flags & 0x80:
        i += 3 * (2 ** ((flags & 7) + 1))
    while i < len(data):
        b = data[i]
        if b == 0x3B:
            break
        if b == 0x2C:                       # image descriptor = one frame
            n += 1
            i += 10
            lf = data[i - 1]
            if lf & 0x80:
                i += 3 * (2 ** ((lf & 7) + 1))
            i += 1
            while i < len(data) and data[i]:
                i += data[i] + 1
            i += 1
        elif b == 0x21:                     # extension block
            i += 2
            while i < len(data) and data[i]:
                i += data[i] + 1
            i += 1
        else:
            i += 1
    return n


def slide_usage(z):
    """media filename -> sorted list of slide numbers that reference it."""
    use = defaultdict(set)
    for n in z.namelist():
        m = re.fullmatch(r'ppt/slides/_rels/slide(\d+)\.xml\.rels', n)
        if not m:
            continue
        idx = int(m.group(1))
        for t in re.findall(r'Target="\.\./media/([^"]+)"', z.read(n).decode('utf8', 'ignore')):
            use[t].add(idx)
    return {k: sorted(v) for k, v in use.items()}


def external_links(z):
    out = []
    for n in z.namelist():
        m = re.fullmatch(r'ppt/slides/_rels/slide(\d+)\.xml\.rels', n)
        if not m:
            continue
        idx = int(m.group(1))
        x = z.read(n).decode('utf8', 'ignore')
        for tag in re.findall(r'<Relationship\b[^>]*TargetMode="External"[^>]*/>', x):
            t = re.search(r'Target="([^"]+)"', tag)
            if not t:
                continue
            url = t.group(1).replace('&amp;', '&').replace('&quot;', '"')
            if (os.path.splitext(url)[1].lower() in VIDEO | AUDIO
                    or re.search(r'youtube|youtu\.be|vimeo|wistia|loom|\.mp4|\.mov', url, re.I)):
                out.append({'slide': idx, 'url': url})
    return out


def extract(src, outdir, want_svg=False, all_gifs=False):
    stem = os.path.splitext(os.path.basename(src))[0]
    items, links = [], []
    with zipfile.ZipFile(src) as z:
        usage = slide_usage(z)
        links = external_links(z)
        media = [n for n in z.namelist() if n.startswith('ppt/media/')]
        for n in sorted(media):
            base = os.path.basename(n)
            ext = os.path.splitext(base)[1].lower()
            if ext in NEVER:
                continue
            if ext == '.svg' and not want_svg:
                continue
            data = z.read(n)
            kind, frames = None, None
            if ext in VIDEO:
                kind = 'video'
            elif ext in AUDIO:
                kind = 'audio'
            elif ext == '.gif':
                frames = gif_frames(data)
                if frames is not None and frames <= 1 and not all_gifs:
                    continue
                kind = 'animated-gif'
            elif ext == '.svg':
                kind = 'svg'
            else:
                kind = 'other'
            slides = usage.get(base, [])
            tag = ('slide%03d' % slides[0]) if slides else 'unplaced'
            name = '%s-%s-%s' % (stem, tag, base)
            os.makedirs(outdir, exist_ok=True)
            with open(os.path.join(outdir, name), 'wb') as f:
                f.write(data)
            items.append({'file': name, 'kind': kind, 'source': base,
                          'slides': slides, 'bytes': len(data),
                          **({'frames': frames} if frames is not None else {})})
    if items or links:
        os.makedirs(outdir, exist_ok=True)
        with open(os.path.join(outdir, 'media.json'), 'w') as f:
            json.dump({'deck': os.path.basename(src), 'items': items,
                       'external': links}, f, indent=1)
    return items, links


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    items, links = extract(args[0], args[1],
                           want_svg='--svg' in sys.argv,
                           all_gifs='--all-gifs' in sys.argv)
    if items or links:
        by = {}
        for i in items:
            by[i['kind']] = by.get(i['kind'], 0) + 1
        parts = ['%d %s' % (v, k) for k, v in sorted(by.items())]
        if links:
            parts.append('%d external link%s' % (len(links), '' if len(links) == 1 else 's'))
        print('  media: ' + ', '.join(parts))
