"""
Expand a slide's click-triggered animation into one static slide per build state.

The XML is edited as text, not re-serialised: OOXML slides carry a dozen namespace
declarations (mc, p14, dgm, pvml ...) and any parser that drops the unused ones
produces a file PowerPoint refuses to open without repair.
"""
import zipfile, re, os, sys
import xml.etree.ElementTree as ET

NSP = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NSA = 'http://schemas.openxmlformats.org/drawingml/2006/main'
SHAPE_TAGS = ('p:sp', 'p:grpSp', 'p:pic', 'p:graphicFrame', 'p:cxnSp')


def click_steps(xml_bytes):
    """Ordered click steps as (entering spids, exiting spids)."""
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return []
    timing = root.find('{%s}timing' % NSP)
    if timing is None:
        return []
    steps = []
    for ctn in timing.iter('{%s}cTn' % NSP):
        if ctn.get('nodeType') != 'clickEffect':
            continue
        ent, exi = set(), set()
        for sub in ctn.iter('{%s}cTn' % NSP):
            pc = sub.get('presetClass')
            if pc not in ('entr', 'exit'):
                continue
            for spt in sub.iter('{%s}spTgt' % NSP):
                if spt.get('spid'):
                    (ent if pc == 'entr' else exi).add(spt.get('spid'))
        if ent or exi:
            steps.append((ent, exi))
    return steps


_SLD_ROOT = re.compile(rb'<p:sld\b[^>]*>')


def is_hidden(xml_bytes):
    """True if the slide is marked show="0".

    PowerPoint omits hidden slides from a PDF export. If we count them anyway
    the page map runs ahead of the real PDF and every note after the first
    hidden slide lands on the wrong page - silently, because the PDF still
    looks fine. Verified: a deck with one hidden 5-state slide mid-deck
    produced a 27-page map for a 22-page PDF.
    """
    m = _SLD_ROOT.search(xml_bytes)
    return bool(m and b'show="0"' in m.group(0))


def _top_level_shapes(xml):
    """(start, end, spid) for each direct child of p:spTree that is a shape."""
    m = re.search(r'<p:spTree[ >]', xml)
    if not m:
        return []
    i, depth, out = m.start(), 0, []
    tag_re = re.compile(r'<(/?)([A-Za-z0-9]+:[A-Za-z0-9]+)([^>]*?)(/?)>')
    open_start = None
    open_tag = None
    for t in tag_re.finditer(xml, m.start()):
        closing, name, attrs, selfclose = t.group(1), t.group(2), t.group(3), t.group(4)
        if name == 'p:spTree':
            if closing:
                break
            depth = 1
            continue
        if depth == 0:
            continue
        if not closing and not selfclose:
            if depth == 1 and name in SHAPE_TAGS:
                open_start, open_tag = t.start(), name
            depth += 1
        elif closing:
            depth -= 1
            if depth == 1 and name == open_tag and open_start is not None:
                block = xml[open_start:t.end()]
                sp = re.search(r'<p:cNvPr[^>]*\bid="(\d+)"', block)
                if sp:
                    out.append((open_start, t.end(), sp.group(1)))
                open_start = open_tag = None
        elif selfclose and depth == 1 and name in SHAPE_TAGS:
            block = t.group(0)
            sp = re.search(r'<p:cNvPr[^>]*\bid="(\d+)"', block)
            if sp:
                out.append((t.start(), t.end(), sp.group(1)))
    return out


def _drop_element(xml, tag):
    """Remove <tag ...>...</tag> or <tag .../> at any position."""
    m = re.search(r'<%s(\s[^>]*)?(/)?>' % re.escape(tag), xml)
    if not m:
        return xml
    if m.group(2):
        return xml[:m.start()] + xml[m.end():]
    close = xml.find('</%s>' % tag, m.end())
    if close == -1:
        return xml
    return xml[:m.start()] + xml[close + len(tag) + 3:]


def state_xml(xml_bytes, k, steps):
    """Slide XML as it looks after k clicks. k=0 is the base state."""
    xml = xml_bytes.decode('utf-8')
    later, gone = set(), set()
    for i, (ent, exi) in enumerate(steps, 1):
        if i > k:
            later |= ent
        else:
            gone |= exi
    hide = later | gone
    for start, end, spid in sorted(_top_level_shapes(xml), reverse=True):
        if spid in hide:
            xml = xml[:start] + xml[end:]
    xml = _drop_element(xml, 'p:timing')      # states are static
    xml = _drop_element(xml, 'p:transition')
    return xml.encode('utf-8')


def variants(src, slide_part, outdir):
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
        blobs = {i.filename: z.read(i.filename) for i in infos}
    xml = blobs[slide_part]
    steps = click_steps(xml)
    os.makedirs(outdir, exist_ok=True)
    made = []
    for k in range(len(steps) + 1):
        out = os.path.join(outdir, "state%02d.pptx" % k)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as w:
            for i in infos:
                data = state_xml(xml, k, steps) if i.filename == slide_part else blobs[i.filename]
                w.writestr(i.filename, data)
        made.append(out)
    return steps, made


# ---------------------------------------------------------------------------
# Whole-deck expansion: rewrite a deck so every animated slide is followed by
# one static slide per build state. A single PowerPoint export then yields one
# PDF carrying slides and sub-slides in reading order.
# ---------------------------------------------------------------------------
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
SLIDE_CT = ('application/vnd.openxmlformats-officedocument.'
            'presentationml.slide+xml')


def slide_order(blobs):
    """Slide parts in presentation order, as (part_name, sldId_element_text)."""
    pres = blobs['ppt/presentation.xml'].decode('utf-8')
    rels = blobs['ppt/_rels/presentation.xml.rels'].decode('utf-8')
    target = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', rels):
        tag = m.group(0)
        rid = re.search(r'Id="([^"]+)"', tag)
        tgt = re.search(r'Target="([^"]+)"', tag)
        if rid and tgt and '/slide' in ('/' + tgt.group(1)):
            t = tgt.group(1).lstrip('/')
            if t.startswith('slides/slide'):
                target[rid.group(1)] = 'ppt/' + t
    out = []
    lst = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', pres, re.S)
    if not lst:
        return out
    for m in re.finditer(r'<p:sldId\b[^>]*/>', lst.group(1)):
        tag = m.group(0)
        rid = re.search(r'r:id="([^"]+)"', tag)
        if rid and rid.group(1) in target:
            out.append((target[rid.group(1)], tag))
    return out


def expand_deck(src, dst):
    """Write `dst`: `src` with each animated slide followed by its build states.

    Returns [(slide_number, n_states), ...] for the slides that were expanded.
    """
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
        blobs = {i.filename: z.read(i.filename) for i in infos}

    order = slide_order(blobs)
    used_nums = {int(re.search(r'slide(\d+)\.xml$', n).group(1))
                 for n in blobs if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)}
    next_num = max(used_nums) + 1 if used_nums else 1

    rels = blobs['ppt/_rels/presentation.xml.rels'].decode('utf-8')
    used_rids = set(re.findall(r'Id="(rId\d+)"', rels))
    next_rid = max((int(r[3:]) for r in used_rids), default=0) + 1

    pres = blobs['ppt/presentation.xml'].decode('utf-8')
    used_sids = {int(s) for s in re.findall(r'<p:sldId\b[^>]*\bid="(\d+)"', pres)}
    next_sid = max(used_sids | {255}) + 1

    new_parts, new_rels, new_cts, expanded = {}, [], [], []
    insert_after = {}                        # original sldId tag -> [new tags]

    for slide_no, (part, sld_tag) in enumerate(order, 1):
        xml = blobs[part]
        if is_hidden(xml):
            continue                         # PowerPoint will not export it
        steps = click_steps(xml)
        if not steps:
            continue
        rels_part = part.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
        base_rels = blobs.get(rels_part, b'')
        # A notesSlide belongs to exactly one slide; don't let copies claim it.
        copy_rels = re.sub(
            r'<Relationship\b[^>]*notesSlide[^>]*/>', '',
            base_rels.decode('utf-8')).encode('utf-8') if base_rels else b''

        blobs[part] = state_xml(xml, 0, steps)      # original becomes state 0
        tags = []
        for k in range(1, len(steps) + 1):
            name = 'ppt/slides/slide%d.xml' % next_num
            next_num += 1
            new_parts[name] = state_xml(xml, k, steps)
            if copy_rels:
                new_parts[name.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'] = copy_rels
            rid = 'rId%d' % next_rid
            next_rid += 1
            new_rels.append(
                '<Relationship Id="%s" Type="%s/slide" Target="slides/%s"/>'
                % (rid, NS_R, os.path.basename(name)))
            new_cts.append('<Override PartName="/%s" ContentType="%s"/>' % (name, SLIDE_CT))
            tags.append('<p:sldId id="%d" r:id="%s"/>' % (next_sid, rid))
            next_sid += 1
        insert_after[sld_tag] = tags
        expanded.append((slide_no, len(steps) + 1))

    if not expanded:
        with open(dst, 'wb') as f, open(src, 'rb') as g:
            f.write(g.read())
        return []

    for old, tags in insert_after.items():
        pres = pres.replace(old, old + ''.join(tags), 1)
    blobs['ppt/presentation.xml'] = pres.encode('utf-8')

    rels = rels.replace('</Relationships>', ''.join(new_rels) + '</Relationships>')
    blobs['ppt/_rels/presentation.xml.rels'] = rels.encode('utf-8')

    ct = blobs['[Content_Types].xml'].decode('utf-8')
    ct = ct.replace('</Types>', ''.join(new_cts) + '</Types>')
    blobs['[Content_Types].xml'] = ct.encode('utf-8')

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as w:
        for i in infos:
            w.writestr(i.filename, blobs[i.filename])
        for n, d in new_parts.items():
            w.writestr(n, d)
    return expanded


def page_map(src, expanded):
    """Page number -> what it shows, for the expanded deck.

    Walks presentation order (sldIdLst), not the slideN.xml file names: the two
    agree in most decks but nothing guarantees it, and `expand_deck` numbers
    slides by presentation position. Hidden slides are skipped because
    PowerPoint does not export them.
    """
    with zipfile.ZipFile(src) as z:
        blobs = {n: z.read(n) for n in z.namelist()
                 if n in ('ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels')
                 or re.fullmatch(r'ppt/slides/slide\d+\.xml', n)}
    order = slide_order(blobs)
    extra = {slide: pages for slide, pages in expanded}
    pages, page = [], 1
    for slide_no, (part, _tag) in enumerate(order, 1):
        if is_hidden(blobs[part]):
            continue
        n = extra.get(slide_no, 1)
        for k in range(n):
            pages.append({"page": page, "slide": slide_no,
                          "state": k if n > 1 else None,
                          "of": n if n > 1 else None})
            page += 1
    return pages


if __name__ == "__main__":
    if sys.argv[1] == "expand":
        src, dst = sys.argv[2], sys.argv[3]
        expanded = expand_deck(src, dst)
        if len(sys.argv) > 4:
            import json
            json.dump({"source": os.path.basename(src),
                       "expanded_slides": [{"slide": s, "pages": n} for s, n in expanded],
                       "pages": page_map(src, expanded)},
                      open(sys.argv[4], "w"), indent=1)
        if expanded:
            print("  builds: %d slide%s expanded into %d extra page%s"
                  % (len(expanded), "" if len(expanded) == 1 else "s",
                     sum(n - 1 for _, n in expanded),
                     "" if sum(n - 1 for _, n in expanded) == 1 else "s"))
    else:
        steps, made = variants(sys.argv[1], sys.argv[2], sys.argv[3])
        print("%d click steps -> %d static states" % (len(steps), len(made)))
