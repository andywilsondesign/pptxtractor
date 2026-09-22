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
    """Ordered click steps.

    Each step is (entering spids, exiting spids, entering paragraphs,
    exiting paragraphs), where a paragraph is (spid, paragraph index).

    The distinction matters. PowerPoint's default way to build a bulleted
    list is to animate one text box paragraph by paragraph, which reaches the
    XML as several steps that all name the same shape and differ only in a
    <p:txEl><p:pRg>. Reading the spid alone makes those steps look identical,
    and hiding the whole shape for each of them produces a run of identical
    frames followed by one where everything appears at once.
    """
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
        ent, exi, entp, exip = set(), set(), set(), set()
        for sub in ctn.iter('{%s}cTn' % NSP):
            pc = sub.get('presetClass')
            if pc not in ('entr', 'exit'):
                continue
            for spt in sub.iter('{%s}spTgt' % NSP):
                spid = spt.get('spid')
                if not spid:
                    continue
                rng = []
                for prg in spt.iter('{%s}pRg' % NSP):
                    try:
                        st = int(prg.get('st', '0'))
                        en = int(prg.get('end', st))
                    except ValueError:
                        continue
                    if en >= st:
                        rng.extend(range(st, en + 1))
                if rng:
                    (entp if pc == 'entr' else exip).update((spid, i) for i in rng)
                else:
                    (ent if pc == 'entr' else exi).add(spid)
        if ent or exi or entp or exip:
            steps.append((ent, exi, entp, exip))
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


# <a:p> is a paragraph; <a:pPr> is its properties and must not match. The
# lookahead keeps the two apart without matching every tag starting "a:p".
_PARA_OPEN = re.compile(r'<a:p(?=[\s/>])[^>]*>')


def _paragraph_spans(block):
    """(start, end) of each paragraph in this shape's text body, in order.

    Paragraphs do not nest, so a flat scan is enough - but it has to start at
    the text body, or the properties block above it contributes spurious
    matches and every index shifts.
    """
    tb = re.search(r'<p:txBody[\s>]', block)
    if not tb:
        return []
    spans, pos = [], tb.start()
    while True:
        m = _PARA_OPEN.search(block, pos)
        if not m:
            return spans
        if m.group(0).endswith('/>'):         # <a:p/>, an empty paragraph
            spans.append((m.start(), m.end()))
            pos = m.end()
            continue
        close = block.find('</a:p>', m.end())
        if close == -1:
            return spans
        spans.append((m.start(), close + len('</a:p>')))
        pos = close + len('</a:p>')


def _keep_one_paragraph(block):
    """Leave every text body with at least one paragraph.

    A <p:txBody> must contain one or more <a:p>; the schema has no way to say
    "no text". Hiding every paragraph of a box therefore produces a file
    PowerPoint offers to repair rather than open - and driven by AppleScript
    that prompt is never answered, so the deck reports as a timeout with no
    clue as to why. An empty paragraph shows nothing and keeps the file legal.
    """
    out, pos = [], 0
    for m in re.finditer(r'<p:txBody>(.*?)</p:txBody>', block, re.S):
        if re.search(r'<a:p(?=[\s/>])', m.group(1)):
            continue
        out.append((m.end() - len('</p:txBody>'), '<a:p/>'))
    if not out:
        return block
    parts, last = [], 0
    for at, ins in out:
        parts.append(block[last:at]); parts.append(ins); last = at
    parts.append(block[last:])
    return "".join(parts)


def state_xml(xml_bytes, k, steps):
    """Slide XML as it looks after k clicks. k=0 is the base state."""
    xml = xml_bytes.decode('utf-8')
    later, gone = set(), set()
    para_hide = {}
    for i, (ent, exi, entp, exip) in enumerate(steps, 1):
        if i > k:
            later |= ent
            for spid, idx in entp:
                para_hide.setdefault(spid, set()).add(idx)
        else:
            gone |= exi
            for spid, idx in exip:
                para_hide.setdefault(spid, set()).add(idx)
    hide = later | gone
    # A shape whose paragraphs are animated must survive as a shape, or the
    # whole list vanishes and every state before the last looks the same.
    hide -= set(para_hide)
    for start, end, spid in sorted(_top_level_shapes(xml), reverse=True):
        if spid in hide:
            xml = xml[:start] + xml[end:]
            continue
        drop = para_hide.get(spid)
        if not drop:
            continue
        block = xml[start:end]
        spans = _paragraph_spans(block)
        # Reverse order so earlier offsets stay valid as later ones are cut.
        for idx in sorted(drop, reverse=True):
            if 0 <= idx < len(spans):
                ps, pe = spans[idx]
                block = block[:ps] + block[pe:]
        block = _keep_one_paragraph(block)
        xml = xml[:start] + block + xml[end:]
    xml = _drop_element(xml, 'p:timing')      # states are static
    xml = _drop_element(xml, 'p:transition')
    return xml.encode('utf-8')


def distinct_states(xml, steps):
    """Every build state, with consecutive duplicates removed.

    A step that changes nothing visible still costs a page, and a run of them
    reads as the exporter repeating itself. Comparing the generated XML catches
    that exactly and cheaply - two states that produce the same markup produce
    the same page, whatever the animation claimed to do. Doing it here rather
    than by comparing rendered images matters: the rendered pages are *not*
    identical, because each carries its own slide number, so a pixel comparison
    finds a difference and keeps the duplicate.
    """
    kept = []
    for k in range(len(steps) + 1):
        body = state_xml(xml, k, steps)
        if not kept or body != kept[-1]:
            kept.append(body)
    return kept


def variants(src, slide_part, outdir):
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
        blobs = {i.filename: z.read(i.filename) for i in infos}
    xml = blobs[slide_part]
    steps = click_steps(xml)
    os.makedirs(outdir, exist_ok=True)
    made = []
    for k, body in enumerate(distinct_states(xml, steps)):
        out = os.path.join(outdir, "state%02d.pptx" % k)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as w:
            for i in infos:
                data = body if i.filename == slide_part else blobs[i.filename]
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

        states = distinct_states(xml, steps)
        blobs[part] = states[0]                     # original becomes state 0
        if len(states) == 1:
            # Every step produced the same slide, so there is nothing to show
            # between them. The base state stands alone, timing stripped.
            continue
        tags = []
        for body in states[1:]:
            name = 'ppt/slides/slide%d.xml' % next_num
            next_num += 1
            new_parts[name] = body
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
        expanded.append((slide_no, len(states)))

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

    # Keep each part packed the way it arrived. A deck can store its media
    # uncompressed - one 349 MB deck here stores 341 of its 858 parts that way,
    # including an 86 MB TIFF - and deflating all of it turns a package
    # PowerPoint could read straight through into 350 MB it must inflate before
    # it can show a slide. Re-packing is not our business; only the slides we
    # rewrote have changed, so everything else goes back as it came.
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as w:
        for i in infos:
            info = zipfile.ZipInfo(i.filename, date_time=i.date_time)
            info.compress_type = i.compress_type
            info.external_attr = i.external_attr
            info.internal_attr = i.internal_attr
            info.create_system = i.create_system
            w.writestr(info, blobs[i.filename])
        for n, d in new_parts.items():
            w.writestr(n, d)                     # new slide XML: deflate is right
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


def validate_deck(path):
    """Check a deck we generated is one PowerPoint will actually open.

    This exists because the alternative is so expensive. Handed a deck it
    considers damaged, PowerPoint does not fail - it puts up a "found a problem
    with content" dialog and waits. Driven by AppleScript nobody answers it, so
    the run burns a full deadline, the watchdog kills it, and the deck is
    reported as a timeout: the one explanation that is certainly wrong. Two
    separate faults reached a real library that way, each costing about forty
    minutes per deck to learn nothing.

    Returns a list of problems; empty means it is worth handing over.
    """
    problems = []
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return ["not a readable package: %s" % e]
    with z:
        names = set(z.namelist())
        for n in sorted(names):
            if not (n.endswith('.xml') or n.endswith('.rels')):
                continue
            try:
                ET.fromstring(z.read(n))
            except ET.ParseError as e:
                problems.append("%s is not well-formed XML: %s" % (n, e))
        slides = sorted(n for n in names
                        if re.fullmatch(r'ppt/slides/slide\d+\.xml', n))
        # A text body must hold at least one paragraph; the schema cannot say
        # "no text", and an empty one is enough on its own to trigger repair.
        for n in slides:
            body = z.read(n).decode('utf-8', 'ignore')
            for m in re.finditer(r'<p:txBody>(.*?)</p:txBody>', body, re.S):
                if not re.search(r'<a:p(?=[\s/>])', m.group(1)):
                    problems.append("%s has a text body with no paragraphs" % n)
                    break
            rels = n.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
            if rels not in names:
                problems.append("%s has no relationships part" % n)
            elif 'slideLayout' not in z.read(rels).decode('utf-8', 'ignore'):
                problems.append("%s is not related to a slide layout" % n)
        try:
            ct = z.read('[Content_Types].xml').decode('utf-8', 'ignore')
        except KeyError:
            return problems + ["no [Content_Types].xml"]
        for n in slides:
            if 'PartName="/%s"' % n not in ct:
                problems.append("%s has no content-type override" % n)
        try:
            pres = z.read('ppt/presentation.xml').decode('utf-8', 'ignore')
            prels = z.read('ppt/_rels/presentation.xml.rels').decode('utf-8', 'ignore')
        except KeyError as e:
            return problems + ["missing %s" % e]
        target = {}
        for m in re.finditer(r'<Relationship\b[^>]*>', prels):
            i = re.search(r'Id="([^"]+)"', m.group(0))
            t = re.search(r'Target="([^"]+)"', m.group(0))
            if i and t:
                target[i.group(1)] = t.group(1)
        seen_id, seen_rid = set(), set()
        for m in re.finditer(r'<p:sldId\b[^>]*\bid="(\d+)"[^>]*r:id="([^"]+)"', pres):
            sid, rid = m.group(1), m.group(2)
            if sid in seen_id:
                problems.append("duplicate slide id %s" % sid)
            if rid in seen_rid:
                problems.append("slide relationship %s used twice" % rid)
            seen_id.add(sid); seen_rid.add(rid)
            if rid not in target:
                problems.append("slide relationship %s resolves to nothing" % rid)
            elif ('ppt/' + target[rid].lstrip('/')) not in names:
                problems.append("slide %s points at a missing part" % rid)
    # De-duplicate but keep order, and keep the list readable.
    out = []
    for p in problems:
        if p not in out:
            out.append(p)
    return out


if __name__ == "__main__":
    # `map` writes the page map for a deck nobody is expanding: one page per
    # visible slide. Worth having even then, because hidden slides mean page
    # number and slide number stop agreeing after the first one.
    if sys.argv[1] == "map":
        import json
        src, out = sys.argv[2], sys.argv[3]
        json.dump({"source": os.path.basename(src), "expanded_slides": [],
                   "pages": page_map(src, [])}, open(out, "w"), indent=1)
        sys.exit(0)
    if sys.argv[1] == "expand":
        src, dst = sys.argv[2], sys.argv[3]
        expanded = expand_deck(src, dst)
        # Check our own work before PowerPoint has to. Exiting non-zero here
        # makes the caller fall back to the untouched deck, which costs the
        # build pages and nothing else - the alternative is a modal dialog and
        # a dead deadline.
        faults = validate_deck(dst)
        if faults:
            sys.stderr.write(
                "expansion produced a deck PowerPoint would refuse - "
                "falling back to the original:\n")
            for f in faults[:6]:
                sys.stderr.write("    %s\n" % f)
            if len(faults) > 6:
                sys.stderr.write("    ... and %d more\n" % (len(faults) - 6))
            try:
                os.remove(dst)
            except OSError:
                pass
            sys.exit(1)
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
