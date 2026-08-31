#!/usr/bin/env python3
"""Build the test fixtures from scratch - no real decks, nothing proprietary."""
import os, struct, zipfile

NS_P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
NS_A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
NS_MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'


def shape(sid, name, text):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/>'
            f'<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/>'
            f'<a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>'
            f'<a:rPr lang="en-GB"><a:latin typeface="Helvetica"/></a:rPr>'
            f'<a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>')


def timing(spids):
    """A main sequence with one click-triggered entrance per shape id."""
    pars = ""
    for i, sp in enumerate(spids):
        pars += (f'<p:par><p:cTn id="{10+i}" fill="hold" nodeType="clickEffect">'
                 f'<p:childTnLst><p:par><p:cTn id="{20+i}" presetClass="entr" fill="hold">'
                 f'<p:childTnLst><p:set><p:cBhvr><p:cTn id="{30+i}" dur="1"/>'
                 f'<p:tgtEl><p:spTgt spid="{sp}"/></p:tgtEl></p:cBhvr></p:set>'
                 f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    return ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never"'
            ' nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">'
            '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            + pars + '</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par>'
            '</p:tnLst></p:timing>')


def slide(shapes, tim=""):
    return (XD + f'<p:sld {NS_P} {NS_A} {NS_R} {NS_MC}><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name="Shape 1"/><p:cNvGrpSpPr/><p:nvPr/>'
            '</p:nvGrpSpPr><p:grpSpPr/>' + "".join(shapes) +
            '</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping/></p:clrMapOvr>'
            + tim + '</p:sld>')


def notes(text):
    body = "".join(f'<a:p><a:r><a:rPr lang="en-GB"/><a:t>{ln}</a:t></a:r></a:p>'
                   for ln in text.split("\n"))
    return (XD + f'<p:notes {NS_P} {NS_A} {NS_R}><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="img"/><p:cNvSpPr/>'
            '<p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/>'
            '<p:txBody><a:bodyPr/><a:p><a:r><a:t>7</a:t></a:r></a:p></p:txBody></p:sp>'
            '<p:sp><p:nvSpPr><p:cNvPr id="3" name="body"/><p:cNvSpPr/>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
            f'<p:txBody><a:bodyPr/>{body}</p:txBody></p:sp>'
            '</p:spTree></p:cSld></p:notes>')


def animated_gif():
    """Smallest honest 2-frame GIF: header, 2 image descriptors, trailer."""
    out = bytearray(b'GIF89a\x01\x00\x01\x00\x80\x00\x00'
                    b'\x00\x00\x00\xff\xff\xff')          # 2-colour global table
    for _ in range(2):
        out += b'\x21\xf9\x04\x00\x0a\x00\x00\x00'        # graphic control
        out += b'\x2c\x00\x00\x00\x00\x01\x00\x01\x00\x00'  # image descriptor
        out += b'\x02\x02\x44\x01\x00'                    # 1px LZW data
    out += b'\x3b'
    return bytes(out)


def build_pptx(path):
    p = {}
    p['[Content_Types].xml'] = XD + (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="gif" ContentType="image/gif"/>'
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '</Types>')
    p['_rels/.rels'] = XD + (
        f'<Relationships xmlns="{REL}"><Relationship Id="rId1" '
        f'Type="{REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
    p['ppt/presentation.xml'] = XD + (
        f'<p:presentation {NS_P} {NS_A} {NS_R}><p:sldIdLst>'
        '<p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/>'
        '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/>'
        '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
    p['ppt/_rels/presentation.xml.rels'] = XD + (
        f'<Relationships xmlns="{REL}">'
        f'<Relationship Id="rId1" Type="{REL}/slide" Target="slides/slide1.xml"/>'
        f'<Relationship Id="rId2" Type="{REL}/slide" Target="slides/slide2.xml"/>'
        f'<Relationship Id="rId3" Type="{REL}/theme" Target="theme/theme1.xml"/>'
        '</Relationships>')
    # slide 1: three shapes, two of them revealed by clicks -> 3 build states
    p['ppt/slides/slide1.xml'] = slide(
        [shape(2, 'Title', 'Fixture title'), shape(3, 'One', 'First'), shape(4, 'Two', 'Second')],
        timing([3, 4]))
    p['ppt/slides/slide2.xml'] = slide([shape(2, 'Plain', 'No animation here')])
    p['ppt/slides/_rels/slide1.xml.rels'] = XD + (
        f'<Relationships xmlns="{REL}">'
        f'<Relationship Id="rId1" Type="{REL}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>'
        f'<Relationship Id="rId2" Type="{REL}/image" Target="../media/loop.gif"/>'
        f'<Relationship Id="rId3" Type="{REL}/video" Target="https://www.youtube.com/watch?v=TESTID" TargetMode="External"/>'
        '</Relationships>')
    p['ppt/slides/_rels/slide2.xml.rels'] = XD + f'<Relationships xmlns="{REL}"/>'
    p['ppt/notesSlides/notesSlide1.xml'] = notes("Fixture speaker notes.\nSecond line.")
    p['ppt/notesSlides/_rels/notesSlide1.xml.rels'] = XD + f'<Relationships xmlns="{REL}"/>'
    p['ppt/theme/theme1.xml'] = XD + (
        f'<a:theme {NS_A} name="Fixture"><a:themeElements><a:fontScheme name="F">'
        '<a:majorFont><a:latin typeface="Helvetica"/></a:majorFont>'
        '<a:minorFont><a:latin typeface="NotARealFontFace"/></a:minorFont>'
        '</a:fontScheme></a:themeElements></a:theme>')

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, body in p.items():
            z.writestr(name, body)
        z.writestr('ppt/media/loop.gif', animated_gif())
    return path


def build_pdf(path, pages=2):
    """A minimal, valid classic-xref PDF - what add_notes.py must update."""
    objs, body = [], b'%PDF-1.4\n'
    kids = " ".join("%d 0 R" % (3 + i) for i in range(pages))

    def add(num, payload):
        nonlocal body
        objs.append((num, len(body)))
        body += b'%d 0 obj\n' % num + payload + b'\nendobj\n'

    add(1, b'<< /Type /Catalog /Pages 2 0 R >>')
    add(2, ('<< /Type /Pages /Count %d /Kids [%s] >>' % (pages, kids)).encode())
    for i in range(pages):
        add(3 + i, ('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 720 405] '
                    '/Contents %d 0 R >>' % (3 + pages + i)).encode())
    for i in range(pages):
        stream = b'BT /F1 24 Tf 60 200 Td (Page %d) Tj ET' % (i + 1)
        add(3 + pages + i, b'<< /Length %d >>\nstream\n' % len(stream) + stream + b'\nendstream')

    xref_at = len(body)
    size = 3 + 2 * pages
    body += b'xref\n0 %d\n' % size
    body += b'0000000000 65535 f \n'
    table = {n: off for n, off in objs}
    for n in range(1, size):
        body += b'%010d %05d n \n' % (table.get(n, 0), 0)
    body += (b'trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n'
             % (size, xref_at))
    open(path, 'wb').write(body)
    return path


if __name__ == "__main__":
    here = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")
    os.makedirs(here, exist_ok=True)
    print(build_pptx(os.path.join(here, "fixture.pptx")))
    print(build_pdf(os.path.join(here, "fixture.pdf")))
