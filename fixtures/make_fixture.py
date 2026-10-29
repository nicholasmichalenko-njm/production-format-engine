"""Build a proposal-shaped .docx from the PE sample master template.

Usage: python -I make_fixture.py <template.dotx> <logo.png> <out.docx> <chapters> <manifest.json>

The raw output is then opened and saved by real Word (harness/word_normalize.py), so the engine
is always tested on files that Word itself wrote.

Planted defects (with the correct answer in the manifest):
 - caption fields without \\s, with MERGEFORMAT, with a wrong SEQ identifier, appendix captions on the body rule
 - table rows with direct run formatting, no header repeat, empty note rows
 - a REF to a bookmark that does not exist (every 7th chapter)
 - PAGE fields with a format switch in the footer, STYLEREF with MERGEFORMAT in the header
 - no page number settings on the sections
"""
import json
import random
import sys
import zipfile
from xml.sax.saxutils import escape

WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

rnd = random.Random(20261028)
NOUNS = ["mission assurance", "ground segment", "payload integration", "flight software", "thermal model", "launch campaign",
         "risk register", "configuration baseline", "test campaign", "systems engineering plan", "supply chain", "crew module",
         "docking adapter", "power subsystem", "verification matrix", "program schedule", "quality plan", "safety review"]
VERBS = ["delivers", "verifies", "integrates", "reduces", "tracks", "documents", "controls", "validates", "supports", "maintains"]
ADJ = ["proven", "integrated", "traceable", "low-risk", "modular", "flight-qualified", "repeatable", "independent", "measured"]


def sentence():
    return (f"The {rnd.choice(ADJ)} {rnd.choice(NOUNS)} {rnd.choice(VERBS)} the {rnd.choice(NOUNS)} "
            f"through a {rnd.choice(ADJ)} {rnd.choice(NOUNS)} and a {rnd.choice(ADJ)} {rnd.choice(NOUNS)}.")


def para(n=6):
    return " ".join(sentence() for _ in range(n))


def run(text, rpr=""):
    t = escape(text)
    sp = ' xml:space="preserve"' if text != text.strip() or "  " in text else ""
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t{sp}>{t}</w:t></w:r>'


def fld(instr, result, rpr=""):
    rp = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return (f'<w:r>{rp}<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r>{rp}<w:instrText xml:space="preserve"> {escape(instr)} </w:instrText></w:r>'
            f'<w:r>{rp}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'{run(result, rpr)}<w:r>{rp}<w:fldChar w:fldCharType="end"/></w:r>')


def p(content, style=None, ppr_extra=""):
    st = f'<w:pStyle w:val="{style}"/>' if style else ""
    ppr = f"<w:pPr>{st}{ppr_extra}</w:pPr>" if (st or ppr_extra) else ""
    return f"<w:p>{ppr}{content}</w:p>"


class Doc:
    def __init__(self):
        self.body = []
        self.bm = 100
        self.ref_targets = []  # (bookmark, label)
        self.manifest = {"captions": [], "broken_refs": [], "tables": 0, "figures": 0, "sections": []}

    def bookmark(self, inner):
        self.bm += 1
        name = f"_Ref{100000000 + self.bm}"
        return name, f'<w:bookmarkStart w:id="{self.bm}" w:name="{name}"/>{inner}<w:bookmarkEnd w:id="{self.bm}"/>'


def caption_runs(label, ident, level, appendix, defect, typed=None):
    """Label + number fields. defect in: None, 'no_s', 'mergeformat', 'wrong_ident', 'typed'."""
    if defect == "typed":
        return run(f"{label} {typed}")
    sref = "STYLEREF 1 \\s" if not appendix else "STYLEREF 1 \\s"  # appendix on the body rule = defect
    seq = f"SEQ {ident} \\* ARABIC \\s 1"
    if defect == "no_s":
        seq = f"SEQ {ident} \\* ARABIC"
    if defect == "mergeformat":
        sref += " \\* MERGEFORMAT"
    if defect == "wrong_ident":
        seq = f"SEQ Tbl \\* ARABIC \\s 1"
    return run(label + " ") + fld(sref, "1") + run("-") + fld(seq, "1")


def table_xml(d, ch, ti, appendix, defect, cols=4, rows=12, landscape=False):
    W_ = 9360 if not landscape else 12960
    cw = W_ // cols
    grid = "".join(f'<w:gridCol w:w="{cw}"/>' for _ in range(cols))
    cap_inner = caption_runs("Table", "Table", 1, appendix, defect, typed=f"{ch}-{ti}")
    name, cap = d.bookmark(cap_inner)
    if defect != "typed":
        d.ref_targets.append((name, "Table"))
    title = f"{rnd.choice(ADJ).title()} {rnd.choice(NOUNS).title()} Summary"
    cap_p = p(cap + run(".") + '<w:r><w:tab/></w:r>' + run(title), "Caption")
    d.manifest["captions"].append({"chapter": ch, "kind": "Table", "appendix": appendix, "defect": defect, "title": title})
    trs = [f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{W_}" w:type="dxa"/><w:gridSpan w:val="{cols}"/></w:tcPr>{cap_p}</w:tc></w:tr>']
    hdr = "".join(f'<w:tc><w:tcPr><w:tcW w:w="{cw}" w:type="dxa"/></w:tcPr>{p(run(h, "<w:b/>"))}</w:tc>'
                  for h in ["Item", "Description", "Owner", "Status"][:cols])
    trs.append(f"<w:tr>{hdr}</w:tr>")
    for r in range(rows):
        noise = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="18"/>' if r % 3 == 0 else ""
        if r == 4:  # subhead row, shaded, one merged cell
            trs.append(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{W_}" w:type="dxa"/><w:gridSpan w:val="{cols}"/>'
                       f'<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr>{p(run("Phase " + str(r)))}</w:tc></w:tr>')
            continue
        cells = []
        for c in range(cols):
            vm = ""
            if c == 0 and r in (6, 7):
                vm = '<w:vMerge w:val="restart"/>' if r == 6 else "<w:vMerge/>"
            txt = "" if (c == 0 and r == 7) else (f"{ch}.{ti}.{r}" if c == 0 else rnd.choice(NOUNS))
            cells.append(f'<w:tc><w:tcPr><w:tcW w:w="{cw}" w:type="dxa"/>{vm}</w:tcPr>{p(run(txt, noise) if txt else "")}</w:tc>')
        trs.append("<w:tr>" + "".join(cells) + "</w:tr>")
    trs.append("<w:tr>" + "".join(f'<w:tc><w:tcPr><w:tcW w:w="{cw}" w:type="dxa"/></w:tcPr>{p(run("Total" if c == 0 else str(rnd.randint(1, 99))))}</w:tc>' for c in range(cols)) + "</w:tr>")
    note = "Note:" if ti % 2 else "Note: values are for the base period."
    trs.append(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{W_}" w:type="dxa"/><w:gridSpan w:val="{cols}"/></w:tcPr>{p(run(note))}</w:tc></w:tr>')
    tblpr = (f'<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="{W_}" w:type="dxa"/>'
             f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>')
    d.manifest["tables"] += 1
    return f"<w:tbl>{tblpr}<w:tblGrid>{grid}</w:tblGrid>{''.join(trs)}</w:tbl>" + p("")


def figure_xml(d, ch, fi, appendix, rid):
    pic = (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="914400"/>'
           f'<wp:docPr id="{d.bm + 5000}" name="Picture {d.bm}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
           f'<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/>'
           f'<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>'
           f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    name, cap = d.bookmark(caption_runs("Figure", "Figure", 1, appendix, None))
    d.ref_targets.append((name, "Figure"))
    title = f"{rnd.choice(NOUNS).title()} Flow"
    # two-part caption: title paragraph ends with a style separator (vanish + specVanish), then the action caption
    sep_ppr = '<w:rPr><w:vanish/><w:specVanish/></w:rPr>'
    cap_p1 = p(cap + run(".") + '<w:r><w:tab/></w:r>' + run(title), "Caption", sep_ppr)
    cap_p2 = p(run(" The " + rnd.choice(NOUNS) + " " + rnd.choice(VERBS) + " the schedule."), "Quote")
    d.manifest["captions"].append({"chapter": ch, "kind": "Figure", "appendix": appendix, "defect": None, "title": title})
    d.manifest["figures"] += 1
    pic_p = p(pic, None, '<w:jc w:val="center"/>')
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>{pic_p}</w:tc></w:tr>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>{cap_p1}{cap_p2}</w:tc></w:tr></w:tbl>' + p(""))


def sect(kind, hdr=True, title_pg=False, landscape=False):
    refs = '<w:headerReference w:type="default" r:id="rIdHdr1"/><w:footerReference w:type="default" r:id="rIdFtr1"/>' if hdr else ""
    if kind == "cover":
        refs = '<w:headerReference w:type="default" r:id="rIdHdr0"/><w:footerReference w:type="default" r:id="rIdFtr0"/>'
    sz = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>' if landscape else '<w:pgSz w:w="12240" w:h="15840"/>'
    return (f'<w:sectPr>{refs}<w:type w:val="nextPage"/>{sz}<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            f'w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/>{"<w:titlePg/>" if title_pg else ""}</w:sectPr>')


def build(chapters):
    d = Doc()
    B = d.body
    # cover section
    B.append(p(run("Example Program Technical Volume", "<w:b/><w:sz w:val=\"52\"/>"), "Title"))
    B.append(p(run("Submitted to Example Agency, Solicitation EX-26-R-0001")))
    B.append(p("", None, sect("cover", hdr=False)))
    d.manifest["sections"].append("cover")
    # front matter: TOC and list of tables
    B.append(p(run("Table of Contents"), "TOCHeading"))
    B.append(p(fld('TOC \\o "1-3" \\h \\z \\u', "Right-click to update the table of contents.")))
    B.append(p(run("List of Tables"), "TOCHeading"))
    B.append(p(fld('TOC \\h \\z \\c "Table"', "No table of figures entries found.")))
    B.append(p("", None, sect("front")))
    d.manifest["sections"].append("front")
    footnotes = []
    for ch in range(1, chapters + 1):
        appendix = False
        B.append(p(run(f"{rnd.choice(ADJ).title()} {rnd.choice(NOUNS).title()} Approach"), "Heading1"))
        for s in range(1, 4):
            B.append(p(run(f"{rnd.choice(NOUNS).title()} {s}"), "Heading2"))
            for k in range(3):
                txt = para()
                if k == 1 and s == 1:
                    # tracked insertion and deletion
                    B.append(p(run(txt[:200]) + f'<w:ins w:id="{90000 + ch}" w:author="Fixture" w:date="2026-10-01T00:00:00Z">{run(" Added by review.")}</w:ins>'
                               + f'<w:del w:id="{91000 + ch}" w:author="Fixture" w:date="2026-10-01T00:00:00Z"><w:r><w:delText xml:space="preserve"> Removed text.</w:delText></w:r></w:del>'))
                elif k == 2 and s == 2 and ch % 3 == 0:
                    fid = len(footnotes) + 1
                    footnotes.append(fid)
                    B.append(p(run(txt) + f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{fid}"/></w:r>'))
                else:
                    B.append(p(run(txt)))
            B.append(p(run(f"{rnd.choice(NOUNS).title()} detail"), "Heading3"))
            B.append(p(run(para(4))))
            defect = [None, "no_s", "mergeformat", "wrong_ident", None, "typed", None, None][(ch + s) % 8]
            B.append(table_xml(d, ch, s, appendix, defect))
            if d.ref_targets:
                name, lab = d.ref_targets[-1]
                B.append(p(run("As shown in ") + fld(f"REF {name} \\h", f"{lab} 1-1") + run(" on page ") + fld(f"PAGEREF {name} \\h", "1") + run(", the plan holds.")))
            if s != 2:
                B.append(figure_xml(d, ch, s, appendix, "rIdImg1"))
        if ch % 7 == 0 or ch == 4:
            missing = f"_Ref{900000000 + ch}"
            d.manifest["broken_refs"].append(missing)
            B.append(p(run("See ") + fld(f"REF {missing} \\h", "Table 9-9") + run(" for the earlier result.")))
        if ch % 10 == 5:
            B.append(p("", None, sect("chapter")))
            d.manifest["sections"].append("chapter")
            B.append(table_xml(d, ch, 9, appendix, None, cols=6, rows=8, landscape=True))
            B.append(p("", None, sect("landscape", landscape=True)))
            d.manifest["sections"].append("landscape")
            continue
        B.append(p("", None, sect("chapter")))
        d.manifest["sections"].append("chapter")
    for a in range(2):
        B.append(p(run(f"{rnd.choice(NOUNS).title()} Data"), "Heading7"))
        B.append(p(run(f"{rnd.choice(NOUNS).title()}"), "Heading8"))
        B.append(p(run(para())))
        B.append(table_xml(d, 100 + a, 1, True, None))
        B.append(table_xml(d, 100 + a, 2, True, None))
        last = a == 1
        if not last:
            B.append(p("", None, sect("appendix")))
            d.manifest["sections"].append("appendix")
    d.manifest["sections"].append("appendix")
    body = "".join(B) + sect("appendix")
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {WNS}><w:body>{body}</w:body></w:document>'
    return doc, d, footnotes


HDR = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {WNS}>'
       + p(fld("STYLEREF 1 \\* MERGEFORMAT", "Chapter") + run("  |  Example Program"), "Header") + '</w:hdr>')
FTR = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {WNS}>'
       + p(run("Example Agency") + '<w:r><w:tab/></w:r><w:r><w:tab/></w:r>' + fld("PAGE \\* roman \\* MERGEFORMAT", "i"), "Footer") + '</w:ftr>')
EMPTY_H = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {WNS}><w:p/></w:hdr>'
EMPTY_F = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {WNS}><w:p/></w:ftr>'


def footnotes_xml(ids):
    sep = ('<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
           '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>')
    notes = "".join(f'<w:footnote w:id="{i}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>{run(" Source: program records, note " + str(i) + ".")}</w:p></w:footnote>' for i in ids)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {WNS}>{sep}{notes}</w:footnotes>'


def main(tpl, logo, out, chapters, manifest_path):
    doc, d, fns = build(int(chapters))
    zin = zipfile.ZipFile(tpl)
    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for i in zin.infolist():
        data = zin.read(i)
        n = i.filename
        if n == "word/document.xml":
            data = doc.encode("utf-8")
        elif n == "[Content_Types].xml":
            data = data.decode().replace("wordprocessingml.template.main+xml", "wordprocessingml.document.main+xml")
            data = data.replace("</Types>",
                '<Default Extension="png" ContentType="image/png"/>'
                '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
                '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                '<Override PartName="/word/header0.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
                '<Override PartName="/word/footer0.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
                "</Types>").encode()
        elif n == "word/_rels/document.xml.rels":
            data = data.decode().replace("</Relationships>",
                '<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
                '<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
                '<Relationship Id="rIdHdr0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header0.xml"/>'
                '<Relationship Id="rIdFtr0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer0.xml"/>'
                '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
                '<Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>'
                "</Relationships>").encode()
        zout.writestr(n, data)
    zout.writestr("word/header1.xml", HDR)
    zout.writestr("word/footer1.xml", FTR)
    zout.writestr("word/header0.xml", EMPTY_H)
    zout.writestr("word/footer0.xml", EMPTY_F)
    zout.writestr("word/footnotes.xml", footnotes_xml(fns))
    zout.writestr("word/media/logo.png", open(logo, "rb").read())
    zout.close()
    json.dump(d.manifest, open(manifest_path, "w"), indent=1)
    print(out, "tables", d.manifest["tables"], "figures", d.manifest["figures"], "sections", len(d.manifest["sections"]),
          "captions", len(d.manifest["captions"]), "broken refs", len(d.manifest["broken_refs"]))


if __name__ == "__main__":
    main(*sys.argv[1:6])
