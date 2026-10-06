"""Write semantic DOCX from JSON, or repair a DOCX using an explicit JSON plan."""

import argparse
from collections import defaultdict
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import sys
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as ET

from word_native import (
    Bookmarks, M, NS, W, add_numbering, caption_nodes, child, element,
    ensure_styles, field, literal_span, math_nodes, parts, plain_runs,
    replace_span, run, set_style, validate_roots, visible, xml_bytes,
)

C = "http://schemas.openxmlformats.org/package/2006/content-types"
R = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def parse(data):
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))


def story_roots(package):
    return [
        parse(data) for name, data in package.items()
        if name == "word/document.xml"
        or (
            name.startswith("word/") and name.endswith(".xml")
            and Path(name).stem.startswith(("header", "footer", "footnotes", "endnotes"))
        )
    ]


def save(package, output):
    for name, data in package.items():
        if name.endswith((".xml", ".rels")):
            parse(data)
    validate_roots(story_roots(package))
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("xb") as stream, ZipFile(stream, "w", ZIP_DEFLATED) as archive:
        for name, data in package.items():
            archive.writestr(name, data)


def numbered_heading(paragraph, item, heading_num):
    level = item["level"]
    if not isinstance(level, int) or not 1 <= level <= 9:
        raise ValueError("Heading level must be 1..9")
    set_style(
        paragraph, f"Heading{level}",
        heading_num if item.get("numbered", True) else 0, level - 1,
    )


def bookmark_paragraph(paragraph, name, bookmarks):
    content = [node for node in paragraph if node.tag != f"{{{W}}}pPr"]
    for node in content:
        paragraph.remove(node)
    paragraph.extend(bookmarks.wrap(name, content))


def new_package(spec):
    if set(spec) - {"title", "blocks"}:
        raise ValueError("Document spec supports only title and blocks")
    document = ET.Element(f"{{{W}}}document", nsmap={"w": W, "m": M})
    body = child(document, "w:body")
    styles = ET.Element(f"{{{W}}}styles", nsmap={"w": W})
    normal = child(styles, "w:style", type="paragraph", default="1", styleId="Normal")
    child(normal, "w:name", val="Normal")
    props = child(normal, "w:rPr")
    child(props, "w:rFonts", ascii="Times New Roman", hAnsi="Times New Roman", eastAsia="\u5b8b\u4f53")
    child(props, "w:sz", val=24)
    ensure_styles(styles)
    numbering = ET.Element(f"{{{W}}}numbering", nsmap={"w": W})
    heading_num = add_numbering(numbering, headings=True)
    bibliography_num = add_numbering(numbering)
    bookmarks = Bookmarks([document])

    def paragraph(item, parent):
        kind = item["kind"]
        allowed = {
            "paragraph": {"parts", "bookmark"},
            "heading": {"parts", "bookmark", "level", "numbered"},
            "bibliography": {"parts", "bookmark"},
            "caption": {"parts", "bookmark", "label", "sequence", "cached", "chapter"},
            "equation": {"math", "bookmark", "sequence", "cached", "chapter"},
            "toc": {"mode", "sequence"},
        }
        if kind not in allowed or set(item) - (allowed[kind] | {"kind"}):
            raise ValueError(f"Unsupported block or options: {item}")
        node = child(parent, "w:p")
        if kind in ("paragraph", "heading", "bibliography"):
            node.extend(parts(item["parts"]))
            if kind == "heading":
                numbered_heading(node, item, heading_num)
            elif kind == "bibliography":
                set_style(node, "Normal", bibliography_num)
            if "bookmark" in item:
                bookmark_paragraph(node, item["bookmark"], bookmarks)
        elif kind == "caption":
            node.extend(caption_nodes(item, bookmarks))
            node.extend(parts(item.get("parts", [])))
        elif kind == "equation":
            math = child(node, "m:oMath")
            math.extend(math_nodes(item["math"]))
            node.append(run("  "))
            node.extend(caption_nodes(item, bookmarks, equation=True))
        elif kind == "toc":
            mode = item.get("mode", "headings")
            if mode == "headings":
                code = 'TOC \\o "1-3" \\h \\z \\u'
            elif mode in ("figures", "tables"):
                sequence = item["sequence"]
                from word_native import identifier
                code = f'TOC \\h \\z \\c "{identifier(sequence)}"'
            else:
                raise ValueError(f"Unknown TOC mode: {mode}")
            node.extend(field(code, "\u76ee\u5f55\u5f85 Word \u66f4\u65b0"))
        else:
            raise ValueError(f"Unsupported block kind: {kind}")

    if spec.get("title"):
        title = child(body, "w:p")
        title.append(run(spec["title"]))
    for item in spec["blocks"]:
        if item["kind"] != "table":
            paragraph(item, body)
            continue
        if set(item) != {"kind", "rows"}:
            raise ValueError("Table block supports only kind and rows")
        rows = item["rows"]
        if not rows or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
            raise ValueError("Table must be a nonempty rectangular array of parts arrays")
        table = child(body, "w:tbl")
        props = child(table, "w:tblPr")
        width = 9026
        widths = [width // len(rows[0])] * len(rows[0])
        widths[-1] += width - sum(widths)
        child(props, "w:tblW", w=width, type="dxa")
        grid = child(table, "w:tblGrid")
        for size in widths:
            child(grid, "w:gridCol", w=size)
        for row in rows:
            tr = child(table, "w:tr")
            for size, value in zip(widths, row):
                cell = child(tr, "w:tc")
                child(child(cell, "w:tcPr"), "w:tcW", w=size, type="dxa")
                paragraph({"kind": "paragraph", "parts": value}, cell)
    section = child(body, "w:sectPr")
    child(section, "w:pgSz", w=11906, h=16838)
    child(section, "w:pgMar", top=1440, bottom=1440, left=1440, right=1440, header=720, footer=720, gutter=0)
    settings = ET.Element(f"{{{W}}}settings", nsmap={"w": W})
    child(settings, "w:updateFields", val="false")
    types = ET.Element(f"{{{C}}}Types", nsmap={None: C})
    for ext, mime in (("rels", "application/vnd.openxmlformats-package.relationships+xml"), ("xml", "application/xml")):
        ET.SubElement(types, f"{{{C}}}Default", Extension=ext, ContentType=mime)
    for name, suffix in (
        ("document", "document.main"), ("styles", "styles"),
        ("numbering", "numbering"), ("settings", "settings"),
    ):
        ET.SubElement(types, f"{{{C}}}Override", PartName=f"/word/{name}.xml",
                      ContentType=f"application/vnd.openxmlformats-officedocument.wordprocessingml.{suffix}+xml")
    rels = ET.Element(f"{{{R}}}Relationships", nsmap={None: R})
    ET.SubElement(rels, f"{{{R}}}Relationship", Id="rId1", Type=OFFICE + "/officeDocument", Target="word/document.xml")
    document_rels = ET.Element(f"{{{R}}}Relationships", nsmap={None: R})
    for index, name in enumerate(("styles", "numbering", "settings"), 1):
        ET.SubElement(document_rels, f"{{{R}}}Relationship", Id=f"rId{index}", Type=OFFICE + "/" + name, Target=name + ".xml")
    return {
        "[Content_Types].xml": xml_bytes(types), "_rels/.rels": xml_bytes(rels),
        "word/_rels/document.xml.rels": xml_bytes(document_rels),
        "word/document.xml": xml_bytes(document), "word/styles.xml": xml_bytes(styles),
        "word/numbering.xml": xml_bytes(numbering), "word/settings.xml": xml_bytes(settings),
    }


def repair(package, plan):
    if set(plan) - {"source_sha256", "operations"}:
        raise ValueError("Repair plan supports only source_sha256 and operations")
    document = parse(package["word/document.xml"])
    styles = parse(package["word/styles.xml"])
    document_rels = parse(package.get(
        "word/_rels/document.xml.rels",
        xml_bytes(ET.Element(f"{{{R}}}Relationships", nsmap={None: R})),
    ))
    if any(
        node.get("Type") == OFFICE + "/numbering" and node.get("Target") != "numbering.xml"
        for node in document_rels
    ):
        raise ValueError("Nonstandard numbering part target; use a targeted package edit")
    numbering = parse(package.get("word/numbering.xml", xml_bytes(element("w:numbering"))))
    paragraphs = document.findall(".//w:p", NS)
    bookmarks = Bookmarks(story_roots(package))
    edits = defaultdict(list)
    structural = {}
    heading_num = None
    bibliography_num = None
    for item in plan["operations"]:
        kind = item["kind"]
        allowed = {
            "heading": {"level", "numbered", "prefix", "bookmark"},
            "bibliography": {"prefix", "bookmark"},
            "caption": {"prefix", "bookmark", "label", "sequence", "cached", "chapter"},
            "equation": {"math", "bookmark", "sequence", "cached", "chapter"},
            "script": {"match", "occurrence", "text", "script"},
            "replace": {"match", "occurrence", "parts"},
        }
        if kind not in allowed or set(item) - (allowed[kind] | {"kind", "paragraph", "expect"}):
            raise ValueError(f"Unsupported operation or options: {item}")
        number = item["paragraph"]
        if not isinstance(number, int) or not 1 <= number <= len(paragraphs):
            raise ValueError(f"Invalid paragraph index: {number}")
        node = paragraphs[number - 1]
        if visible(node) != item["expect"]:
            raise ValueError(f"Paragraph {number} no longer matches expect; re-audit input")
        if kind in ("heading", "bibliography", "caption", "equation"):
            if number in structural:
                raise ValueError(f"Multiple structural operations on paragraph {number}")
            structural[number] = item
        if kind in ("heading", "bibliography"):
            plain_runs(node)
            prefix = item.get("prefix", "")
            if prefix:
                if not visible(node).startswith(prefix):
                    raise ValueError(f"Paragraph {number} does not start with prefix")
                edits[number].append((0, len(prefix), lambda props: []))
        elif kind == "caption":
            prefix = item["prefix"]
            if not prefix or not visible(node).startswith(prefix):
                raise ValueError(f"Paragraph {number} does not start with caption prefix")
            edits[number].append(
                (0, len(prefix), lambda props, i=item: caption_nodes(i, bookmarks, props))
            )
        elif kind == "equation":
            plain_runs(node)
        elif kind in ("script", "replace"):
            start, end = literal_span(node, item["match"], item.get("occurrence"))
            if kind == "script":
                if item["script"] not in ("superscript", "subscript"):
                    raise ValueError("script must be superscript or subscript")
                replacement = lambda props, i=item: [run(i.get("text", i["match"]), props, i["script"])]
            else:
                replacement = lambda props, i=item: parts(i["parts"], props)
            edits[number].append((start, end, replacement))
        else:
            raise ValueError(f"Unknown repair operation: {kind}")
    for number, changes in edits.items():
        if structural.get(number, {}).get("kind") == "equation":
            raise ValueError("Equation replacement cannot be combined with text edits")
        ordered = sorted(changes, key=lambda change: change[:2])
        if any(left[1] > right[0] for left, right in zip(ordered, ordered[1:])):
            raise ValueError(f"Overlapping edits in paragraph {number}")
        # Preflight once; later edits may introduce fields or bookmarks.
        node = paragraphs[number - 1]
        plain_runs(node)
        for start, end, replacement in reversed(ordered):
            replace_span(node, start, end, replacement, preflight=False)
    for number, item in structural.items():
        node = paragraphs[number - 1]
        kind = item["kind"]
        if kind == "heading":
            if heading_num is None:
                ensure_styles(styles)
                heading_num = add_numbering(numbering, headings=True)
            numbered_heading(node, item, heading_num)
        elif kind == "bibliography":
            if bibliography_num is None:
                bibliography_num = add_numbering(numbering)
            style = node.find("w:pPr/w:pStyle", NS)
            set_style(node, style.get(f"{{{W}}}val") if style is not None else "Normal", bibliography_num)
        elif kind == "equation":
            for old in list(node):
                if old.tag != f"{{{W}}}pPr":
                    node.remove(old)
            child(node, "m:oMath").extend(math_nodes(item["math"]))
            node.append(run("  "))
            node.extend(caption_nodes(item, bookmarks, equation=True))
        if kind in ("heading", "bibliography") and "bookmark" in item:
            bookmark_paragraph(node, item["bookmark"], bookmarks)
    updated = dict(package)
    updated["word/document.xml"] = xml_bytes(document)
    if heading_num is not None:
        updated["word/styles.xml"] = xml_bytes(styles)
    if heading_num is not None or bibliography_num is not None:
        updated["word/numbering.xml"] = xml_bytes(numbering)
        rels = document_rels
        if not any(node.get("Type") == OFFICE + "/numbering" for node in rels):
            ids = {node.get("Id") for node in rels}
            index = 1
            while f"rId{index}" in ids:
                index += 1
            ET.SubElement(rels, f"{{{R}}}Relationship", Id=f"rId{index}", Type=OFFICE + "/numbering", Target="numbering.xml")
            updated["word/_rels/document.xml.rels"] = xml_bytes(rels)
        types = parse(updated["[Content_Types].xml"])
        if not any(node.get("PartName") == "/word/numbering.xml" for node in types):
            ET.SubElement(types, f"{{{C}}}Override", PartName="/word/numbering.xml",
                          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")
            updated["[Content_Types].xml"] = xml_bytes(types)
    if "word/settings.xml" in updated:
        settings = parse(updated["word/settings.xml"])
        update = settings.find("w:updateFields", NS)
        if update is not None:
            update.set(f"{{{W}}}val", "false")
            updated["word/settings.xml"] = xml_bytes(settings)
    return updated


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    write = commands.add_parser("write")
    write.add_argument("spec", type=Path)
    write.add_argument("output", type=Path)
    fix = commands.add_parser("repair")
    fix.add_argument("input", type=Path)
    fix.add_argument("plan", type=Path)
    fix.add_argument("output", type=Path)
    args = parser.parse_args()
    if args.output.suffix.lower() != ".docx":
        raise ValueError("Output must be a .docx file")
    inputs = [args.spec] if args.command == "write" else [args.input, args.plan]
    if args.output.exists() or any(args.output.resolve() == item.resolve() for item in inputs):
        raise FileExistsError("Output must be a new file, not an input")
    if args.command == "write":
        spec = json.loads(args.spec.read_text(encoding="utf-8-sig"))
        package = new_package(spec)
    else:
        if args.input.suffix.lower() != ".docx":
            raise ValueError("Repair input must be a .docx file")
        plan = json.loads(args.plan.read_text(encoding="utf-8-sig"))
        source = args.input.read_bytes()
        if plan.get("source_sha256") != sha256(source).hexdigest():
            raise ValueError("source_sha256 must match the audited input file")
        with ZipFile(BytesIO(source)) as archive:
            if len(archive.namelist()) != len(set(archive.namelist())):
                raise ValueError("Duplicate package part names; cannot preserve this input safely")
            if any(name.startswith("_xmlsignatures/") for name in archive.namelist()):
                raise ValueError("Signed documents require an explicit signature handling decision")
            package = {name: archive.read(name) for name in archive.namelist()}
        package = repair(package, plan)
    save(package, args.output)
    print(f"Created: {args.output}; field caches still require Word verification")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
