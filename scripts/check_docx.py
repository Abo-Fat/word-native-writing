"""Check field targets, native script characters and optional source invariants."""

import argparse
from collections import Counter
from hashlib import sha256
import json
from pathlib import Path
import re
import sys
import unicodedata
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from audit_docx import M, NS, W, complex_fields, script_characters, text


def visible_text(element):
    return "".join(
        node.text or ""
        for node in element.iter()
        if node.tag in (f"{{{W}}}t", f"{{{M}}}t")
    )


def read(path):
    with ZipFile(path) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        images = Counter(
            sha256(package.read(name)).hexdigest()
            for name in package.namelist()
            if name.startswith("word/media/") and not name.endswith("/")
        )
        cells = [
            [
                visible_text(cell)
                for row in table.findall("w:tr", NS)
                for cell in row.findall("w:tc", NS)
            ]
            for table in document.findall(".//w:tbl", NS)
        ]
        return document, images, cells


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--source", type=Path)
    args = parser.parse_args()
    document, images, cells = read(args.input)
    errors = []
    with ZipFile(args.input) as package:
        roots = [
            ET.fromstring(package.read(name))
            for name in package.namelist()
            if name == "word/document.xml"
            or re.fullmatch(r"word/(?:header\d*|footer\d*|footnotes|endnotes)\.xml", name)
        ]
    fields = [field for root in roots for field in complex_fields(root)]
    bookmark_names = [
        node.get(f"{{{W}}}name") for root in roots
        for node in root.findall(".//w:bookmarkStart", NS)
    ]
    bookmarks = set(bookmark_names)
    if len(bookmark_names) != len(bookmarks):
        errors.append("Duplicate bookmark names")
    for root in roots:
        starts = [node.get(f"{{{W}}}id") for node in root.findall(".//w:bookmarkStart", NS)]
        ends = [node.get(f"{{{W}}}id") for node in root.findall(".//w:bookmarkEnd", NS)]
        if len(starts) != len(set(starts)) or Counter(starts) != Counter(ends):
            errors.append("Duplicate or unmatched bookmark boundaries")
    for field in fields:
        match = re.match(r'^(REF|PAGEREF)\s+(?:"([^"]+)"|(\S+))', field["code"], re.I)
        if not match:
            continue
        target = match[2] or match[3]
        if target not in bookmarks:
            errors.append(f"Missing target: {target}")
        if any(
            marker in field["result"]
            for marker in (
                "Error! Reference source not found",
                "\u9519\u8bef\uff01\u672a\u627e\u5230\u5f15\u7528\u6e90",
                "\u9519\u8bef\uff01\u672a\u5b9a\u4e49\u4e66\u7b7e",
            )
        ):
            errors.append(f"Reference error result: {field['code']}")
    special = Counter(
        char for root in roots for char in script_characters(text(root) + text(root, "m"))
    )
    if special:
        errors.append(f"Review remaining Unicode script glyphs: {dict(special)}")
    if args.source:
        original, source_images, source_cells = read(args.source)
        if source_images != images:
            errors.append("Embedded image bytes changed")
        normalize = lambda tables: [
            [unicodedata.normalize("NFKC", cell) for cell in table] for table in tables
        ]
        if normalize(source_cells) != normalize(cells):
            errors.append("Table text differs after NFKC normalization")
        if len(original.findall(".//w:drawing", NS)) != len(document.findall(".//w:drawing", NS)):
            errors.append("Drawing count changed")
    summary = {
        "field_types": dict(Counter(f["code"].split()[0] for f in fields if f["code"])),
        "bookmarks": len(bookmarks),
        "tables": len(document.findall(".//w:tbl", NS)),
        "drawings": len(document.findall(".//w:drawing", NS)),
        "native_math": len(document.findall(".//m:oMath", NS)),
        "superscript_subscript_runs": dict(
            Counter(node.get(f"{{{W}}}val") for node in document.findall(".//w:vertAlign", NS))
        ),
        "errors": errors,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if errors:
        raise SystemExit(1)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
