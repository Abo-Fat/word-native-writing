"""Read-only DOCX audit. Output matches the case scripts' REPORT/TEMPLATE layout."""

import argparse
from collections import Counter
from hashlib import sha256
import json
from pathlib import Path
import sys
import unicodedata
from zipfile import ZipFile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}


def text(element, namespace="w"):
    return "".join(node.text or "" for node in element.findall(f".//{namespace}:t", NS))


def properties(element):
    if element is None:
        return {}
    return {
        child.tag.rsplit("}", 1)[-1]: {
            key.rsplit("}", 1)[-1]: value for key, value in child.attrib.items()
        }
        for child in element
    }


def complex_fields(root):
    fields = []
    stack = []
    for node in root.iter():
        if node.tag == f"{{{W}}}fldSimple":
            fields.append({"code": node.get(f"{{{W}}}instr", ""), "result": text(node)})
        elif node.tag == f"{{{W}}}fldChar":
            kind = node.get(f"{{{W}}}fldCharType")
            if kind == "begin":
                stack.append({"code": [], "result": [], "separated": False})
            elif kind == "separate":
                if not stack:
                    raise ValueError("Field separator without a field begin")
                stack[-1]["separated"] = True
            elif kind == "end":
                if not stack:
                    raise ValueError("Field end without a field begin")
                frame = stack.pop()
                record = {
                    "code": "".join(frame["code"]).strip(),
                    "result": "".join(frame["result"]),
                }
                fields.append(record)
                if stack and stack[-1]["separated"]:
                    stack[-1]["result"].append(record["result"])
        elif node.tag == f"{{{W}}}instrText" and stack and not stack[-1]["separated"]:
            stack[-1]["code"].append(node.text or "")
        elif node.tag == f"{{{W}}}t" and stack and stack[-1]["separated"]:
            stack[-1]["result"].append(node.text or "")
    if stack:
        raise ValueError("Unclosed complex Word fields")
    return fields


def script_characters(value):
    return [
        char
        for char in value
        if any(
            category in unicodedata.name(char, "")
            for category in ("SUPERSCRIPT", "SUBSCRIPT", "MODIFIER LETTER")
        )
    ]


def audit(path):
    with ZipFile(path) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        styles = ET.fromstring(package.read("word/styles.xml"))
        paragraphs = []
        for index, paragraph in enumerate(document.findall(".//w:p", NS), 1):
            style = paragraph.find("w:pPr/w:pStyle", NS)
            paragraphs.append(
                {
                    "index": index,
                    "text": text(paragraph),
                    "style": style.get(f"{{{W}}}val") if style is not None else "(default)",
                    "p": properties(paragraph.find("w:pPr", NS)),
                    "runs": [
                        {"text": text(run), "r": properties(run.find("w:rPr", NS))}
                        for run in paragraph.findall(".//w:r", NS)
                    ],
                    "math_text": text(paragraph, "m"),
                    "drawings": len(paragraph.findall(".//w:drawing", NS)),
                }
            )
        definitions = {}
        for style in styles.findall("w:style", NS):
            name = style.find("w:name", NS)
            definitions[style.get(f"{{{W}}}styleId")] = {
                "name": name.get(f"{{{W}}}val") if name is not None else "",
                "properties": properties(style),
                "p": properties(style.find("w:pPr", NS)),
                "r": properties(style.find("w:rPr", NS)),
            }
        defaults = styles.find("w:docDefaults", NS)
        fields = complex_fields(document)
        visible = text(document) + text(document, "m")
        return {
            "path": str(path.resolve()),
            "sha256": sha256(path.read_bytes()).hexdigest(),
            "paragraphs": paragraphs,
            "styles": definitions,
            "default_xml": ET.tostring(defaults, encoding="unicode") if defaults is not None else "",
            "fields": fields,
            "bookmarks": [
                node.get(f"{{{W}}}name")
                for node in document.findall(".//w:bookmarkStart", NS)
            ],
            "sections": [properties(node) for node in document.findall(".//w:sectPr", NS)],
            "counts": {
                "paragraphs": len(paragraphs),
                "tables": len(document.findall(".//w:tbl", NS)),
                "drawings": len(document.findall(".//w:drawing", NS)),
                "native_math": len(document.findall(".//m:oMath", NS)),
                "field_types": dict(Counter(f["code"].split()[0] for f in fields if f["code"])),
                "vert_align": dict(
                    Counter(
                        node.get(f"{{{W}}}val")
                        for node in document.findall(".//w:vertAlign", NS)
                    )
                ),
                "unicode_script_candidates": dict(Counter(script_characters(visible))),
            },
        }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--template", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    inputs = [args.input] + ([args.template] if args.template else [])
    if any(args.output.resolve() == path.resolve() for path in inputs):
        raise ValueError("Audit output must not replace an input document")
    result = {"REPORT": audit(args.input)}
    if args.template:
        result["TEMPLATE"] = audit(args.template)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    for label, data in result.items():
        print(label, json.dumps(data["counts"], ensure_ascii=False))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
