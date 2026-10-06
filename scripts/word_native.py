"""Native Word structures shared by writing and explicit, non-destructive repair."""

from copy import deepcopy
import re

from lxml import etree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def element(tag, value=None, **attrs):
    prefix, local = tag.split(":")
    namespace = NS[prefix]
    node = ET.Element(f"{{{namespace}}}{local}")
    for key, item in attrs.items():
        node.set(f"{{{namespace}}}{key}", str(item))
    if value is not None:
        node.text = str(value)
    return node


def child(parent, tag, value=None, **attrs):
    node = element(tag, value, **attrs)
    parent.append(node)
    return node


def visible(paragraph):
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def run(text, properties=None, script=None):
    if not isinstance(text, str) or "\n" in text or "\r" in text:
        raise ValueError("Run text must be a string without paragraph breaks")
    node = element("w:r")
    if properties is not None:
        node.append(deepcopy(properties))
    if script is not None:
        if script not in ("superscript", "subscript"):
            raise ValueError(f"Unknown script: {script}")
        props = node.find("w:rPr", NS)
        if props is None:
            props = element("w:rPr")
            node.insert(0, props)
        for old in props.findall("w:vertAlign", NS):
            props.remove(old)
        later = {"rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"}
        index = next(
            (i for i, node in enumerate(props) if ET.QName(node).localname in later),
            len(props),
        )
        props.insert(index, element("w:vertAlign", val=script))
    child(node, "w:t", text).set(XML_SPACE, "preserve")
    return node


def field(code, cached, properties=None):
    if not isinstance(cached, str):
        raise ValueError("Field cached result must be a string")
    result = []
    for kind in ("begin", "instruction", "separate", "result", "end"):
        node = element("w:r")
        if properties is not None:
            node.append(deepcopy(properties))
        if kind == "instruction":
            child(node, "w:instrText", f" {code} ").set(XML_SPACE, "preserve")
        elif kind == "result":
            node = run(cached, properties)
        else:
            child(node, "w:fldChar", fldCharType=kind)
        result.append(node)
    return result


def identifier(value, bookmark=False):
    pattern = r"[A-Za-z][A-Za-z0-9_]{0,39}" if bookmark else r"[^\W\d]\w{0,39}"
    if not isinstance(value, str) or not re.fullmatch(pattern, value):
        raise ValueError(f"Invalid {'bookmark' if bookmark else 'sequence'} name: {value!r}")
    return value


class Bookmarks:
    def __init__(self, roots):
        starts = [node for root in roots for node in root.findall(".//w:bookmarkStart", NS)]
        self.names = {node.get(f"{{{W}}}name") for node in starts}
        ids = [
            int(node.get(f"{{{W}}}id"))
            for root in roots
            for node in root.xpath(".//w:bookmarkStart | .//w:bookmarkEnd", namespaces=NS)
        ]
        self.next_id = max(ids, default=-1) + 1

    def wrap(self, name, nodes):
        identifier(name, bookmark=True)
        if name in self.names:
            raise ValueError(f"Bookmark already exists: {name}")
        self.names.add(name)
        number = self.next_id
        self.next_id += 1
        return [
            element("w:bookmarkStart", id=number, name=name),
            *nodes,
            element("w:bookmarkEnd", id=number),
        ]


def math_nodes(spec):
    if isinstance(spec, str):
        node = element("m:r")
        child(node, "m:t", spec).set(XML_SPACE, "preserve")
        return [node]
    if isinstance(spec, list):
        return [node for item in spec for node in math_nodes(item)]
    if not isinstance(spec, dict) or len(spec) != 1:
        raise ValueError("Math must be text, a list, or one supported operator")
    operator, args = next(iter(spec.items()))
    structures = {
        "sub": ("m:sSub", ("m:e", "m:sub")),
        "sup": ("m:sSup", ("m:e", "m:sup")),
        "frac": ("m:f", ("m:num", "m:den")),
        "sum": ("m:nary", ("m:sub", "m:sup", "m:e")),
    }
    if operator == "roman":
        if not isinstance(args, str):
            raise ValueError("roman requires literal text")
        node = element("m:r")
        child(child(node, "m:rPr"), "m:sty", val="p")
        child(node, "m:t", args).set(XML_SPACE, "preserve")
        return [node]
    if operator in ("paren", "bracket", "abs"):
        node = element("m:d")
        props = child(node, "m:dPr")
        left, right = {"paren": ("(", ")"), "bracket": ("[", "]"), "abs": ("|", "|")}[operator]
        child(props, "m:begChr", val=left)
        child(props, "m:endChr", val=right)
        child(node, "m:e").extend(math_nodes(args))
        return [node]
    if operator not in structures:
        raise ValueError(f"Unsupported math operator: {operator}")
    tag, slots = structures[operator]
    if not isinstance(args, list) or len(args) != len(slots):
        raise ValueError(f"{operator} requires {len(slots)} arguments")
    node = element(tag)
    if operator == "sum":
        props = child(node, "m:naryPr")
        child(props, "m:chr", val="\u2211")
        child(props, "m:limLoc", val="undOvr")
        child(props, "m:subHide", val="1" if args[0] == "" else "0")
        child(props, "m:supHide", val="1" if args[1] == "" else "0")
    for slot, arg in zip(slots, args):
        child(node, slot).extend(math_nodes(arg))
    return [node]


def parts(items, properties=None):
    if not isinstance(items, list):
        raise ValueError("parts must be an array")
    nodes = []
    for item in items:
        if isinstance(item, str):
            nodes.append(run(item, properties))
        elif isinstance(item, dict) and "ref" in item:
            if set(item) - {"ref", "mode", "cached"}:
                raise ValueError(f"Unknown reference options: {item}")
            target = identifier(item["ref"], bookmark=True)
            mode = item.get("mode", "text")
            if mode not in ("text", "number", "page"):
                raise ValueError(f"Unknown reference mode: {mode}")
            code = f'PAGEREF {target} \\h' if mode == "page" else f'REF {target} \\h'
            if mode == "number":
                code += " \\n"
            nodes.extend(field(code, item["cached"], properties))
        elif isinstance(item, dict) and set(item) == {"math"}:
            node = element("m:oMath")
            node.extend(math_nodes(item["math"]))
            nodes.append(node)
        elif isinstance(item, dict) and "text" in item:
            if set(item) - {"text", "script"}:
                raise ValueError(f"Unknown run options: {item}")
            nodes.append(run(item["text"], properties, item.get("script")))
        else:
            raise ValueError(f"Unsupported inline item: {item!r}")
    return nodes


def plain_runs(paragraph):
    """Only edit plain runs; do not flatten fields, drawings, links or revisions."""
    if any(node.tag not in (f"{{{W}}}pPr", f"{{{W}}}r") for node in paragraph):
        raise ValueError("Paragraph has non-plain content; use a targeted XML/Word edit")
    runs = paragraph.findall("w:r", NS)
    if any(
        any(node.tag not in (f"{{{W}}}rPr", f"{{{W}}}t") for node in item)
        for item in runs
    ):
        raise ValueError("Paragraph has fields, tabs, breaks or drawings; cannot replace text")
    return runs


def replace_span(paragraph, start, end, replacement, preflight=True):
    runs = plain_runs(paragraph) if preflight else paragraph.findall("w:r", NS)
    if not 0 <= start < end <= len(visible(paragraph)):
        raise ValueError("Invalid replacement span")
    position = 0
    inserted = False
    for original in runs:
        value = "".join(original.xpath("w:t/text()", namespaces=NS))
        finish = position + len(value)
        if position < end and finish > start:
            props = original.find("w:rPr", NS)
            new = []
            if start > position:
                new.append(run(value[: start - position], props))
            if not inserted:
                new.extend(replacement(props))
                inserted = True
            if finish > end:
                new.append(run(value[end - position :], props))
            index = paragraph.index(original)
            paragraph.remove(original)
            for offset, node in enumerate(new):
                paragraph.insert(index + offset, node)
        position = finish


def literal_span(paragraph, value, occurrence=None):
    if not value:
        raise ValueError("Empty match is not allowed")
    matches = list(re.finditer(re.escape(value), visible(paragraph)))
    if occurrence is None and len(matches) != 1:
        raise ValueError(f"Expected exactly one match for {value!r}, found {len(matches)}")
    index = 1 if occurrence is None else occurrence
    if not isinstance(index, int) or index < 1 or index > len(matches):
        raise ValueError(f"Invalid occurrence for {value!r}: {index}")
    match = matches[index - 1]
    return match.start(), match.end()


def paragraph_properties(paragraph):
    props = paragraph.find("w:pPr", NS)
    if props is None:
        props = element("w:pPr")
        paragraph.insert(0, props)
    return props


def set_style(paragraph, style, numbering=None, level=0):
    props = paragraph_properties(paragraph)
    for node in props.xpath("w:pStyle | w:numPr | w:outlineLvl", namespaces=NS):
        props.remove(node)
    props.insert(0, element("w:pStyle", val=style))
    if numbering is not None:
        num = element("w:numPr")
        child(num, "w:ilvl", val=level)
        child(num, "w:numId", val=numbering)
        # numPr precedes spacing, indentation, alignment and paragraph-mark rPr.
        later = {
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
            "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
            "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
            "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
            "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
        }
        index = next(
            (i for i, node in enumerate(props) if ET.QName(node).localname in later),
            len(props),
        )
        props.insert(index, num)


def add_numbering(root, headings=False):
    abstract_id = max(
        (int(node.get(f"{{{W}}}abstractNumId")) for node in root.findall("w:abstractNum", NS)),
        default=-1,
    ) + 1
    num_id = max(
        (int(node.get(f"{{{W}}}numId")) for node in root.findall("w:num", NS)),
        default=0,
    ) + 1
    abstract = element("w:abstractNum", abstractNumId=abstract_id)
    child(abstract, "w:multiLevelType", val="multilevel" if headings else "singleLevel")
    for level in range(9 if headings else 1):
        node = child(abstract, "w:lvl", ilvl=level)
        child(node, "w:start", val=1)
        child(node, "w:numFmt", val="decimal")
        if headings:
            child(node, "w:pStyle", val=f"Heading{level + 1}")
        child(node, "w:suff", val="space")
        child(node, "w:lvlText", val=".".join(f"%{i + 1}" for i in range(level + 1)))
        child(node, "w:lvlJc", val="left")
    index = next(
        (i for i, node in enumerate(root) if node.tag in (f"{{{W}}}num", f"{{{W}}}numIdMacAtCleanup")),
        len(root),
    )
    root.insert(index, abstract)
    num = element("w:num", numId=num_id)
    child(num, "w:abstractNumId", val=abstract_id)
    cleanup = root.find("w:numIdMacAtCleanup", NS)
    root.insert(root.index(cleanup) if cleanup is not None else len(root), num)
    return num_id


def ensure_styles(root):
    for level in range(1, 10):
        name = f"Heading{level}"
        style = next((s for s in root if s.get(f"{{{W}}}styleId") == name), None)
        if style is None:
            style = child(root, "w:style", type="paragraph", styleId=name)
            child(style, "w:name", val=f"heading {level}")
            child(style, "w:basedOn", val="Normal")
            child(style, "w:next", val="Normal")
            child(style, "w:qFormat")
            fonts = child(style, "w:rPr")
            child(fonts, "w:b")
            child(fonts, "w:sz", val=max(24, 34 - level * 2))
        props = style.find("w:pPr", NS)
        if props is None:
            props = element("w:pPr")
            run_props = style.find("w:rPr", NS)
            style.insert(style.index(run_props) if run_props is not None else len(style), props)
        outline = props.find("w:outlineLvl", NS)
        if outline is None:
            outline = element("w:outlineLvl")
            index = next(
                (i for i, node in enumerate(props)
                 if ET.QName(node).localname in {"divId", "cnfStyle", "rPr", "sectPr", "pPrChange"}),
                len(props),
            )
            props.insert(index, outline)
        outline.set(f"{{{W}}}val", str(level - 1))


def caption_nodes(item, bookmarks, properties=None, equation=False):
    sequence = identifier(item.get("sequence", "Equation" if equation else item["label"]))
    cached = str(item["cached"])
    chapter = item.get("chapter", False)
    if chapter:
        if not re.fullmatch(r"[1-9]\d*\.[1-9]\d*", cached):
            raise ValueError("Chapter caption cached number must be chapter.sequence")
        first, second = cached.split(".")
        numbered = field("STYLEREF 1 \\s", first, properties) + [run(".", properties)]
        numbered += field(f"SEQ {sequence} \\* ARABIC \\s 1", second, properties)
    else:
        if not re.fullmatch(r"[1-9]\d*", cached):
            raise ValueError("Caption cached number must be a positive integer")
        numbered = field(f"SEQ {sequence} \\* ARABIC", cached, properties)
    if equation:
        numbered = [run("(", properties), *numbered, run(")", properties)]
    else:
        numbered = [run(item["label"], properties), *numbered]
    return bookmarks.wrap(item["bookmark"], numbered)


def validate_roots(roots):
    from audit_docx import complex_fields

    names = set()
    fields = []
    for root in roots:
        starts = root.findall(".//w:bookmarkStart", NS)
        ends = root.findall(".//w:bookmarkEnd", NS)
        start_ids = [node.get(f"{{{W}}}id") for node in starts]
        end_ids = [node.get(f"{{{W}}}id") for node in ends]
        if len(start_ids) != len(set(start_ids)) or sorted(start_ids) != sorted(end_ids):
            raise ValueError("Duplicate or unmatched bookmark boundaries")
        for node in starts:
            name = node.get(f"{{{W}}}name")
            if name in names:
                raise ValueError(f"Duplicate bookmark name: {name}")
            names.add(name)
        fields.extend(complex_fields(root))
    for record in fields:
        match = re.match(r'^(REF|PAGEREF)\s+(?:"([^"]+)"|(\S+))', record["code"], re.I)
        if match and (match[2] or match[3]) not in names:
            raise ValueError(f"Missing reference target: {record['code']}")


def xml_bytes(root):
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
