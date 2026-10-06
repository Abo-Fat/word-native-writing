from base64 import b64decode
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import sys
from lxml import etree

source = Path(sys.argv[1])
target = Path(sys.argv[2])
if source.resolve() == target.resolve() or target.exists():
    raise FileExistsError(f"Refusing to overwrite: {target}")
P = "{http://schemas.microsoft.com/office/2006/xmlPackage}"
C = "{http://schemas.openxmlformats.org/package/2006/content-types}"
package = etree.parse(str(source))
types = etree.Element(C + "Types", nsmap={None: C[1:-1]})
for extension, content_type in (
    ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
    ("xml", "application/xml"),
    ("png", "image/png"),
):
    etree.SubElement(types, C + "Default", Extension=extension, ContentType=content_type)
declared = {"rels", "xml", "png"}
for part in package.getroot():
    name = part.get(P + "name", "")
    extension = name.rsplit(".", 1)[-1].lower()
    content_type = part.get(P + "contentType")
    if part.find(P + "binaryData") is not None and extension not in declared and content_type:
        etree.SubElement(types, C + "Default", Extension=extension, ContentType=content_type)
        declared.add(extension)
with ZipFile(target, "w", ZIP_DEFLATED) as output:
    for part in package.getroot():
        name = part.get(P + "name")
        if not name:
            continue
        data = part.find(P + "xmlData")
        if data is None:
            content = b64decode(part.find(P + "binaryData").text)
        else:
            content = etree.tostring(
                data[0], xml_declaration=True, encoding="UTF-8", standalone=True
            )
        output.writestr(name.lstrip("/"), content)
        if not name.endswith(".rels"):
            etree.SubElement(
                types, C + "Override", PartName=name, ContentType=part.get(P + "contentType")
            )
    output.writestr(
        "[Content_Types].xml",
        etree.tostring(types, xml_declaration=True, encoding="UTF-8", standalone=True),
    )
print(f"Packaged {target.name}: {target.stat().st_size} bytes")
