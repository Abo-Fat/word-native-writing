"""Unpack or pack a DOCX without converting its content. No XSD validation."""

import argparse
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET


def unpack(source, destination):
    if destination.exists():
        raise FileExistsError(f"Use a new unpack directory: {destination}")
    with ZipFile(source) as package:
        base = destination.resolve()
        for entry in package.infolist():
            target = (base / entry.filename).resolve()
            if target != base and base not in target.parents:
                raise ValueError(f"Unsafe package path: {entry.filename}")
        destination.mkdir(parents=True)
        package.extractall(destination)


def pack(source, destination):
    if destination.exists():
        raise FileExistsError(f"Refusing to overwrite: {destination}")
    files = sorted(path for path in source.rglob("*") if path.is_file())
    if not (source / "[Content_Types].xml").is_file():
        raise ValueError("Not an unpacked Office package")
    for path in files:
        if path.suffix.lower() in (".xml", ".rels"):
            ET.parse(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(destination, "w", ZIP_DEFLATED) as package:
        for path in files:
            package.write(path, path.relative_to(source).as_posix())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=("unpack", "pack"))
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    args = parser.parse_args()
    if args.operation == "unpack":
        unpack(args.source, args.destination)
    else:
        pack(args.source, args.destination)
    print(f"{args.operation}: {args.destination}")


if __name__ == "__main__":
    main()
