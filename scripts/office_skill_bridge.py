"""Invoke an installed office skill with the namespace-import workaround."""

import argparse
import importlib
from pathlib import Path
import runpy
import sys


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--office-dir", required=True, type=Path)
    parser.add_argument("operation", choices=("unpack", "pack", "validate"))
    parser.add_argument("arguments", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    script = args.office_dir / f"{args.operation}.py"
    if not script.is_file():
        raise FileNotFoundError(script)
    sys.path.insert(0, str(args.office_dir.resolve()))
    validators = importlib.import_module("validators")
    for module_name, class_name in (
        ("docx", "DOCXSchemaValidator"),
        ("pptx", "PPTXSchemaValidator"),
        ("redlining", "RedliningValidator"),
    ):
        module = importlib.import_module(f"validators.{module_name}")
        setattr(validators, class_name, getattr(module, class_name))
    sys.argv = [str(script)] + args.arguments
    runpy.run_path(str(script), run_name="__main__")


if __name__ == "__main__":
    main()
