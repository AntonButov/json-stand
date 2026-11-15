from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from json_stand.pipeline import convert_file


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="json-stand", description="Convert pdf-stand JSON to Excel")
    sub = parser.add_subparsers(dest="command", required=True)
    c = sub.add_parser("convert", help="JSON → xlsx")
    c.add_argument("file", type=Path)
    c.add_argument("--out", type=Path, default=None)
    c.add_argument("--no-normalize", action="store_true")
    c.add_argument("--base64", action="store_true")
    args = parser.parse_args(argv)

    if args.command == "convert":
        result = convert_file(
            args.file,
            out_path=args.out,
            normalize=not args.no_normalize,
            include_base64=args.base64,
        )
        payload = result.to_dict()
        if not args.base64:
            payload.pop("workbook_base64", None)
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0 if result.ok else 1
    return 2


if __name__ == "__main__":
    sys.exit(main())
