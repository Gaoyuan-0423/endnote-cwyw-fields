#!/usr/bin/env python3
"""Read-only structural audit for EndNote CWYW fields in DOCX files."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from endnote_ooxml import EndNoteError, audit_document


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path, help="Word DOCX to audit")
    parser.add_argument(
        "--json", dest="json_path", type=Path, help="Optional JSON report path"
    )
    parser.add_argument(
        "--locations", action="store_true", help="Print every visible citation location"
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        report = audit_document(args.docx)
    except EndNoteError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    summary_keys = (
        "document",
        "visible_citation_groups",
        "active_endnote_cite_fields",
        "active_endnote_data_fields",
        "plain_citation_groups",
        "endnote_reflist_fields",
        "tracked_changes_present",
        "embedded_record_candidates",
        "embedded_unique_dois",
    )
    for key in summary_keys:
        print(f"{key}: {report[key]}")
    if report["mixed_paragraphs"]:
        print("mixed_paragraphs: " + ", ".join(map(str, report["mixed_paragraphs"])))
    if args.locations:
        for item in report["citation_locations"]:
            where = "table" if item["in_table"] else "body"
            print(
                f"p{item['paragraph']} {where} {item['display']} -> {item['numbers']}"
            )
    if args.json_path:
        args.json_path.parent.mkdir(parents=True, exist_ok=True)
        args.json_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    return 1 if report["plain_citation_groups"] or report["mixed_paragraphs"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
