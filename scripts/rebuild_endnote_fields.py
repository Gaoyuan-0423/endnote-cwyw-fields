#!/usr/bin/env python3
"""Plan or rebuild plain numeric DOCX citations as EndNote CWYW fields."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from endnote_ooxml import (
    EndNoteError,
    audit_document,
    collect_occurrences,
    extract_embedded_records,
    flatten_endnote_citation_fields,
    has_tracked_changes,
    load_document_xml,
    parse_bibliography,
    parse_reference_map,
    rebuild_plain_citations,
    records_from_library,
    resolve_records,
    write_docx_with_document_xml,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "docx", type=Path, help="Clean Word DOCX containing numeric citations"
    )
    parser.add_argument(
        "--source-docx",
        action="append",
        type=Path,
        default=[],
        help="DOCX containing original EndNote traveling-library records; repeatable",
    )
    parser.add_argument(
        "--library",
        action="append",
        type=Path,
        default=[],
        help="Read-only EndNote .enl library; repeatable and ordered by preference",
    )
    parser.add_argument(
        "--reference-map",
        type=Path,
        help="Optional JSON mapping from bibliography number to DOI",
    )
    parser.add_argument(
        "--output", type=Path, help="New output DOCX; omit for dry-run planning"
    )
    parser.add_argument("--report", type=Path, help="Optional JSON plan/result report")
    return parser.parse_args()


def emit_report(report: dict, path: Path | None) -> None:
    text = json.dumps(report, ensure_ascii=False, indent=2)
    print(text)
    if path:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    report: dict = {
        "input": str(args.docx.resolve()),
        "mode": "write" if args.output else "dry-run",
        "record_sources": [],
        "issues": [],
    }
    try:
        root, _ = load_document_xml(args.docx)
        if has_tracked_changes(root):
            raise EndNoteError(
                "Tracked changes are present; rebuild an accepted/clean copy instead"
            )

        target_audit = audit_document(args.docx)
        occurrences = collect_occurrences(root)
        cited_numbers = [number for item in occurrences for number in item.numbers]
        ref_map = parse_reference_map(args.reference_map)
        references = parse_bibliography(root, ref_map)

        candidates = extract_embedded_records(args.docx)
        report["record_sources"].append(
            {
                "source": str(args.docx.resolve()),
                "type": "target_docx",
                "records": len(candidates),
            }
        )
        for path in args.source_docx:
            records = extract_embedded_records(path)
            candidates.extend(records)
            report["record_sources"].append(
                {
                    "source": str(path.resolve()),
                    "type": "source_docx",
                    "records": len(records),
                }
            )
        for path in args.library:
            records = records_from_library(path)
            candidates.extend(records)
            report["record_sources"].append(
                {
                    "source": str(path.resolve()),
                    "type": "endnote_library",
                    "records": len(records),
                }
            )

        resolved, issues = resolve_records(references, cited_numbers, candidates)
        report.update(
            {
                "visible_citation_groups": len(occurrences),
                "active_citation_fields_before": target_audit[
                    "active_endnote_cite_fields"
                ],
                "plain_citation_groups_before": target_audit[
                    "plain_citation_groups"
                ],
                "unique_cited_references": len(set(cited_numbers)),
                "resolved_references": len(resolved),
                "issues": issues,
            }
        )
        if issues:
            report["status"] = "stopped_without_writing"
            emit_report(report, args.report)
            return 3
        if not args.output:
            report["status"] = "ready_to_write"
            emit_report(report, args.report)
            return 0

        flattened, flatten_issues = flatten_endnote_citation_fields(root)
        report["existing_fields_replaced"] = flattened
        report["issues"].extend(flatten_issues)
        if flatten_issues:
            report["status"] = "stopped_without_writing"
            emit_report(report, args.report)
            return 4
        conversions, conversion_issues = rebuild_plain_citations(root, resolved)
        report["issues"].extend(conversion_issues)
        if conversion_issues:
            report["status"] = "stopped_without_writing"
            emit_report(report, args.report)
            return 4
        write_docx_with_document_xml(args.docx, args.output, root)
        after = audit_document(args.output)
        report.update(
            {
                "output": str(args.output.resolve()),
                "converted_plain_groups": conversions,
                "audit_after": after,
            }
        )
        if after["plain_citation_groups"] != 0:
            raise EndNoteError(
                f"Post-write audit found {after['plain_citation_groups']} plain citation groups"
            )
        if (
            after["active_endnote_cite_fields"]
            != after["active_endnote_data_fields"]
        ):
            raise EndNoteError("Post-write EN.CITE and EN.CITE.DATA counts differ")
        if (
            after["active_endnote_cite_fields"]
            != after["visible_citation_groups"]
        ):
            raise EndNoteError(
                "Post-write active-field count does not equal visible citation-group count"
            )
        report["status"] = "structurally_rebuilt_requires_word_endnote_round_trip"
        emit_report(report, args.report)
        return 0
    except (EndNoteError, OSError, ValueError, json.JSONDecodeError) as exc:
        report["issues"].append({"issue": str(exc)})
        report["status"] = "stopped_without_writing"
        emit_report(report, args.report)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
