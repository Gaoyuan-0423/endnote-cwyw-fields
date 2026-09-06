#!/usr/bin/env python3
"""Shared, fail-closed helpers for EndNote CWYW DOCX fields."""

from __future__ import annotations

import base64
import copy
import json
import re
import shutil
import sqlite3
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Iterator, Sequence

try:
    from lxml import etree
except ImportError as exc:  # pragma: no cover - environment diagnostic
    raise SystemExit("lxml is required: install it in the active Python environment") from exc


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W = f"{{{W_NS}}}"
XML_SPACE = f"{{{XML_NS}}}space"
NS = {"w": W_NS}

CITATION_RE = re.compile(
    r"\[(?:\s*\d+\s*(?:(?:,|;|[-\u2013\u2014])\s*\d+\s*)*)\]"
)
DOI_RE = re.compile(r"(?:doi\s*:\s*|https?://(?:dx\.)?doi\.org/)(10\.\d{4,9}/\S+)", re.I)
INSTR_CITE_RE = re.compile(r"^ADDIN\s+EN\.CITE$")
INSTR_DATA_RE = re.compile(r"^ADDIN\s+EN\.CITE\.DATA$")
INSTR_REFLIST_RE = re.compile(r"^ADDIN\s+EN\.REFLIST$")


class EndNoteError(RuntimeError):
    """An integrity condition that requires the caller to stop."""


@dataclass(frozen=True)
class ReferenceEntry:
    number: int
    text: str
    doi: str | None


@dataclass
class RecordCandidate:
    record: etree._Element
    source: str
    doi: str | None
    title_key: str | None


@dataclass(frozen=True)
class CitationOccurrence:
    paragraph_index: int
    display: str
    numbers: tuple[int, ...]
    in_table: bool


def normalize_instruction(value: str) -> str:
    return " ".join(value.split())


def normalize_doi(value: str | None) -> str | None:
    if not value:
        return None
    value = value.strip()
    value = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", value, flags=re.I)
    value = re.sub(r"^doi\s*:\s*", "", value, flags=re.I)
    value = value.rstrip(".,;:)]}\u0000").lower()
    return value or None


def normalize_title(value: str | None) -> str | None:
    if not value:
        return None
    value = re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()
    return value or None


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def paragraph_style(paragraph: etree._Element) -> str:
    nodes = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return nodes[0] if nodes else ""


def paragraph_in_table(paragraph: etree._Element) -> bool:
    return bool(paragraph.xpath("ancestor::w:tbl", namespaces=NS))


def field_instructions(paragraph: etree._Element) -> list[str]:
    values = paragraph.xpath(".//w:instrText/text()", namespaces=NS)
    return [normalize_instruction(value) for value in values]


def is_reference_heading(text: str) -> bool:
    key = re.sub(r"\s+", " ", text).strip().casefold().rstrip(":")
    return key in {"references", "reference list", "bibliography"}


def load_document_xml(path: Path) -> tuple[etree._Element, bytes]:
    if not path.is_file():
        raise EndNoteError(f"DOCX not found: {path}")
    try:
        with zipfile.ZipFile(path) as archive:
            xml_bytes = archive.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as exc:
        raise EndNoteError(f"Not a readable DOCX package: {path}") from exc
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)
    return etree.fromstring(xml_bytes, parser=parser), xml_bytes


def body_paragraphs(root: etree._Element) -> list[etree._Element]:
    return root.xpath(".//w:body//w:p", namespaces=NS)


def has_tracked_changes(root: etree._Element) -> bool:
    tags = ("ins", "del", "moveFrom", "moveTo")
    return any(root.xpath(f".//w:{tag}", namespaces=NS) for tag in tags)


def split_citation_display(display: str) -> tuple[int, ...]:
    if not (display.startswith("[") and display.endswith("]")):
        raise EndNoteError(f"Malformed citation display: {display}")
    body = display[1:-1].strip().replace("\u2013", "-").replace("\u2014", "-")
    if not body:
        raise EndNoteError(f"Empty citation display: {display}")
    numbers: list[int] = []
    for token in re.split(r"[,;]", body):
        token = token.strip()
        if not token:
            raise EndNoteError(f"Malformed citation display: {display}")
        if "-" in token:
            parts = [part.strip() for part in token.split("-")]
            if len(parts) != 2 or not all(part.isdigit() for part in parts):
                raise EndNoteError(f"Malformed citation range: {display}")
            start, end = map(int, parts)
            if start <= 0 or end < start or end - start > 500:
                raise EndNoteError(f"Invalid citation range: {display}")
            numbers.extend(range(start, end + 1))
        elif token.isdigit() and int(token) > 0:
            numbers.append(int(token))
        else:
            raise EndNoteError(f"Malformed citation number: {display}")
    ordered = tuple(dict.fromkeys(numbers))
    if not ordered:
        raise EndNoteError(f"No citation numbers found: {display}")
    return ordered


def iter_pre_reference_paragraphs(root: etree._Element) -> Iterator[tuple[int, etree._Element]]:
    found_heading = False
    for index, paragraph in enumerate(body_paragraphs(root), start=1):
        text = paragraph_text(paragraph)
        if is_reference_heading(text):
            found_heading = True
            break
        yield index, paragraph
    if not found_heading:
        raise EndNoteError("Could not locate a References, Reference List, or Bibliography heading")


def collect_occurrences(root: etree._Element) -> list[CitationOccurrence]:
    occurrences: list[CitationOccurrence] = []
    for p_index, paragraph in iter_pre_reference_paragraphs(root):
        for match in CITATION_RE.finditer(paragraph_text(paragraph)):
            occurrences.append(
                CitationOccurrence(
                    paragraph_index=p_index,
                    display=match.group(0),
                    numbers=split_citation_display(match.group(0)),
                    in_table=paragraph_in_table(paragraph),
                )
            )
    return occurrences


def audit_document(path: Path) -> dict:
    root, _ = load_document_xml(path)
    visible = collect_occurrences(root)
    active = data = reflist = 0
    mixed: list[int] = []
    plain = 0
    for p_index, paragraph in iter_pre_reference_paragraphs(root):
        instructions = field_instructions(paragraph)
        p_active = sum(bool(INSTR_CITE_RE.match(item)) for item in instructions)
        p_data = sum(bool(INSTR_DATA_RE.match(item)) for item in instructions)
        p_visible = len(CITATION_RE.findall(paragraph_text(paragraph)))
        active += p_active
        data += p_data
        p_plain = p_visible - p_active
        if p_plain < 0:
            raise EndNoteError(f"Paragraph {p_index} has more EN.CITE fields than visible citation groups")
        plain += p_plain
        if p_active and p_plain:
            mixed.append(p_index)
    for item in root.xpath(".//w:instrText/text()", namespaces=NS):
        if INSTR_REFLIST_RE.match(normalize_instruction(item)):
            reflist += 1
    embedded = extract_embedded_records_from_root(root, source=str(path))
    return {
        "document": str(path.resolve()),
        "visible_citation_groups": len(visible),
        "active_endnote_cite_fields": active,
        "active_endnote_data_fields": data,
        "plain_citation_groups": plain,
        "mixed_paragraphs": mixed,
        "endnote_reflist_fields": reflist,
        "tracked_changes_present": has_tracked_changes(root),
        "embedded_record_candidates": len(embedded),
        "embedded_unique_dois": len({item.doi for item in embedded if item.doi}),
        "citation_locations": [
            {
                "paragraph": item.paragraph_index,
                "display": item.display,
                "numbers": list(item.numbers),
                "in_table": item.in_table,
            }
            for item in visible
        ],
    }


def extract_embedded_records_from_root(root: etree._Element, source: str) -> list[RecordCandidate]:
    candidates: list[RecordCandidate] = []
    seen: set[tuple[str | None, str | None, bytes]] = set()
    for encoded in root.xpath(".//w:fldData/text()", namespaces=NS):
        try:
            payload = base64.b64decode("".join(encoded.split()), validate=True).rstrip(b"\x00")
            endnote = etree.fromstring(payload, parser=etree.XMLParser(resolve_entities=False, huge_tree=True))
        except Exception:
            continue
        for record in endnote.xpath(".//record"):
            candidate = candidate_from_record(record, source)
            serialized = etree.tostring(candidate.record)
            key = (candidate.doi, candidate.title_key, serialized)
            if key not in seen:
                seen.add(key)
                candidates.append(candidate)
    return candidates


def extract_embedded_records(path: Path) -> list[RecordCandidate]:
    root, _ = load_document_xml(path)
    return extract_embedded_records_from_root(root, source=str(path.resolve()))


def candidate_from_record(record: etree._Element, source: str) -> RecordCandidate:
    record = copy.deepcopy(record)
    doi_nodes = record.xpath("./electronic-resource-num/text()")
    title_nodes = record.xpath("./titles/title/text()")
    return RecordCandidate(
        record=record,
        source=source,
        doi=normalize_doi(doi_nodes[0] if doi_nodes else None),
        title_key=normalize_title(title_nodes[0] if title_nodes else None),
    )


def _open_library_copy(path: Path) -> tuple[sqlite3.Connection, tempfile.TemporaryDirectory | None]:
    connection: sqlite3.Connection | None = None
    try:
        connection = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
        connection.execute("SELECT 1 FROM refs LIMIT 1").fetchone()
        return connection, None
    except sqlite3.Error:
        if connection is not None:
            connection.close()
        holder = tempfile.TemporaryDirectory(prefix="endnote-library-")
        copied = Path(holder.name) / path.name
        shutil.copy2(path, copied)
        connection = sqlite3.connect(copied)
        return connection, holder


def _append_text(parent: etree._Element, tag: str, value: object) -> etree._Element | None:
    text = str(value or "").strip()
    if not text:
        return None
    node = etree.SubElement(parent, tag)
    node.text = text
    return node


def _split_endnote_list(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in re.split(r"\r\n|\r|\n|\^M", value) if item.strip()]


def record_from_library_row(row: sqlite3.Row, db_id: str, source: str) -> RecordCandidate:
    ref_type = int(row["reference_type"])
    if ref_type != 0:
        raise EndNoteError(
            f"Unsupported .enl reference_type={ref_type} for record {row['id']} in {source}; "
            "supply a DOCX containing its traveling-library record"
        )
    record = etree.Element("record")
    _append_text(record, "rec-number", row["id"])
    foreign = etree.SubElement(record, "foreign-keys")
    key = etree.SubElement(
        foreign,
        "key",
        app="EN",
        **{
            "db-id": db_id,
            "timestamp": str(row["record_last_updated"] or row["added_to_library"] or 0),
        },
    )
    key.text = str(row["id"])
    type_node = etree.SubElement(record, "ref-type", name="Journal Article")
    type_node.text = "17"

    authors = _split_endnote_list(row["author"])
    if authors:
        contributors = etree.SubElement(record, "contributors")
        author_nodes = etree.SubElement(contributors, "authors")
        for author in authors:
            _append_text(author_nodes, "author", author)

    _append_text(record, "auth-address", row["author_address"])
    titles = etree.SubElement(record, "titles")
    _append_text(titles, "title", row["title"])
    _append_text(titles, "secondary-title", row["secondary_title"])
    if row["secondary_title"]:
        periodical = etree.SubElement(record, "periodical")
        _append_text(periodical, "full-title", row["secondary_title"])
    for key_name, xml_name in (
        ("pages", "pages"),
        ("volume", "volume"),
        ("number", "number"),
        ("edition", "edition"),
    ):
        _append_text(record, xml_name, row[key_name])

    keywords = _split_endnote_list(row["keywords"])
    if keywords:
        keyword_nodes = etree.SubElement(record, "keywords")
        for keyword in keywords:
            _append_text(keyword_nodes, "keyword", keyword)

    if row["year"] or row["date"]:
        dates = etree.SubElement(record, "dates")
        _append_text(dates, "year", row["year"])
        if row["date"]:
            pub_dates = etree.SubElement(dates, "pub-dates")
            _append_text(pub_dates, "date", row["date"])
    _append_text(record, "isbn", row["isbn"])
    _append_text(record, "accession-num", row["accession_number"])
    urls = _split_endnote_list(row["url"])
    if urls:
        urls_node = etree.SubElement(record, "urls")
        related = etree.SubElement(urls_node, "related-urls")
        for url in urls:
            _append_text(related, "url", url)
    _append_text(record, "electronic-resource-num", row["electronic_resource_number"])
    _append_text(record, "remote-database-name", row["name_of_database"])
    _append_text(record, "remote-database-provider", row["database_provider"])
    return candidate_from_record(record, source)


def records_from_library(path: Path) -> list[RecordCandidate]:
    if not path.is_file():
        raise EndNoteError(f"EndNote library not found: {path}")
    connection, holder = _open_library_copy(path)
    try:
        connection.row_factory = sqlite3.Row
        table_names = {
            row[0]
            for row in connection.execute("SELECT name FROM sqlite_master WHERE type='table'")
        }
        if not {"refs", "misc"}.issubset(table_names):
            raise EndNoteError(f"Unsupported EndNote library schema: {path}")
        db_row = connection.execute(
            "SELECT value FROM misc WHERE code=14 AND subcode=0"
        ).fetchone()
        if not db_row or not db_row[0]:
            raise EndNoteError(f"Could not determine EndNote library db-id: {path}")
        db_id = db_row[0].decode() if isinstance(db_row[0], bytes) else str(db_row[0])
        rows = connection.execute(
            "SELECT * FROM refs WHERE trash_state=0 ORDER BY id"
        ).fetchall()
        records: list[RecordCandidate] = []
        for row in rows:
            try:
                records.append(
                    record_from_library_row(row, db_id, str(path.resolve()))
                )
            except EndNoteError:
                # Unsupported types remain unresolved unless a DOCX supplies the
                # canonical traveling-library record. Do not coerce or fabricate.
                continue
        return records
    finally:
        connection.close()
        if holder is not None:
            holder.cleanup()


def parse_reference_map(path: Path | None) -> dict[int, str]:
    if path is None:
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise EndNoteError("Reference map must be a JSON object")
    result: dict[int, str] = {}
    for raw_number, raw_value in data.items():
        try:
            number = int(raw_number)
        except (TypeError, ValueError) as exc:
            raise EndNoteError(f"Invalid reference-map key: {raw_number!r}") from exc
        value = raw_value.get("doi") if isinstance(raw_value, dict) else raw_value
        doi = normalize_doi(str(value) if value is not None else None)
        if number <= 0 or not doi:
            raise EndNoteError(f"Reference-map entry {raw_number!r} must contain a DOI")
        result[number] = doi
    return result


def parse_bibliography(
    root: etree._Element, explicit_map: dict[int, str] | None = None
) -> dict[int, ReferenceEntry]:
    explicit_map = explicit_map or {}
    paragraphs = body_paragraphs(root)
    start = None
    for index, paragraph in enumerate(paragraphs):
        if is_reference_heading(paragraph_text(paragraph)):
            start = index + 1
            break
    if start is None:
        raise EndNoteError("Could not locate the bibliography heading")
    entries: dict[int, ReferenceEntry] = {}
    fallback_number = 1
    for paragraph in paragraphs[start:]:
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        style = paragraph_style(paragraph).casefold()
        match = re.match(r"^\s*(\d+)\s*[.\t)]\s*(.*)$", text, flags=re.S)
        if match:
            number = int(match.group(1))
        elif "endnote" in style and "bibliograph" in style:
            number = fallback_number
        else:
            continue
        fallback_number = max(fallback_number, number + 1)
        if number in entries:
            raise EndNoteError(f"Duplicate bibliography number: {number}")
        doi_match = DOI_RE.search(text)
        doi = normalize_doi(doi_match.group(1) if doi_match else explicit_map.get(number))
        entries[number] = ReferenceEntry(number=number, text=text, doi=doi)
    if not entries:
        raise EndNoteError("No numbered bibliography entries were found")
    return entries


def index_records(
    candidates: Iterable[RecordCandidate],
) -> dict[str, list[RecordCandidate]]:
    by_doi: dict[str, list[RecordCandidate]] = {}
    for candidate in candidates:
        if candidate.doi:
            by_doi.setdefault(candidate.doi, []).append(candidate)
    return by_doi


def resolve_records(
    references: dict[int, ReferenceEntry],
    cited_numbers: Iterable[int],
    candidates: Sequence[RecordCandidate],
) -> tuple[dict[int, RecordCandidate], list[dict]]:
    by_doi = index_records(candidates)
    resolved: dict[int, RecordCandidate] = {}
    issues: list[dict] = []
    for number in sorted(set(cited_numbers)):
        reference = references.get(number)
        if reference is None:
            issues.append(
                {"number": number, "issue": "citation number absent from bibliography"}
            )
            continue
        if not reference.doi:
            issues.append(
                {
                    "number": number,
                    "issue": "bibliography entry has no DOI or explicit DOI map",
                    "reference": reference.text,
                }
            )
            continue
        matches = by_doi.get(reference.doi, [])
        if not matches:
            issues.append(
                {
                    "number": number,
                    "doi": reference.doi,
                    "issue": "no EndNote record matched",
                }
            )
            continue
        first_source = matches[0].source
        preferred = [item for item in matches if item.source == first_source]
        unique_payloads = {etree.tostring(item.record) for item in preferred}
        if len(unique_payloads) != 1:
            issues.append(
                {
                    "number": number,
                    "doi": reference.doi,
                    "issue": "multiple non-identical records in preferred source",
                    "source": first_source,
                }
            )
            continue
        resolved[number] = preferred[0]
    return resolved, issues


def _first_text(record: etree._Element, xpath: str, fallback: str = "") -> str:
    values = record.xpath(xpath)
    return str(values[0]) if values else fallback


def build_endnote_payload(
    records: Sequence[RecordCandidate], display: str
) -> str:
    endnote = etree.Element("EndNote")
    for index, candidate in enumerate(records):
        record = copy.deepcopy(candidate.record)
        cite = etree.SubElement(endnote, "Cite")
        authors = record.xpath("./contributors/authors/author/text()")
        first_author = authors[0] if authors else ""
        surname = first_author.split(",", 1)[0].strip()
        if "," not in first_author and first_author:
            surname = first_author.split()[-1]
        _append_text(cite, "Author", surname)
        _append_text(cite, "Year", _first_text(record, "./dates/year/text()"))
        _append_text(cite, "RecNum", _first_text(record, "./rec-number/text()"))
        if index == 0:
            display_node = etree.SubElement(cite, "DisplayText")
            style = etree.SubElement(display_node, "style", size="10")
            style.text = display
        cite.append(record)
    payload = (
        etree.tostring(endnote, encoding="utf-8", xml_declaration=False) + b"\x00"
    )
    return base64.encodebytes(payload).decode("ascii").rstrip("\n")


def _new_run(
    child: etree._Element, run_properties: etree._Element | None = None
) -> etree._Element:
    run = etree.Element(W + "r")
    if run_properties is not None:
        run.append(copy.deepcopy(run_properties))
    run.append(child)
    return run


def _text_run(
    text: str, run_properties: etree._Element | None
) -> etree._Element:
    node = etree.Element(W + "t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        node.set(XML_SPACE, "preserve")
    node.text = text
    return _new_run(node, run_properties)


def _field_runs(
    display: str,
    records: Sequence[RecordCandidate],
    run_properties: etree._Element | None,
) -> list[etree._Element]:
    encoded = build_endnote_payload(records, display)

    def fld_char(kind: str, with_data: bool = False) -> etree._Element:
        node = etree.Element(W + "fldChar")
        node.set(W + "fldCharType", kind)
        if with_data:
            data = etree.SubElement(node, W + "fldData")
            data.set(XML_SPACE, "preserve")
            data.text = encoded
        return node

    def instruction(value: str) -> etree._Element:
        node = etree.Element(W + "instrText")
        node.set(XML_SPACE, "preserve")
        node.text = value
        return node

    return [
        _new_run(fld_char("begin", with_data=True)),
        _new_run(instruction(" ADDIN EN.CITE ")),
        _new_run(fld_char("begin", with_data=True)),
        _new_run(instruction(" ADDIN EN.CITE.DATA ")),
        _new_run(fld_char("end")),
        _new_run(fld_char("separate")),
        _text_run(display, run_properties),
        _new_run(fld_char("end")),
    ]


def _run_field_kind(node: etree._Element) -> str | None:
    if node.tag != W + "r":
        return None
    values = node.xpath("./w:fldChar/@w:fldCharType", namespaces=NS)
    return values[0] if len(values) == 1 else None


def flatten_endnote_citation_fields(root: etree._Element) -> tuple[int, list[dict]]:
    """Replace each complete EN.CITE complex field with its visible result text."""
    flattened = 0
    issues: list[dict] = []
    for p_index, paragraph in iter_pre_reference_paragraphs(root):
        children = list(paragraph)
        index = 0
        while index < len(children):
            if _run_field_kind(children[index]) != "begin":
                index += 1
                continue
            lookahead = children[index + 1 : index + 4]
            instructions = [
                normalize_instruction("".join(node.xpath(".//w:instrText/text()", namespaces=NS)))
                for node in lookahead
            ]
            if not any(INSTR_CITE_RE.match(item) for item in instructions):
                index += 1
                continue

            depth = 0
            separated = False
            end_index: int | None = None
            display_parts: list[str] = []
            display_rpr: etree._Element | None = None
            for cursor in range(index, len(children)):
                child = children[cursor]
                kind = _run_field_kind(child)
                if kind == "begin":
                    depth += 1
                elif kind == "separate" and depth == 1:
                    separated = True
                elif kind == "end":
                    depth -= 1
                    if depth == 0:
                        end_index = cursor
                        break
                elif separated and depth == 1:
                    text = "".join(child.xpath(".//w:t/text()", namespaces=NS))
                    if text:
                        display_parts.append(text)
                        if display_rpr is None and child.tag == W + "r":
                            display_rpr = child.find(W + "rPr")
            display = "".join(display_parts)
            if end_index is None or not separated or not CITATION_RE.fullmatch(display):
                issues.append(
                    {
                        "paragraph": p_index,
                        "issue": "existing EN.CITE field is incomplete or has a non-numeric display",
                    }
                )
                index += 1
                continue

            for child in children[index : end_index + 1]:
                paragraph.remove(child)
            replacement = _text_run(display, display_rpr)
            paragraph.insert(index, replacement)
            children[index : end_index + 1] = [replacement]
            flattened += 1
            index += 1
    return flattened, issues


def rebuild_plain_citations(
    root: etree._Element, resolved: dict[int, RecordCandidate]
) -> tuple[int, list[dict]]:
    conversions = 0
    issues: list[dict] = []
    for p_index, paragraph in iter_pre_reference_paragraphs(root):
        instructions = field_instructions(paragraph)
        active = sum(bool(INSTR_CITE_RE.match(item)) for item in instructions)
        visible = len(CITATION_RE.findall(paragraph_text(paragraph)))
        if active and visible > active:
            issues.append(
                {
                    "paragraph": p_index,
                    "issue": "paragraph mixes active and plain citation groups",
                }
            )
            continue
        if active:
            continue
        targets = []
        for text_node in paragraph.xpath(".//w:t", namespaces=NS):
            value = text_node.text or ""
            matches = list(CITATION_RE.finditer(value))
            if not matches:
                continue
            run = text_node.getparent()
            if run is None or run.tag != W + "r":
                issues.append(
                    {
                        "paragraph": p_index,
                        "issue": "citation text is not in a normal Word run",
                    }
                )
                continue
            if run.xpath("ancestor::w:hyperlink", namespaces=NS):
                issues.append(
                    {
                        "paragraph": p_index,
                        "issue": "citation occurs inside a hyperlink",
                    }
                )
                continue
            text_children = run.xpath("./w:t", namespaces=NS)
            unsafe_children = [
                child for child in run if child.tag not in {W + "rPr", W + "t"}
            ]
            if len(text_children) != 1 or unsafe_children:
                issues.append(
                    {
                        "paragraph": p_index,
                        "issue": "citation run contains non-text content and cannot be split safely",
                    }
                )
                continue
            targets.append((run, value, matches))

        for run, value, matches in targets:
            parent = run.getparent()
            if parent is None:
                issues.append(
                    {"paragraph": p_index, "issue": "citation run has no parent"}
                )
                continue
            rpr = run.find(W + "rPr")
            replacement: list[etree._Element] = []
            cursor = 0
            local_conversions = 0
            for match in matches:
                if match.start() > cursor:
                    replacement.append(_text_run(value[cursor : match.start()], rpr))
                display = match.group(0)
                numbers = split_citation_display(display)
                try:
                    records = [resolved[number] for number in numbers]
                except KeyError as exc:
                    issues.append(
                        {
                            "paragraph": p_index,
                            "display": display,
                            "issue": f"reference {exc.args[0]} is unresolved",
                        }
                    )
                    replacement = []
                    break
                replacement.extend(_field_runs(display, records, rpr))
                local_conversions += 1
                cursor = match.end()
            if not replacement:
                continue
            if cursor < len(value):
                replacement.append(_text_run(value[cursor:], rpr))
            index = parent.index(run)
            parent.remove(run)
            for offset, item in enumerate(replacement):
                parent.insert(index + offset, item)
            conversions += local_conversions
    return conversions, issues


def write_docx_with_document_xml(
    input_path: Path, output_path: Path, root: etree._Element
) -> None:
    if input_path.resolve() == output_path.resolve():
        raise EndNoteError("Refusing to overwrite the input DOCX")
    if output_path.exists():
        raise EndNoteError(f"Output already exists: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    xml_bytes = etree.tostring(
        root, encoding="UTF-8", xml_declaration=True, standalone=True
    )
    temporary = output_path.with_name(output_path.name + ".partial")
    try:
        with zipfile.ZipFile(input_path, "r") as source, zipfile.ZipFile(
            temporary, "w"
        ) as target:
            for info in source.infolist():
                payload = (
                    xml_bytes
                    if info.filename == "word/document.xml"
                    else source.read(info.filename)
                )
                target.writestr(info, payload)
        with zipfile.ZipFile(temporary) as check:
            bad_member = check.testzip()
            if bad_member:
                raise EndNoteError(f"Generated DOCX failed ZIP validation at {bad_member}")
            etree.fromstring(check.read("word/document.xml"))
        temporary.replace(output_path)
    finally:
        if temporary.exists():
            temporary.unlink()
