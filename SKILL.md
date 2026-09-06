---
name: endnote-cwyw-fields
description: Audit, rebuild, and verify editable EndNote Cite While You Write (CWYW) citation fields in Microsoft Word DOCX files. Use when numeric citations look correct but may be plain text, when revisions flattened EndNote fields, or when a manuscript must be restored to EndNote-managed citations without overwriting the source document.
---

# EndNote CWYW Fields

Restore genuine Word fields containing `ADDIN EN.CITE`, `ADDIN EN.CITE.DATA`, embedded traveling-library records, and an EndNote-managed bibliography. Do not treat an `EndNoteBibliography` paragraph style or visible text such as `[12]` as proof that a citation is editable by EndNote.

## Workflow

1. Locate the manuscript, the best pre-edit DOCX that still contains EndNote fields, and any `.enl` libraries. Keep each `.enl` beside its matching `.Data` directory.
2. Audit every candidate document:

   ```bash
   python scripts/audit_endnote_fields.py manuscript.docx --json audit.json
   ```

   Report active `EN.CITE` fields, paired `EN.CITE.DATA` fields, plain numeric citation groups, `EN.REFLIST` fields, tracked changes, and embedded traveling-library records. Active citation fields and visible citation groups must be counted separately.
3. Select record sources in this order:
   - the target DOCX and pre-edit DOCX files containing embedded traveling-library records;
   - the EndNote library intended to manage the manuscript;
   - another library only to identify records missing from the intended library.
4. Run a dry plan before writing:

   ```bash
   python scripts/rebuild_endnote_fields.py manuscript.docx \
     --source-docx pre_edit.docx \
     --library manuscript.enl
   ```

   The command must stop on missing or ambiguous records, citation numbers absent from the bibliography, malformed ranges, unsupported EndNote reference types, or tracked changes. A write rebuilds both existing and plain citation groups from the resolved records so stale active fields cannot survive unnoticed.
5. If every citation resolves uniquely, write to a new file:

   ```bash
   python scripts/rebuild_endnote_fields.py manuscript.docx \
     --source-docx pre_edit.docx \
     --library manuscript.enl \
     --output manuscript_endnote.docx
   ```

6. Audit the output. Require:
   - zero plain citation groups before the References section;
   - one `EN.CITE` and one `EN.CITE.DATA` field per visible citation group;
   - no out-of-range citation numbers;
   - one `EN.REFLIST` field when the source bibliography was EndNote-managed;
   - unchanged visible manuscript text outside field markup.
7. Open the new file in Word with the EndNote CWYW add-in. Use **Update Citations and Bibliography** once, resolve any matching-reference dialog by DOI/title, save, close, reopen, and run the audit again. The Word/EndNote round-trip is required before calling the result fully validated.

## Safety boundaries

- Never overwrite the input DOCX. Preserve the tracked manuscript separately; rebuild fields in an accepted/clean copy because tracked revisions can split or delete field boundaries.
- Never insert field-code-looking text without an embedded EndNote `<record>`. Such text is not a portable CWYW citation.
- Never renumber citations independently of the bibliography. Visible numbers are only lookup keys for the current document.
- Prefer DOI matching. For sources without a DOI, provide an explicit JSON map from reference number to DOI rather than guessing from a partial title.
- Treat the SQLite layout of `.enl` as a read-only adapter. The scripts never modify an EndNote library.
- Do not claim EndNote compatibility from OOXML inspection alone. Distinguish structural validation from the final Word/EndNote application round-trip.

For the field structure, failure conditions, and manual recovery procedure, read [references/ooxml-safety.md](references/ooxml-safety.md).

## Scripts

The scripts require Python 3.10+ and `lxml`.

- `scripts/audit_endnote_fields.py`: read-only structural audit of DOCX citations and bibliography fields.
- `scripts/rebuild_endnote_fields.py`: dry-run planner and fail-closed full rebuild of numeric citations as CWYW fields.
- `scripts/endnote_ooxml.py`: shared DOCX, EndNote-record, matching, and OOXML implementation.
