# OOXML structure and safety notes

## What constitutes a genuine EndNote citation

A portable CWYW citation is an outer Word complex field whose instruction is `ADDIN EN.CITE`. Its field-begin element contains base64-encoded, NUL-terminated XML with one or more `<Cite>` elements and complete `<record>` metadata. EndNote also writes a nested `ADDIN EN.CITE.DATA` field carrying the same data. The visible bracketed citation appears after the outer field separator.

An EndNote bibliography is normally a complex field with instruction `ADDIN EN.REFLIST`. Bibliography paragraph styles alone do not establish field linkage.

## Record matching

Numbered citation displays are document-local. Resolve each cited number to the corresponding bibliography entry, then match that entry to an embedded traveling record or `.enl` record by normalized DOI. Do not assume that bibliography number 27 equals EndNote record number 27.

When the same DOI occurs in multiple record sources, prefer the first source supplied. The rebuild command reports duplicate candidates and uses a record only when its normalized DOI resolves deterministically. Add the desired managing library before supplemental libraries.

The `.enl` reader supports EndNote's common Journal Article record type. For other reference types, supply a DOCX containing the original traveling-library record rather than coercing the database row to a generic type.

## Failure conditions

Stop without writing when any of the following applies:

- the document contains tracked insertions, deletions, or moves;
- an existing EndNote field is incomplete or does not have a numeric citation display;
- a bracketed group cannot be parsed as a positive integer list/range;
- a cited number has no bibliography entry;
- a bibliography entry lacks both a DOI and an explicit mapping;
- no unique record matches the DOI;
- a citation occurs inside a hyperlink or a non-text run that cannot be split safely;
- the requested output is the input path or already exists.

These are integrity failures, not invitations to guess or fabricate a record.

## Manual recovery

If records are missing, import them into the intended EndNote library using DOI, PubMed, RIS, or EndNote XML, verify metadata inside EndNote, and rerun the dry plan. Do not modify the `.enl` SQLite database directly.

After structural rebuilding, Word and the EndNote add-in remain authoritative for application-level validation:

1. Work on a new DOCX copy.
2. Open the intended EndNote library, then the rebuilt DOCX.
3. In Word's EndNote tab, run **Update Citations and Bibliography**.
4. Resolve matching-reference prompts by DOI and title; reject incorrect duplicates.
5. Save, close, reopen, update once more, and audit the saved DOCX.

If Word or EndNote removes citations, changes visible numbering unexpectedly, or cannot read the traveling records, discard the generated copy and return to the untouched input.

## Validation levels

- **Audit verified:** field instructions and visible groups were counted correctly.
- **Structurally rebuilt:** every target group has paired `EN.CITE`/`EN.CITE.DATA` markup with embedded record XML and the visible text is preserved.
- **EndNote validated:** the rebuilt file survived an Update/Save/Close/Reopen cycle in Word with the EndNote add-in and passed a second audit.

Use the final label only after the application round-trip.
