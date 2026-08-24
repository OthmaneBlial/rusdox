# International and Accessible Output

RusDox preserves Unicode text, meaningful image descriptions, and document
language metadata across its supported authoring paths. Those guarantees are
machine-tested. They are not a claim that every script is fully laid out or
that the current PDF is an accessible tagged document.

## Support decision

Latin text is supported. Arabic/right-to-left, Hebrew, CJK, and emoji remain
experimental in the native PDF renderer. The checked-in international fixture
passes semantic DOCX reopen, PDF text projection, language metadata, and the
deterministic Linux page baseline, but RusDox does not yet implement:

- the Unicode bidirectional algorithm and contextual Arabic shaping;
- CJK kinsoku line-breaking rules or a pinned representative CJK font profile;
- color or ZWJ emoji composition;
- a dated review in representative Word, LibreOffice, Preview, and Acrobat
  versions for these script combinations.

[`compatibility/international-readiness.json`](../compatibility/international-readiness.json)
is the executable graduation gate. A track cannot change from `experimental`
to `supported` while any required gate is false. This intentionally prevents a
green semantic round-trip from being presented as typographic parity.

## Font embedding and fallback contract

DOCX stores the requested family names and lets the reader choose an installed
font when those families are absent. RusDox does not embed fonts in DOCX, so
Word or LibreOffice substitution can change line breaks and pagination.

The native PDF path resolves the requested family, common sans/serif/mono
fallbacks, and then script-specific Arabic, Hebrew, CJK, or Devanagari
fallbacks. It admits only single-face outline TrueType fonts for which OpenType
reports installable, preview-and-print, or editable embedding permission.
Restricted fonts and fonts that forbid outline embedding are rejected. The
admitted font program is embedded in full and accompanied by a Unicode map.

Parity evidence records, for every font used:

- the resolved family and PostScript name;
- the OpenType embedding permission;
- the number of mapped Unicode characters;
- every character that fell back to the font's missing-glyph box.

This is a technical permission check, not a font licence grant. Distributors
must confirm that their selected font licence permits the intended embedding
and redistribution. Production deployments that need reproducible pagination
must pin and licence an explicit font set in the container or host image;
system font discovery is intentionally outside cross-machine byte parity.

## Alternative text and language

Every visual block must provide non-blank `alt_text`. RusDox writes the same
description to both relevant DOCX drawing-property locations, reopens it, and
compares it with the source and PDF semantic projection in the
`image_alt_text` parity check. Decorative-image semantics are not yet exposed;
authors should describe the visual's purpose, not its file name.

`metadata.language` accepts a bounded BCP 47-style language tag such as
`en-US`, `fr`, or `mul`. It is written as `dc:language` in DOCX core properties
and `/Lang` in the PDF catalog. The `document_language` parity check compares
the source, reopened DOCX, PDF projection, and rendered PDF evidence. Documents
that intentionally mix languages can use `mul`; future run-level language
changes will require their own contract.

The current PDF records image descriptions in its semantic projection, but it
does not attach them to a PDF structure tree. Assistive technology therefore
cannot consume those descriptions as tagged figure alternatives. See the
[PDF conformance research](pdf-conformance-research.md) before making any PDF
accessibility or archival claim.

## Accessibility parity boundary

The 21-check parity report adds two explicit accessibility-oriented assertions:

- `image_alt_text` passes when every visual has identical meaningful text in
  the source, reopened DOCX, and PDF projection; it skips when there are no
  visuals;
- `document_language` passes when all four representations agree; it skips when
  the author did not declare a language.

These checks prevent metadata loss. They do not replace keyboard, screen-reader,
reading-order, contrast, or representative-viewer review. Tagged PDF and PDF/A
remain separate, unclaimed conformance projects.

## Graduation checklist

RTL or CJK can become supported only after the implementation adds the missing
layout algorithm, representative licensed font profiles, deterministic parity
fixtures, and dated viewer evidence, then every machine-readable graduation
gate passes. Tagged PDF requires the separate structure, validation, and human
assistive-technology work described in the research note.
