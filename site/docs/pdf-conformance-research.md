# Tagged PDF, PDF/UA, and PDF/A Research

Status: researched on 2026-08-24. Current RusDox output is **not tagged PDF**,
**not PDF/UA**, and **not PDF/A**. `PdfRenderEvidence` records these claims as
false. A language entry, embedded fonts, Unicode maps, and source-level image
descriptions are useful prerequisites, but they are not conformance.

## Tagged PDF and PDF/UA

The current accessibility standard is
[ISO 14289-2:2024 (PDF/UA-2)](https://www.iso.org/standard/82278.html). The
[PDF Association sponsored-standards programme](https://pdfa.org/sponsored-standards/)
provides free access to the relevant ISO standards, and its
[Tagged PDF Best Practice Guide](https://pdfa.org/resource/tagged-pdf-best-practice-guide-syntax/)
explains the syntax implementation boundary.

An honest first tagged-PDF milestone needs all of the following before any
accessibility claim:

1. `/MarkInfo` with marked content and a `/StructTreeRoot` rooted in a semantic
   document element;
2. marked-content IDs, structure elements, and a correct parent tree that bind
   every meaningful page object to the logical structure;
3. semantic roles for headings, paragraphs, lists, tables, links, notes, and
   figures, including correct table headers and nesting;
4. logical reading order independent of paint order, with decorative content
   marked as artifacts;
5. alternative text for figures, replacement or actual text where needed, a
   catalog language, and language changes within mixed-language content;
6. embedded fonts with reliable Unicode mapping, tagged annotations, document
   title, viewer preferences, and the required XMP identification metadata.

The W3C PDF techniques provide focused implementation checks for
[reading order](https://www.w3.org/WAI/WCAG21/Techniques/pdf/PDF3),
[image alternatives](https://www.w3.org/WAI/WCAG21/Techniques/pdf/PDF1),
[document language](https://www.w3.org/WAI/WCAG21/Techniques/pdf/PDF16),
[language changes](https://www.w3.org/WAI/WCAG21/Techniques/pdf/PDF19), and
[artifacts](https://www.w3.org/WAI/WCAG21/Techniques/pdf/PDF4). The
[Matterhorn Protocol](https://pdfa.org/resource/the-matterhorn-protocol/) is a
useful test catalogue, not a substitute for the standard.

CI should run the exact PDF/UA flavour of
[veraPDF validation](https://docs.verapdf.org/cli/validation/) and retain the
machine-readable report. A veraPDF pass is necessary but not sufficient:
release evidence must also include manual reading-order inspection and tests in
representative assistive technology. RusDox will not infer conformance from a
generic PDF validity check.

## PDF/A archival output

The [PDF/A overview](https://pdfa.org/resource/iso-19005-pdfa) separates
archival profiles by PDF generation and semantic level. The recommended first
RusDox target is PDF/A-2u: it fits the current PDF generation more closely and
requires reliable Unicode mapping without prematurely claiming a complete
logical structure. PDF/A-2a plus PDF/UA-1 can follow after tagged structure is
implemented. PDF/A-4 remains a later evaluation because the
[ISO 19005-4:2020 page](https://www.iso.org/standard/71832.html) currently marks
that edition as under revision.

PDF/A-2u implementation must, at minimum:

- embed every font and provide Unicode mapping for all text;
- embed a valid ICC output intent and make colour use self-contained;
- add correct PDF/A XMP identification and any required extension schemas;
- prohibit encryption, unresolved external rendering dependencies, and actions
  or JavaScript forbidden by the selected profile;
- define policy for attachments and annotations under the exact chosen part and
  conformance level;
- pass the explicit PDF/A-2u flavour in veraPDF, not merely auto-detection.

The PDF Association publishes guidance on
[XMP identification](https://pdfa.org/future-proofing-xmp-identification-schema/)
and [extension-schema templates](https://pdfa.org/resource/xmp-extension-schema-templates/).
The open [veraPDF validation profiles](https://github.com/veraPDF/veraPDF-validation-profiles)
make the automated rule set reviewable.

## Claim gate

RusDox may advertise tagged PDF, PDF/UA, or PDF/A only after the renderer emits
the complete required structures and metadata, fixtures exercise every relevant
feature, the exact veraPDF flavour passes in CI, and the human checks required
for accessibility are recorded. Until then, API evidence remains `false` and
the public compatibility matrix says unsupported rather than “nearly
compliant.”
