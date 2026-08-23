# Word-native Templates

RusDox template syntax v1 turns a designer-authored DOCX and a JSON object into an edited DOCX, a native PDF, deterministic page snapshots, and an HTML/JSON parity report in one command.

    rusdox template verify proposal.docx data.json --strict

The renderer edits textual OOXML parts directly. It does not rebuild the package from a blank document, so untouched styles, sections, headers, footers, media, relationships, custom XML, and other package parts keep their exact bytes.

## Inspect Before Rendering

    rusdox template inspect proposal.docx
    rusdox template inspect proposal.docx --format json

Inspection lists each expression with the exact OOXML part and a human-readable paragraph or table-row location. Unclosed or mismatched blocks fail before output is written.

## Syntax v1

Scalar values use double braces:

    Prepared for {{ customer.name }}
    Invoice {{ invoice.number }}

Paths traverse nested JSON objects. items.0.name can address an array position. Values are XML-escaped; there is deliberately no raw-XML escape hatch.

Supported deterministic filters are:

    {{ customer.name | title }}
    {{ report.period | upper }}
    {{ code | lower }}
    {{ description | trim }}
    {{ optional | default("Not provided") }}

Filters run from left to right. An unknown filter is an error instead of silently changing meaning.

### Loops

Loop markers must occupy complete Word paragraphs or complete table rows:

    {{#each items}}
    {{ @index }}. {{ description }} — {{ quantity }}
    {{/each}}

this refers to the current item, nested fields resolve against the current item first, and @index is one-based. The start and end markers must use the same container type. This bounded rule makes repetition predictable and keeps the resulting OOXML structurally valid.

### Conditions

Conditions also use complete paragraphs or rows:

    {{#if include_support}}
    Managed support is included.
    {{else}}
    Managed support is optional.
    {{/if}}

False, null, zero, empty strings, empty arrays, and empty objects are false. Conditions are deliberately truthiness-only in syntax v1; arbitrary code and general-purpose expressions are not supported.

### Reusable partials

JSON can define reusable inline subtemplates under $partials:

    {
      "client": {"name": "Northstar Operations"},
      "$partials": {
        "confidentiality": "Confidential. Prepared for {{ client.name | title }}."
      }
    }

Insert one with:

    {{> confidentiality}}

Partials can reuse scalar paths and filters. Recursive partials and block markers inside a partial are errors.

## Missing Values and Strict Mode

Without --strict, a missing or null scalar becomes an empty string and the result includes a warning. This is useful for optional drafts while remaining visible in machine-readable output.

With --strict, every missing reference is an error, the previous DOCX destination remains untouched, and no PDF or parity report is produced:

    rusdox template render invoice.docx invoice.json --strict

Diagnostics include:

- part, such as word/document.xml or word/header1.xml;
- location, such as paragraph 8 or table row 3;
- the exact placeholder;
- the failure and a suggested fix.

## Output Contract

Both template render and template verify produce:

    generated/<name>.docx
    rendered/<name>.pdf
    reports/<name>-pages/page-001.png
    reports/<name>-parity.json
    reports/<name>-parity.html

template verify is the explicit CI spelling. A completed parity failure exits with code 2; an invalid template, missing strict value, unreadable input, or resource-limit failure exits with code 1.

Use a custom name and output root:

    rusdox template verify template.docx data.json \
      --name client-proposal \
      --output-root build \
      --strict \
      --format json

## Bundled Designer Templates

The checked-in sources were styled as RTF documents and exported to DOCX by the macOS text system, giving the workflow package evidence authored outside RusDox.

| Template | Features exercised | Source |
|---|---|---|
| Invoice | nested values, repeated line items, condition, partial, title filter | [templates/invoice/](../templates/invoice/) |
| Proposal | string-array loop, condition/else, nested partial, title filter | [templates/proposal/](../templates/proposal/) |
| Board report | landscape section, object-array loop, nested condition, upper filter, partial | [templates/board-report/](../templates/board-report/) |

Regenerate the external DOCX packages on macOS with ./scripts/generate_word_templates.sh. Render all three samples with ./scripts/verify_word_templates.sh.

## Preservation and Fidelity Boundary

RusDox preserves untouched package parts byte-for-byte and preserves surrounding run nodes when a placeholder is split across multiple Word runs. Automated tests cover sections, headers, footers, styles, media, relationships, complete table-row loops, strict recovery, and an externally authored DOCX.

The native PDF is rendered from the subset of Word body semantics that RusDox can parse. Unsupported arbitrary Word layout features can remain intact in the edited DOCX while not appearing identically in the native PDF. The parity report describes the semantics actually compared; it does not claim universal Word fidelity. Test representative templates in the viewer versions used by recipients.
