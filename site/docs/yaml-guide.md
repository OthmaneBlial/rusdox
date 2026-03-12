# YAML Guide

RusDox is designed so the YAML reads like the document.

The basic shape is always:

```yaml
output_name: my-document
blocks:
  - type: title
    text: My Document
  - type: body
    text: This is a paragraph.
```

## Top-Level Fields

`output_name`

- Optional
- Controls the file name of the generated output
- If omitted, RusDox uses the spec file stem

`blocks`

- Required for real documents
- Ordered from top to bottom
- This is the document structure

## Simple Block Types

Use these when you want readable documents with strong defaults:

- `cover_title`
- `title`
- `subtitle`
- `hero`
- `centered_note`
- `page_heading`
- `section`
- `body`
- `tagline`
- `spacer`

Example:

```yaml
blocks:
  - type: cover_title
    text: Board Report
  - type: subtitle
    text: March 2026
  - type: hero
    text: Prepared automatically with RusDox
  - type: centered_note
    text: Internal use only
  - type: page_heading
    text: Board Narrative
  - type: body
    text: Revenue expanded faster than forecast.
```

`page_heading` starts a new page before the heading.

`spacer` adds vertical space when you want breathing room between sections.

## List Blocks

### `bullets`

```yaml
- type: bullets
  items:
    - Confirm launch date
    - Finalize sales enablement
    - Publish support macros
```

### `label_values`

Good for notes, metadata, meeting headers, and document summaries.

```yaml
- type: label_values
  items:
    - label: Date
      value: 2026-03-10
    - label: Owner
      value: Operations
```

### `metrics`

Good for dashboard-style cards.

Available tones:

- `positive`
- `neutral`
- `warning`
- `risk`

```yaml
- type: metrics
  items:
    - label: ARR
      value: $18.7M
      tone: positive
    - label: Cash Runway
      value: 24 mo
      tone: warning
```

## Tables

Tables use explicit columns and rows.

Each column needs:

- `label`
- `width`

Each row has `cells`.

Each cell can be:

- `kind: text`
- `kind: status`

Example:

```yaml
- type: table
  spec:
    columns:
      - label: Category
        width: 2800
      - label: Current
        width: 2000
      - label: Status
        width: 1600
    rows:
      - cells:
          - kind: text
            text: ARR
          - kind: text
            text: $18.7M
          - kind: status
            text: Strong
            tone: positive
```

`width` values are DOCX table widths in twips.

You usually only need to copy a working table from an example and edit the content.

## Custom Paragraphs

Use `paragraph` when you need mixed formatting inside one paragraph.

Example:

```yaml
- type: paragraph
  spec:
    alignment: center
    spacing_after_twips: 120
    runs:
      - text: "This is "
      - text: important
        bold: true
      - text: " and "
      - text: styled
        italic: true
        color: D00000
```

Supported paragraph fields:

- `runs`
- `alignment`: `left`, `center`, `right`, `justified`
- `spacing_before_twips`
- `spacing_after_twips`
- `page_break_before`

Supported run fields:

- `text`
- `bold`
- `italic`
- `underline`
- `strikethrough`
- `small_caps`
- `shadow`
- `color`
- `font_family`
- `size_pt`
- `vertical_align`

Underline values:

- `single`
- `double`
- `dotted`
- `dash`
- `wavy`
- `words`
- `none`

Vertical align values:

- `superscript`
- `subscript`
- `baseline`

## Best Practices

- Keep content in variables only if you are in Rust. In YAML, keep it in document order.
- Prefer `title`, `section`, `body`, `bullets`, `metrics`, and `table` before reaching for custom paragraphs.
- Let config control styling instead of repeating style values everywhere.
- Copy a close example and edit the content.

## See Real Files

- [../examples/board_report.yaml](../examples/board_report.yaml)
- [../examples/executive_dashboard.yaml](../examples/executive_dashboard.yaml)
- [../examples/formatting_showcase.yaml](../examples/formatting_showcase.yaml)
