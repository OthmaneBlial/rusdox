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

`styles`

- Optional
- Defines reusable named paragraph, run, and table styles
- Supports inheritance through `based_on`
- Paragraph styles can also set `next`

## Named Styles

Use named styles when the document needs stable reusable formatting instead of repeating direct properties on every block.

Example:

```yaml
styles:
  paragraph:
    - id: lead
      based_on: Normal
      next: body
      paragraph:
        alignment: center
        spacing_after: 180
        keep_next: true
      run:
        bold: true
        color: "0F172A"
        font_family: Georgia
        font_size: 28
  run:
    - id: accent
      based_on: DefaultParagraphFont
      properties:
        italic: true
        color: "AA5500"
  table:
    - id: grid
      based_on: TableNormal
      properties:
        width: 9360

blocks:
  - type: paragraph
    spec:
      style_id: lead
      runs:
        - text: Styled
        - text: " emphasis"
          style_id: accent
  - type: table
    spec:
      style_id: grid
      columns:
        - label: Metric
          width: 4680
      rows:
        - cells:
            - kind: text
              text: ARR
```

Style definition groups:

- `styles.paragraph`
- `styles.run`
- `styles.table`

Paragraph style fields:

- `id`
- `name`
- `based_on`
- `next`
- `paragraph.list`
- `paragraph.alignment`
- `paragraph.spacing_before`
- `paragraph.spacing_after`
- `paragraph.keep_next`
- `paragraph.page_break_before`
- `run.bold`
- `run.italic`
- `run.underline`
- `run.strikethrough`
- `run.small_caps`
- `run.shadow`
- `run.color`
- `run.font_family`
- `run.font_size`
- `run.vertical_align`

Run style fields:

- `id`
- `name`
- `based_on`
- `properties.bold`
- `properties.italic`
- `properties.underline`
- `properties.strikethrough`
- `properties.small_caps`
- `properties.shadow`
- `properties.color`
- `properties.font_family`
- `properties.font_size`
- `properties.vertical_align`

Table style fields:

- `id`
- `name`
- `based_on`
- `properties.width`
- `properties.borders`

Units:

- paragraph spacing and table widths use twips
- run `font_size` uses OOXML half-points, so `24` means `12pt`

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
- `image`
- `logo`
- `signature`
- `chart`
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

## Visual Blocks

Use visual blocks when the document needs brand marks, screenshots, signatures, or SVG charts.

Example:

```yaml
- type: logo
  path: ../assets/rusdox-mark.svg
  alt_text: RusDox logo
  max_width_twips: 2200

- type: image
  path: ../assets/template-gallery.png
  alt_text: RusDox template gallery
  max_width_twips: 7200

- type: chart
  path: ../assets/benchmark-stress-1000-pages.svg
  alt_text: RusDox benchmark chart
  max_width_twips: 7200

- type: signature
  path: ../assets/signature-demo.svg
  alt_text: Automated approval signature
  max_width_twips: 2800
```

Supported visual fields:

- `path`
- `alt_text`
- `alignment`: `left`, `center`, `right`, `justified`
- `width_twips`
- `height_twips`
- `max_width_twips`
- `max_height_twips`

Supported file formats:

- PNG
- JPEG
- SVG

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
    style_id: grid
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
- `style_id`
- `alignment`: `left`, `center`, `right`, `justified`
- `spacing_before_twips`
- `spacing_after_twips`
- `page_break_before`

Supported run fields:

- `text`
- `style_id`
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
- [../examples/named_styles_showcase.yaml](../examples/named_styles_showcase.yaml)
