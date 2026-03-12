# Configuration

RusDox keeps styling and output behavior in config so your YAML stays focused on content.

## Where Config Lives

RusDox checks config in this order:

1. `./rusdox.toml`
2. `~/rusdox/config.toml`
3. built-in defaults

That means:

- user-level config gives you a personal default style
- local `./rusdox.toml` overrides that style for one project

## Fastest Way To Change Config

Use the wizard:

```bash
rusdox config wizard --level basic
```

Or the full editor:

```bash
rusdox config wizard --level advanced
```

Find the user config path:

```bash
rusdox config path
```

Show the effective config:

```bash
rusdox config show
```

Create a local project override:

```bash
rusdox config wizard --path ./rusdox.toml --level basic
```

## What Config Controls

### `[output]`

- DOCX output folder
- PDF output folder
- whether PDF preview is enabled

### `[typography]`

- default font family
- cover title size
- title size
- subtitle size
- hero size
- page heading size
- section size
- body size
- tagline size
- note size
- table text size
- metric label and value sizes

### `[spacing]`

- paragraph spacing before and after
- section spacing
- note spacing
- table cell paragraph spacing
- metric spacing

### `[colors]`

- body text colors
- accent colors
- success, warning, and risk colors
- table border color
- card surface colors

### `[table]`

- default table width
- metric card width
- border sizes
- PDF cell padding

### `[pdf]`

- page size
- margins
- default text size
- line height behavior
- text width tuning used by the preview renderer

## Manual Editing

If you prefer to edit the file directly:

```bash
rusdox config init --template --force
```

That writes a commented TOML template.

Every field is optional.

If you remove a field, RusDox falls back to the built-in default.

## Common Patterns

### One personal style for everything

Use only `~/rusdox/config.toml`.

### One special style for one project

Create `./rusdox.toml` in that project.

### DOCX only

Set:

```toml
[output]
emit_pdf_preview = false
```

Or override per run with:

```bash
rusdox mydoc.yaml --docx-only
```
