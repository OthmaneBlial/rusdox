# RusDox Authoring for VS Code

This zero-dependency extension targets RusDox YAML, JSON, and TOML files.

It provides:

- JSON Schema validation and autocomplete for JSON specs;
- block-type and top-level-field completion across all three formats;
- hover documentation for the versioned authoring contract;
- inline diagnostics for missing/unsupported versions, unknown blocks,
  unterminated expressions, and unsupported filters.

For full schema-driven YAML completion, install Red Hat YAML and map
schema/rusdox-spec-v1.schema.json to *.rusdox.yaml. The bundled diagnostics
remain local and require no language server or network request.

Run the dependency-free service tests with npm test.

The extension never reads document content outside VS Code and never uploads it.
