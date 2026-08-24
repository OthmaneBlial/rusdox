# Word-native Templates

These three starter templates were authored as styled RTF documents and exported to DOCX by the macOS text system, so they exercise a package source outside RusDox. Each folder contains the editable source, the checked-in DOCX template, representative JSON, generated output, native PDF, and parity evidence.

All bundled template sources and sample data are licensed under the template
[MIT license](LICENSE.md). Placeholder syntax is documented in
[`docs/word-templates.md`](../docs/word-templates.md).

## Signed registry workflow

The templates are indexed in [`registry/index.json`](../registry/index.json)
instead of being embedded in the Rust crate. This keeps normal installs lean
while making every public entry discoverable with author credit, version bounds,
documented inputs, preview alt text, accessibility notes, and hashes for the
source and verified evidence.

    rusdox template list
    rusdox template add invoice

Before proposing a template, refresh its evidence and verify the complete
registry contract:

    ./scripts/verify_word_templates.sh
    node scripts/build_template_registry.mjs

The pull-request workflow publishes a review artifact containing the generated
registry page, screenshots, sample data, Word sources, and parity reports. Read
the [curation and signing contract](../docs/template-registry.md) before editing
registry metadata.

Regenerate the DOCX source packages on macOS:

    ./scripts/generate_word_templates.sh

Render every sample and refresh evidence:

    ./scripts/verify_word_templates.sh
