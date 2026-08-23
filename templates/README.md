# Word-native Templates

These three starter templates were authored as styled RTF documents and exported to DOCX by the macOS text system, so they exercise a package source outside RusDox. Each folder contains the editable source, the checked-in DOCX template, representative JSON, generated output, native PDF, and parity evidence.

All bundled template sources and sample data are licensed under the repository MIT license. Placeholder syntax is documented in docs/word-templates.md.

Regenerate the DOCX source packages on macOS:

    ./scripts/generate_word_templates.sh

Render every sample and refresh evidence:

    ./scripts/verify_word_templates.sh
