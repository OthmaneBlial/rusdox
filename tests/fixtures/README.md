`external-macos-textutil.docx` is a small package produced outside RusDox by converting `external-macos-source.rtf` with the macOS text system. The integration suite opens it, modifies the body, saves it, validates content types and relationships, reopens it, and confirms untouched theme/meta parts remain.

Other tests generate DOCX archives in memory and in temporary locations. The checked-in external fixture exists specifically to prevent every regression case from sharing RusDox's own writer assumptions.
