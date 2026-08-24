# Curated template registry

RusDox keeps the template registry outside the published Rust crate. The engine
stays lean; templates can carry Word source, sample JSON, previews, verified
DOCX/PDF output, and parity evidence without increasing `cargo install` size.

## Discover and install

    rusdox template list
    rusdox template search compliance
    rusdox template add board-report
    rusdox template update board-report
    rusdox template update --all

List and search are read-only. Add and update download only the selected
`template.docx`, `data.json`, and manifest into the user template data folder.
Every downloaded byte is checked against the SHA-256 hash in the registry.

The default registry index is signed with Ed25519. RusDox pins the registry
public key and verifies the detached signature before it trusts metadata or
downloads. Private registries can pass a local file/HTTPS index and an explicit
public key. Plain HTTP is rejected except for loopback testing.

## Curated acceptance contract

Every entry must include:

- author credit and an SPDX license;
- one or more of invoices, proposals, reports, compliance, HR, education, or operations;
- a screenshot with meaningful alt text;
- supported RusDox version bounds and documented inputs;
- sample JSON and a Word template with SHA-256 hashes;
- verified DOCX, PDF, parity JSON, and parity HTML hashes;
- language, reading-order review, non-color-only meaning, and accessibility notes.

Run the registry verifier before proposing a template:

    node scripts/build_template_registry.mjs
    ./scripts/verify_word_templates.sh

The generator validates the contract and writes the preview page used by pull
requests and the public site. A template PR never needs to modify OOXML engine
code.

## Trust and privacy

The registry contains public templates only. RusDox never uploads local sample
data or rendered documents. Installation is explicit and writes only beneath
the selected install root. Hash or signature mismatches fail before replacement,
and atomic writes preserve an existing installed template.

The initial signing key was generated only to sign the immutable v1 index and
then discarded. This conservative bootstrap means a changed default registry
requires a new RusDox release with a newly pinned key; the remote host alone
cannot silently replace the trusted index.
