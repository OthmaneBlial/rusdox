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

The default v1 registry lives at `registry/v1/index.json` and is signed with
Ed25519. RusDox 1.x pins that versioned URL and its public key, then verifies the
detached signature before it trusts metadata or downloads. The legacy
`registry/index.json` and its key remain byte-for-byte available for 0.1.1
clients. Private registries can pass a local file/HTTPS index and an explicit
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

    RUSDOX_TEMPLATE_EVIDENCE_ROOT=registry/v1/evidence ./scripts/verify_word_templates.sh
    node scripts/build_v1_registry_manifest.mjs --check
    node scripts/build_template_registry.mjs --registry-dir registry/v1 --check

The generator validates the contract and writes the preview page used by pull
requests and the public site. A template PR never needs to modify OOXML engine
code.

## Trust and privacy

The registry contains public templates only. RusDox never uploads local sample
data or rendered documents. Installation is explicit and writes only beneath
the selected install root. Hash or signature mismatches fail before replacement,
and atomic writes preserve an existing installed template.

Each default-channel signing key is generated in memory only to sign its
immutable index and is then discarded. A changed default registry therefore
requires a new versioned directory, a newly pinned key, and a RusDox release;
the remote host alone cannot silently replace either the legacy or v1 index.
`scripts/create_immutable_registry_signature.mjs` refuses to overwrite an
existing signature or public key.
