# Security Policy

## Supported Versions

RusDox is currently maintained as a fast-moving project.

Security fixes are most likely to land in:

- the latest release line
- the current `master` branch

Older versions may not receive fixes.

## Reporting A Vulnerability

Please do not publish sensitive vulnerability details in a public issue.

Preferred path:

1. Use GitHub private vulnerability reporting if it is enabled for the repository.
2. If that is not available yet, open a minimal public issue asking for a private security contact channel without including exploit details.

When reporting, include:

- affected version or commit
- impact summary
- reproduction steps
- whether the issue affects DOCX parsing, DOCX writing, YAML parsing, PDF rendering, or the CLI
- any proof-of-concept details that help confirm the issue

## What To Expect

- an acknowledgement as soon as practical
- clarification questions if needed
- a fix or mitigation plan when the report is confirmed

## Scope

Examples of relevant security issues:

- unsafe parsing behavior with untrusted input
- archive handling bugs that could overwrite files unexpectedly
- vulnerabilities caused by generated output or package preservation behavior

Non-security bugs should go through the normal issue tracker.
