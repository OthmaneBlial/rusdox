# Governance and maintainer policy

RusDox is currently maintained by Othmane Blial. This policy is published before
adding committers so access decisions remain predictable rather than social or
ad hoc.

## Decision model

The maintainer is accountable for releases, security response, compatibility
claims, repository settings, and final merge decisions. Product decisions favor
the documented contracts: editable DOCX, native PDF, measurable parity,
local-first integrations, bounded inputs, and honest evidence.

Normal changes are discussed in issues or pull requests. Decisions that alter a
stable spec, protocol, template syntax, security boundary, compatibility claim,
or release process require a written rationale in the pull request and the
corresponding documentation or migration note. Private vulnerability reports
follow `SECURITY.md`, not public consensus.

## Contribution and review

- Anyone may open an issue, Discussion, or pull request under the Code of Conduct.
- Contributors retain authorship credit in Git history and release notes.
- Pull requests should be small enough to review, include proportionate tests,
  and disclose generated artifacts or viewer evidence.
- The author does not approve their own change when another active maintainer is
  available. Until then, CI evidence and an explicit self-review checklist are
  required.
- Being declined on one proposal does not reduce standing or future access.

## Path to committer

Commit access is never granted for star count, employment, sponsorship, or a
single large contribution. A candidate should demonstrate, over multiple merged
changes:

1. sound judgment across code, tests, documentation, and user impact;
2. respect for privacy, compatibility, accessibility, and scope boundaries;
3. constructive reviews of other contributors;
4. reliable follow-through on regressions they introduced;
5. understanding of the release and security policies.

The current maintainer proposes the role publicly in a governance issue. The
candidate must accept, enable strong account security, and begin with scoped
write access. Maintainer status is recorded in this file and `CODEOWNERS`.

## Releases and security

Only maintainers may publish crates, tags, GitHub Releases, Pages, or signed
artifacts. Releases require green protected CI and the evidence listed in the
release checklist. Credentials must be short-lived or platform-managed; secrets
must not be committed, logged, or copied into issues.

Security reports are acknowledged and handled according to `SECURITY.md`.
Embargoed details remain limited to the smallest response group until a fix and
coordinated disclosure are ready.

## Inactivity, removal, and conflicts

A maintainer may step down at any time. Access may be reduced after six months
of inactivity following a private check-in, or removed immediately for a
security risk or Code of Conduct enforcement. Removal does not erase authorship
credit.

Maintainers disclose material conflicts and recuse from decisions where they
cannot be impartial. Governance disputes are documented in an issue; if there is
only one maintainer, that maintainer records the final decision and rationale.

## Policy changes

Changes to this policy use a normal pull request with at least seven days for
community comment, except urgent security corrections. The changelog records
material governance changes.
