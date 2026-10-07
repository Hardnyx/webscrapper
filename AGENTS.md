# Repository instructions

- This is a personal library of reusable scrapers. Keep company-specific
  eligibility, portfolio and scoring rules in consumer projects.
- Providers must remain independent. Shared infrastructure and explicit
  consumer-level joins are allowed.
- Use English Conventional Commit messages and English code comments.
  User-facing reports and documentation use Spanish.
- When access checks identify a page block from the execution environment,
  deliver a standalone page-function diagnostic script for the user to run
  on their PC. Include dependency version checks and install only missing or
  incompatible packages via `sys.executable`. Preserve TLS verification.
  Distinguish HTTP/WAF blocks, connection failures and parser/structure errors.
  Do not represent a blocked or incomplete capture as a successful scrape.
- Preserve unresolved or ambiguous entity matches. Name changes, mergers and
  type changes require explicit dated evidence; never infer legal eligibility
  from the current published list or from a name match.
