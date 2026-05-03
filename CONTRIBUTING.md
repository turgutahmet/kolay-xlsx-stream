# Contributing

Contributions are welcome. This is a small, focused package — please
keep that in mind when proposing changes.

## Before opening a PR

1. **Open an issue first** for non-trivial changes. The
   [ROADMAP](ROADMAP.md) describes what's planned and what's
   intentionally out of scope.
2. **Run the test suite locally:** `composer test`
3. **Match the existing style** — PSR-12, no Laravel Pint config
   yet (planned for v2.3).

## What I'm looking for

- Bug fixes with regression tests
- Performance improvements with before/after benchmarks
- Documentation improvements (README, CHANGELOG, examples)
- New sink implementations (e.g. for GCS, Azure Blob)

## What I'm not looking for

- Adding read support — this package is write-only by design.
  Use `box/spout` or `phpoffice/phpspreadsheet` for reading.
- Adding chart/image embedding — out of scope.
- Adding cell formulas — also out of scope.
- Cosmetic refactoring without functional change.

## Pull request checklist

- [ ] Tests pass (`composer test`)
- [ ] CHANGELOG.md has an entry under `[Unreleased]`
- [ ] New public API has at least one test
- [ ] If it changes behavior, UPGRADE.md notes the migration

## Reporting bugs / vulnerabilities

- Bug reports: open an issue with the bug template
- Security: see SECURITY.md (do not open a public issue)

Thanks for your interest in improving the package.
