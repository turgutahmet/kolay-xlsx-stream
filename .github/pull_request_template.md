## What

A brief description of the change. One or two sentences.

## Why

What problem does this solve? Link to the issue if there is one.

If there's no issue, explain the use case — non-trivial PRs without
prior discussion may be closed with a request to open an issue first
(see CONTRIBUTING.md).

## How

Key implementation decisions. Anything reviewers should look out for —
tradeoffs, alternatives considered, areas you're uncertain about.

## Tests

- [ ] Added tests covering the new behavior
- [ ] All existing tests pass locally (`composer test`)
- [ ] Tested with a real XLSX output opened in Excel / LibreOffice
      (for changes affecting cell rendering, styling, or schema)

## Performance impact

For changes touching the hot path (`writeRow`, sanitization, ZIP/XML
generation, S3 multipart):

- [ ] Ran `php benchmark-comprehensive.php` before and after
- [ ] Memory profile unchanged or improved
- Before: <!-- e.g. 192K rows/s, 2MB peak -->
- After:  <!-- e.g. 195K rows/s, 2MB peak -->

If this PR doesn't touch the hot path, delete this section.

## Checklist

- [ ] CHANGELOG.md updated under `[Unreleased]`
- [ ] No backwards-incompatible changes (or UPGRADE.md updated)
- [ ] Public API additions have at least one test
- [ ] Documentation updated if behavior or signatures changed
- [ ] Code follows PSR-12 and existing patterns in the codebase

## Screenshots / Output

If your change affects XLSX output (styling, formatting, layout),
attach a screenshot of the resulting file opened in Excel.

## Related

- Issue: #
- Discussion: #
- Related PRs: #
