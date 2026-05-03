# Security Policy

## Supported Versions

The following versions of `kolay/xlsx-stream` receive security updates:

| Version | Supported          |
|---------|--------------------|
| 2.2.x   | :white_check_mark: |
| 2.1.x   | :white_check_mark: |
| 2.0.x   | :white_check_mark: |
| 1.x     | :x:                |

Users on v1.x are encouraged to upgrade to v2.x. See [UPGRADE.md](UPGRADE.md)
for migration steps.

## Reporting a Vulnerability

If you discover a security vulnerability in `kolay/xlsx-stream`, please report
it privately. **Do not open a public GitHub issue.**

Send a report to: **turgutahmt@gmail.com**

Please include:

- A description of the vulnerability and its potential impact
- Steps to reproduce the issue
- Affected version(s)
- Any relevant code snippets, exploits, or PoCs

You can expect:

- An acknowledgment within 72 hours
- An initial assessment within 7 days
- A coordinated disclosure timeline if the issue is confirmed

## Scope

In-scope:

- The `kolay/xlsx-stream` package code
- The included S3 multipart upload implementation
- XML/ZIP construction logic that processes user-supplied data

Out-of-scope:

- Vulnerabilities in upstream dependencies (`aws/aws-sdk-php`,
  Laravel framework). Report those to the respective projects.
- Issues requiring administrator-level access to the host system.
- Denial of service via legitimate but very large inputs (the
  package is designed to handle millions of rows; if you find a
  configuration that causes unexpected memory growth, that's a
  bug not a vulnerability).

## Hall of Fame

If you've responsibly reported a security issue, you'll be credited
in the relevant CHANGELOG entry (with your permission).
