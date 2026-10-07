# Security Policy

## Supported versions

aioopenpyxl is pre-1.0. Only the latest released `0.x` version receives security fixes;
please upgrade before reporting.

| Version | Supported |
| --- | --- |
| latest 0.x release | yes |
| older releases | no |

## Reporting a vulnerability

Please do **not** open a public issue for security problems.

Use GitHub's private vulnerability reporting instead:
<https://github.com/yamaaaaaa31/aioopenpyxl/security/advisories/new>
("Security" tab of the repository, then "Report a vulnerability"). Include the affected version,
a description of the problem and, if possible, a minimal reproduction.

You should receive an acknowledgement within a week. Once a fix is available it is released as
a new version and the advisory is published; you will be credited unless you prefer not to be.

## Scope

aioopenpyxl is a thin wrapper: it does not parse `.xlsx` files itself but hands them to
openpyxl in worker threads. Issues in parsing untrusted spreadsheets (XML entity expansion,
zip bombs, ...) belong to [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) and its
optional `defusedxml` support. Report them upstream; a report here is welcome if the wrapper
makes an upstream mitigation unusable or bypasses it.
