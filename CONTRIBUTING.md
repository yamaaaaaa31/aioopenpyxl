# Contributing to aioopenpyxl

Thanks for your interest. Bug reports, documentation fixes and code changes are all welcome.
This page explains how the project is set up and what a pull request needs.

Please note the project's goal before proposing features: aioopenpyxl exists so that openpyxl
can be used **without blocking the event loop and without violating openpyxl's thread-safety
assumptions**. It does not try to be faster than openpyxl, and we do not add performance claims
or benchmark figures to the documentation. Design decisions are argued from correctness and
safety.

## Development environment

The project is managed with [uv](https://docs.astral.sh/uv/). Python 3.10 to 3.14 are
supported; `.python-version` pins 3.14 for local development.

```sh
git clone https://github.com/yamaaaaaa31/aioopenpyxl.git
cd aioopenpyxl
uv sync            # creates .venv with the runtime and dev dependencies
```

To use another interpreter for a run, set `UV_PYTHON`, e.g. `UV_PYTHON=3.10 uv run pytest -q`.

## Running the tests

```sh
uv run pytest -q
```

Every `async` test is executed on **both anyio backends**: the `anyio_backend` fixture in
`tests/conftest.py` is parametrised with `asyncio` and `trio`, so there is nothing to select.
Tests run in random order (`pytest-randomly`), warnings are errors and each test has a 60 s
timeout (`pyproject.toml`, `[tool.pytest.ini_options]`). To rerun a failing order, pass the
seed pytest printed: `uv run pytest --randomly-seed=<seed>`, or `uv run pytest -p no:randomly`
to disable shuffling.

New behaviour needs tests. Concurrency fixes in particular should come with a test that
fails without the fix; see `tests/test_cancellation.py` and `tests/test_streaming.py` for the
style used (real threads, real cancellation, both backends). The suite is organised by topic:

- `tests/test_runner.py`: the `Runner`, the busy guard, the runner registry and `limiter=`;
- `tests/test_streaming.py`: `iter_rows` / `iter_cols` / `read_rows`, chunks, `prefetch` and
  stream clean-up;
- `tests/test_cancellation.py`: anyio and native asyncio cancellation;
- `tests/test_workbook.py` / `tests/test_worksheet.py`: the wrapper APIs;
- `tests/test_compat.py`: full-surface coverage of openpyxl's public members, sub-module
  aliases and module constants; `tests/test_typing.py` holds static `assert_type` checks.

Shared helpers (`ROWS`, the `sample_xlsx` fixture, `prefetch_threads()`, ...) live in
`tests/conftest.py`.

## Lint and type checking

```sh
uv run ruff check .
uv run ruff format .           # or `--check` to only verify
uv run ty check src tests
uv run mypy
```

CI runs all four and fails on any finding. `ruff` is configured in `pyproject.toml`
(line length 100, target Python 3.10). `ty` and `mypy` both target 3.10 as well, so avoid
syntax or typing features newer than that in `src/` and `tests/`.

## Documentation

The README is deliberately short. The design rationale (semaphore, busy guard, cancellation,
limiter, registry) lives in `docs/design.md`, and the openpyxl compatibility details (forwarded
members, typing, sub-module aliases, subclass attribute rules, the full differences table) in
`docs/compatibility.md`. Put new detail there and link to it from the README rather than growing
the README. Do not add performance claims or benchmark figures anywhere.

## Generated code

Most of the public surface lives in `src/aioopenpyxl/_generated.py` and in the `.pyi` shims
next to the package (`src/aioopenpyxl/styles/__init__.pyi`, ...). Both are produced by
`tools/gen_proxies.py` from the `types-openpyxl` stubs. **Do not edit them by hand.**

Regenerate and reformat after any of these:

- changing `tools/gen_proxies.py` (including its `EXCLUDE` / `EXTRA_MEMBERS` tables),
- bumping the `openpyxl` or `types-openpyxl` pins or the lockfile,
- adding a hand-written member to a wrapper that should no longer be generated.

```sh
uv run python tools/gen_proxies.py
uv run ruff format src/aioopenpyxl/
git status --porcelain -- src/aioopenpyxl/   # commit everything that changed
```

The generator also performs self-checks (signature drift against the runtime, runtime members
unknown to the stubs, unresolvable defaults) and exits non-zero on drift. CI regenerates and
fails if the committed files differ, including untracked new `.pyi` files, so commit them.

## Pull requests

1. Open an issue first for anything beyond a small fix, so the approach can be discussed.
2. Create a branch from `main`, make the change, add tests.
3. Run the test suite, lint, type checks and (if relevant) the generator, as above.
4. Add an entry to `CHANGELOG.md` under `## [Unreleased]` (see below).
5. Open the pull request; the template has a checklist. CI must be green. Keep the PR focused
   on one change; unrelated reformatting makes review harder.

The `src/` wrappers are written to behave the same on asyncio and trio. If something can only
be done for one backend, say so in the PR and in the docstring.

## Changelog

The changelog follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html) (pre-1.0: minor versions may
break the API, patch versions do not).

Add a bullet under `## [Unreleased]` in the matching subsection (`Added`, `Changed`,
`Deprecated`, `Removed`, `Fixed`, `Security`), creating the subsection if it does not exist.
Write it for users, not for reviewers: what changed, and what a caller sees differently.
Documentation-only and CI-only changes do not need an entry. The release workflow copies the
section for the released version into the GitHub Release notes, so keep it self-contained.

## Commit messages

- Imperative mood, short summary line (around 72 characters), optional body explaining *why*.
- Reference issues in the body (`Fixes #12`) rather than in the summary.
- One logical change per commit where practical; generator output can be its own commit
  ("Regenerate forwarders for types-openpyxl X.Y").

## Releasing (maintainers)

1. Move the `[Unreleased]` entries to a new `## [X.Y.Z] - YYYY-MM-DD` section and update the
   comparison links at the bottom of `CHANGELOG.md`.
2. Bump `version` in `pyproject.toml` and `__version__` in `src/aioopenpyxl/__init__.py`.
3. Commit, tag `vX.Y.Z` and push the tag. `.github/workflows/release.yml` builds the sdist and
   wheel, publishes them to PyPI via Trusted Publishing (the `pypi` environment) and creates a
   GitHub Release whose body is that changelog section.

## Code of Conduct

This project follows the [Contributor Covenant](CODE_OF_CONDUCT.md). By participating you
agree to abide by it.
