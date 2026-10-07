## Summary

<!-- What does this change and why? Link the issue if there is one: Fixes #123 -->

## Checklist

- [ ] Tests added or updated for the change; the async ones run on **both** backends
      (the `anyio_backend` fixture parametrises asyncio and trio automatically).
- [ ] `uv run pytest -q` passes locally.
- [ ] `uv run ruff check . && uv run ruff format --check .` pass.
- [ ] `uv run ty check src tests && uv run mypy` pass.
- [ ] If `tools/gen_proxies.py`, `EXCLUDE` / `EXTRA_MEMBERS`, or the openpyxl / `types-openpyxl`
      pins changed: ran `uv run python tools/gen_proxies.py && uv run ruff format src/aioopenpyxl/`
      and committed the regenerated `_generated.py` and `.pyi` shims.
- [ ] `CHANGELOG.md` has an entry under `[Unreleased]` (skip for docs-only or CI-only changes).
- [ ] No new performance claims or numbers in the docs; design decisions are explained by reasoning.
