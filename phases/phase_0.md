# Phase 0 — Excel MCP Foundation ([SN] contract)

Active spec: `spec.md`. Archive: `docs/superpowers/specs/2026-10-13-excel-mcp-design.md`. Plan: `docs/superpowers/plans/2026-10-13-excel-mcp.md`.

## P0.1 Base Architecture & Scaffolding (plan Tasks 1–3)

Goal: repo skeleton + trust primitives (guard, backup, audit). No Excel value logic yet.

Deliverable files:
- `excel-mcp/pyproject.toml`, `excel-mcp/.gitignore`, `excel-mcp/config.py`
- `excel-mcp/styles/ib.py`, `excel-mcp/styles/custom.json`
- `excel-mcp/core/guard.py`, `excel-mcp/core/versions.py`, `excel-mcp/core/audit.py`

Exported interfaces / types:
- `get_style(key: str) -> dict | str`; `resolve_style(key: str, override: dict | None)`; `ALLOWED_ROOTS: list[str]`
- `is_path_allowed(path: str) -> bool`; `make_ops_hash(ops: dict) -> str`; `check_approval(ops: dict, approval_token: str | None) -> dict`
- `backup(path: str, tool: str, label: str) -> str`; `list_versions(path: str) -> list[str]`; `rollback(path: str, backup_id: str) -> dict`; `append(tool, path, detail, backup_id, status) -> None`

Deterministic test cases (must pass, headless-only):
- `pytest excel-mcp/tests/test_styles.py -v` — header `002060`, override wins, roots contain project.
- `pytest excel-mcp/tests/test_guard.py -v` — traversal + outside-root rejected with no side effects; `None` token → `needs_approval`, hash token → allowed.
- `pytest excel-mcp/tests/test_versions.py -v` — backup→corrupt→rollback byte-identical; 21 backups → 20 retained FIFO.

## P0.2 Engines, Shots & Server (plan Tasks 4–9)

Goal: working MCP with headless + headed + IB formatting + screenshots, paranoid flow end-to-end.

Deliverable files:
- `excel-mcp/core/headless.py` (+`preview_diff`), `core/format_headless.py`, `core/advanced_headless.py`
- `excel-mcp/core/headed.py`, `excel-mcp/core/engine.py`, `excel-mcp/core/shots.py`
- `excel-mcp/server.py`, `excel-mcp/opencode_mcp.json.example`, `excel-mcp/tests/golden_ib.json`

Exported interfaces / types:
- `list_sheets(path)`, `read_range(path, sheet, cell_range)`, `write_range(path, sheet, cell_range, values)`, `set_formula(path, sheet, cell, formula)`, `add_sheet(path, name)`, `preview_diff(path, ops)`
- `format_range(path, sheet, cell_range, style)`, `apply_ib_table(path, sheet, cell_range, has_totals)`
- `add_chart(path, sheet, data_range, chart_type, anchor)`, `add_pivot(...)`, `add_image(path, sheet, image_path, anchor)`
- `is_excel_available()`, `headed_read(...)`, `headed_apply(path, ops)`, `dispatch(tool, params)`
- `thumbnail(path, sheet, cell_range) -> str`, `full_shot(path, sheet, cell_range, mode) -> str`
- MCP tools (13): `list_sheets, read_range, preview_diff, screenshot_sheet` + `write_range, set_formula, add_sheet, format_range, add_chart, add_pivot, add_image, macro_run_gated, rollback`

Deterministic test cases:
- `pytest excel-mcp/tests/test_headless.py -v` — roundtrip values/formulas; 200k-cell write rejected untouched; preview never writes.
- `pytest excel-mcp/tests/test_format.py -v` — IB header navy + accounting fmt vs `golden_ib.json`.
- `pytest excel-mcp/tests/test_advanced.py -v` — chart+image drawing present.
- `pytest excel-mcp/tests/test_headed.py excel-mcp/tests/test_engine.py -v` — dispatch defaults headless; headed detached → fallback warning (Excel-marked skipped OK).
- `pytest excel-mcp/tests/test_shots.py -v` — thumbnail <40KB and < full.
- `pytest excel-mcp/tests/test_server.py -v` — macro without flag never executes; write needs approval then succeeds with `backup_id`.

DS order: P0.1 files first, `npm run check` per save, then P0.2 in listed order. OP audits each save against this contract; 3rd consecutive tap → `npm run block`, HALT for DIR.

---
[SN] Phase 0 contract locked. Retiring SN. Ready for Terminal 1 (DS) + Terminal 2 (OP) Sentinel execution.
