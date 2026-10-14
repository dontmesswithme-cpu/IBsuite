# Spec — Excel MCP (Sentinel Genesis, [SN])

Date: 2026-10-13
Archive: `docs/superpowers/specs/2026-10-13-excel-mcp-design.md` (brainstorm record, read-only)
Plan: `docs/superpowers/plans/2026-10-13-excel-mcp.md` (9 tasks, TDD)

## 1. Project Identity & Domain

- Purpose: local MCP server `excel-mcp/` letting the assistant create/edit real `.xlsx` from chat — headless (file lib, no Excel) + headed (COM driving Excel Desktop on same Windows box).
- Domain: `backend/API` + `CLI tool` (stdio MCP server, Python).
- User: single user (Faraz), Windows + Excel Desktop, client Opencode here.
- Non-goals V1: no Excel Online/Mac/Office.js, no remote HTTP (core kept transport-agnostic for later), no touch to `sentinel/`.

## 2. Tech Stack & Tooling

- Language: Python `>=3.11`. PM: pip via `excel-mcp/pyproject.toml`.
- Libs (pinned): `mcp` (Python SDK), `openpyxl`, `pywin32`, `pillow`, `pytest`.
- Greenfield: `excel-mcp/` new beside `sentinel/`. No existing Excel code.
- Test & verify: `pytest excel-mcp/tests/ -k "not needs_excel" -v` (CI, headless-only); `pytest excel-mcp/tests/test_headed.py -v -m needs_excel` when Excel present (skipped otherwise).

## 3. Locked Decisions (D1–D12)

- D1 headless = file `.xlsx` via `openpyxl`. D2 head = `Excel.Application` COM, same box.
- D3 full power: values, formulas, sheets, formatting, tables, validation, conditional, charts, pivots, images, print.
- D4 IB-grade defaults verbatim (§5). D5 `styles/ib.py` + `styles/custom.json` override + per-call override.
- D6 local stdio V1; `server.py` thin, `core/` transport-agnostic.
- D7 stack above. D8 single server dual-engine (rejected split-server).
- D9 paranoid-first: reads always, writes return `needs_approval` until exact `ops_hash` approved; whitelist `ALLOWED_ROOTS`.
- D10 macros: read freely, execute only with `macro_approved:true` + approval hash; log SHA256 only, never auto-run.
- D11 review: diff always, thumbnail every change, full PNG on `show me`, version history + rollback.
- D12 backups: sidecar `<Stem>.versions/<ts>_<tool>_<label>.bak` + `.json`, 20 FIFO, git-ignored.

## 4. Architecture & Module Boundaries

```
excel-mcp/
  server.py, config.py
  core/engine.py, headless.py, headed.py, guard.py, versions.py, audit.py, shots.py, format_headless.py, advanced_headless.py
  styles/ib.py, styles/custom.json
  tests/test_*.py, tests/golden_ib.json
```

- `server.py`: stdio only, no Excel logic. Flow per gated call: `guard.is_path_allowed` → `guard.check_approval` → `versions.backup` → `engine.dispatch` → `audit.append` → `{ok, backup_id, thumbnail}`.
- `core/`: no MCP imports. `headed.py` single-thread COM, retry 200/500/1000ms, atomic restore `.bak` on exception, fallback warning `headed_unavailable_fell_back`.
- Max 100000 cells/write, else `rejected split`. Errors → `audit.log` with `ERROR`, never half-write.
- Storage: live `.xlsx` + sidecar `.versions/` + `audit.log`, both git-ignored.

## 5. IB Styles (§6 of archive)

Header `002060` white bold 10, body `Arial 10`, inputs `D9E1F2`; int `_-* #,##0_-;-* #,##0_-;-* "-"_-;_-@_-`, pct `%_0.0%`; grid `D9D9D9` thin, medium subtotal, double total; freeze top+first col, widths 14–18, height 15, `FitToPagesWide=1 Tall=0` landscape footer `&F | &A`.

## 6. MCP Surface (13 tools)

Read (4, always): `list_sheets`, `read_range`, `preview_diff` (never writes), `screenshot_sheet`.
Gated (9): `write_range`, `set_formula`, `add_sheet`, `format_range`, `add_chart`, `add_pivot`, `add_image`, `macro_run_gated`, `rollback` (itself gated, backups pre-rollback first).
Every tool takes `mode: headless|headed` (default `headless`).

---
[SN] Genesis spec. Active contracts in `phases/phase_0.md`. `conventions.md` still open (naming/commit rules) — to be locked before DS Terminal 1 if DIR wants it.
