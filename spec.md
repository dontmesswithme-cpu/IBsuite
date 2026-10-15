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

## 7. Phase 1 Amendment — IB Hardening (P1.1 real-model golden + P1.2 headed stress)

Date: 2026-10-14. Status: draft for DIR review. Pre-Phase-1 text (§1–§6) traces via D1–D12 + §5 verbatim values (grandfathered; every NEW item below carries an ID).

### Requirements (each names its verifier)

- R1 Real-model golden fixture. `tests/fixtures/make_ib_model.py: build(path: str) -> None` generates `tests/fixtures/ib_model.xlsx` with sheets `IS`, `BS`, `CF`, `Comps`; cross-sheet 3-statement links; one chart on `Comps`; one SUMIFS pivot-summary (D3 fallback); print-fit + footer per §5. `apply_ib_table` over model ranges yields header `002060` white bold 10, body `Arial 10`, inputs `D9E1F2`, int `_-* #,##0_-;-* #,##0_-;-* "-"_-;_-@_-`, pct `%_0.0%`, grid `D9D9D9` thin, medium subtotal, double total, freeze top+first col, widths 14–18, height 15, `FitToPagesWide=1 Tall=0` landscape, `oddFooter.center.text == "&F | &A"` with sides `""`. Verifier: `test_realmodel.py::test_full_ib_grade_vs_golden` green + match vs frozen `tests/golden_ib.json`.
- R2 Formulas verbatim incl. circular. `write_range`/`set_formula` store any formula string verbatim (cross-sheet `=IS!B2-BS!B2`, `SUMIFS`, volatile ok); fixture carries one INTENTIONAL circular cell flagged `CIRCULAR_INTENTIONAL` (e.g. `CF!X1 =CF!X1+1`); `read_range` returns formula strings verbatim with no evaluation and no invented values (extends P0 A7). Never evaluate, never block, never crash — headless or headed. Verifier: `test_realmodel.py::test_formulas_verbatim_and_circular_safe` green (write→reopen→equal; openpyxl load raises nothing).
- R3 Headed verifies on real model. `headed.py` opens strictly by path, never attaches the active workbook (closes archive Q3); `Visible=True DisplayAlerts=False ScreenUpdating=False` (DIR 2026-10-14: headed runs on-screen; focus-steal and no-parallel-headed-runs accepted, see D19); file-locked/Excel-busy → retry 200/500/1000ms then `{status:"needs_retry"}` with `.bak` restored, never half-written; detached COM → `{warning:"headed_unavailable_fell_back"}` + headless result, never crash; headed write of a real-model range roundtrips values+formulas. Verifier: `test_stress.py` headed tests green on this box, skipped (never error) where Excel is absent.
- R4 Regression pins (P0 A6–A12, now law). Single-range `headless` errors read `"bad cell range"`; `advanced`/`headed` multi-range errors read `"bad data range"`/`"bad source range"`/`"bad image anchor"`; `add_pivot` REFUSES an existing dest sheet; oversize >100000 cells raises `split into batches` with file bytes untouched; `golden_ib.json` top-level keys frozen to exactly `fills, number_formats, fonts, borders, layout` (schema change requires its paired test update in the same change). Verifier: `test_realmodel.py::test_regression_pins` green + all P0 suites stay green.
- R5 Shots on real model. `thumbnail` of an `ib_model.xlsx` range is valid base64 PNG <40000 chars and shorter than `full_shot` of the same range. Verifier: `test_realmodel.py::test_shots_on_real_model` green.

### Non-goals

- N1 No calc engine: no evaluation, no dependency graph, no circular-detection UI (R2 stores verbatim only).
- N2 No `style_save_as_preset(name)` (deferred to presets/polish phase).
- N3 No Word/PowerPoint, no HTTP/SSE transport (core stays transport-agnostic per D6).
- N4 No real client data in fixtures: synthetic only.

### Locked decisions (D numbering continues)

- D13 One phase file `phases/phase_1.md` carries P1.1+P1.2 sub-phases under a SINGLE gate invocation (full suite, one command).
- D14 Fixture is synthetic with its builder committed (`make_ib_model.py`), so any reviewer regenerates deterministically; the binary is also committed for zero-setup runs.
- D15 Circular = store-verbatim-never-evaluate (extends A7).
- D16 Headed = open-by-path-only (closes archive Q3).
- D17 `ALLOWED_ROOTS` default = repo root + `~/Documents` + `~/Downloads` (code + example already shipped 2026-10-14; closes archive Q1 as stated; per-write approval stays per D9).
- D18 Clients = Claude, Opencode, Codex, Antigravity + any MCP stdio client (README already reflects).

### Open questions (none blocking)

- Q1 Golden sample real vs synthetic? RESOLVED: synthetic (D14). Proof: gate green.
- Q2 Headed attach-active vs open-by-path? RESOLVED: open-by-path (D16). Proof: headed tests green.

## 8. Phase 2 Amendment — Visible Headed + PowerPoint (P2.1 + P2.2)

Date: 2026-10-14. Status: draft for DIR review.

### Requirements (each names its verifier)

- R6 Excel headed runs visible. `core/headed.py: _excel()` sets `Visible=True` (default, not opt-in); `DisplayAlerts=False` stays. Amends R3's flag in place. Verifier: `test_headed.py::test_headed_visible_default` green (COM-stubbed assert `Visible is True`, `DisplayAlerts is False`) + headed integration on the real model raises nothing.
- R7 PowerPoint server `ppt-mcp/` beside `excel-mcp/`, dual-engine: headless via `python-pptx` (pinned), headed via `PowerPoint.Application` COM on the same box. V1 ops: slides (add/list, layouts, 16:9 + 4:3 sizing), text (read/set), tables, charts, images, theme formatting (fonts, fills, IB navy/white). Same paranoid flow as Excel: whitelist `ALLOWED_ROOTS` (same default trio), exact-`ops_hash` approval, copy-before-write sidecar backups (20 FIFO), append-only audit, `{ok, backup_id, thumbnail}` returns. Verifier: `ppt-mcp` suite green (AC-2).
- R8 PPT headed is visible with shared Excel patterns: open-by-path only, `Visible=True`, busy → retry 200/500/1000ms then `{status:"needs_retry"}` + `.bak` restore, detached → `{warning:"headed_unavailable_fell_back"}` + headless result. Default engine `mode="headless"` on every tool. Verifier: `test_stress_ppt.py` headed tests green here, skipped (never error) elsewhere.
- R9 PPT golden deck fixture. `ppt-mcp/tests/fixtures/make_deck.py: build(path: str) -> None` generates `ib_deck.pptx` (title slide, 3-statement summary, comps table, chart slide, 16:9); theme asserts (heading font white bold on `002060`, body `Arial`, accounting-style tables) match frozen `golden_deck.json`. Verifier: `test_deck.py::test_full_ib_grade_vs_golden` green.

### Non-goals

- N5 Word deferred: PPT first per DIR 2026-10-14.
- N6 PPT deferred: SmartArt native authoring, animations/transitions, slideshow control, macros/VBA.
- N7 No shared-venv restructure: each server keeps its own venv and suite.

### Locked decisions (D numbering continues)

- D19 `Visible=True` per DIR overrule of R3's flag (costs accepted: Excel window takes focus during headed ops, headed runs cannot parallelise, headed is demo/operator mode — CI stays headless-only via `needs_excel` skips).
- D20 PPT headless lib = `python-pptx` (pinned; the only viable pure-python `.pptx` lib — mirrors the D1 `openpyxl` rationale).
- D21 PPT headed = `PowerPoint.Application` COM, visible, reusing the Excel retry/fallback/audit shapes.
- D22 `phases/phase_2.md` carries P2.1 (visible flip) + P2.2 (PPT server) as an ordered gate block (AC-1 then AC-2, both must be green). DIR's grouping directive supersedes the R8 split rule here; each sub-gate stays independently rerunnable.

### Open questions (none blocking)

- Q3 PPT charts in V1? RESOLVED: yes, via `python-pptx` chart parts + headed parity (D20–D21). Proof: gate green.
- Q4 Shared vs separate venv for `ppt-mcp/`? RESOLVED: separate (N7). Proof: AC-2 runs under `ppt-mcp/.venv`.

---
[SN] Genesis spec + Phase 1 amendment (§7) + Phase 2 amendment (§8, draft for DIR review). Active contracts in `phases/phase_0.md`, `phases/phase_1.md`, `phases/phase_2.md`. `conventions.md` still open (naming/commit rules) — to be locked before DS Terminal 1 if DIR wants it.
