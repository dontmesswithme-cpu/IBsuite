# Phase 1 — IB Hardening: Real-Model Golden + Headed Stress ([SN] contract)

Active spec: spec.md §7 (R1–R5, N1–N4, D13–D18, Q1–Q2) | Plan: TBD (writing-plans after DIR approves this contract)

## Goal

Excel survives daily IB use, proven by ONE gate: `excel-mcp\.venv\Scripts\python.exe -m pytest excel-mcp/tests/ -q` green from the repo root (`needs_excel` passes where Excel Desktop is present, skips cleanly elsewhere, never errors).

## Inputs

- spec.md §7 (R1–R5, N1–N4, D13–D18, Q1–Q2)
- phases/phase_0.md (P0.1+P0.2 interfaces, A1–A12)
- excel-mcp/core/headless.py, excel-mcp/core/headed.py, excel-mcp/core/format_headless.py, excel-mcp/core/advanced_headless.py, excel-mcp/core/guard.py, excel-mcp/core/versions.py, excel-mcp/core/shots.py, excel-mcp/core/engine.py
- excel-mcp/config.py, excel-mcp/styles/ib.py, excel-mcp/server.py
- excel-mcp/tests/test_format.py, excel-mcp/tests/golden_ib.json, excel-mcp/tests/test_headed.py
- sentinel/howtowork.md (DS check-per-save, OP watch protocol, 3-strike breaker)

## Scope — In / Out

In (each traces to an R):

- Real-model fixture + deterministic builder (R1)
- Full IB-grade formatting asserted vs frozen golden (R1)
- Formulas verbatim incl. one intentional circular cell (R2)
- Headed open-by-path verify + busy/detached handling on the real model (R3)
- Regression pins: error strings, pivot dest refusal, 100k untouched, frozen golden schema (R4)
- Shots on the real model (R5)

Out (each traces to an N):

- Calc engine, evaluation, circular-detection UI (N1)
- `style_save_as_preset(name)` (N2)
- Word/PowerPoint, HTTP/SSE transport (N3)
- Real client data in fixtures (N4)
- `sentinel/` untouched (spec §1 non-goals)

## Steps

P1.1 files first (`tests/fixtures/make_ib_model.py`, `tests/fixtures/ib_model.xlsx`, `tests/test_realmodel.py`, frozen `tests/golden_ib.json`, tightened `tests/test_format.py`), DS runs `npm run check` per save with OP watching per `sentinel/howtowork.md`, then P1.2 files (`tests/test_stress.py`, headed pins; `core/headed.py` and `core/format_headless.py` touched ONLY if the gate is red, fix-to-green with no new exported interfaces). The single Acceptance gate invocation decides for both sub-phases together.

## Deliverables

- excel-mcp/tests/fixtures/make_ib_model.py: `build(path: str) -> None` (R1)
- excel-mcp/tests/fixtures/ib_model.xlsx: sheets `IS`/`BS`/`CF`/`Comps`, chart, SUMIFS summary, one cell flagged `CIRCULAR_INTENTIONAL` (R1, R2)
- excel-mcp/tests/test_realmodel.py: `test_full_ib_grade_vs_golden`, `test_formulas_verbatim_and_circular_safe`, `test_regression_pins`, `test_shots_on_real_model` (R1, R2, R4, R5)
- excel-mcp/tests/test_stress.py: `test_headed_realmodel_roundtrip`, `test_busy_locked_needs_retry_no_halfwrite`, `test_detached_falls_back` (R3)
- excel-mcp/tests/golden_ib.json: top-level keys frozen to exactly `fills, number_formats, fonts, borders, layout` (R1, R4)
- excel-mcp/tests/test_format.py: tightened to the frozen schema; footer asserts `oddFooter.center.text == "&F | &A"` with sides `""` (R1, R4)
- (conditional) excel-mcp/core/format_headless.py, excel-mcp/core/headed.py: fix-to-green only, no interface changes (R1, R3)

## Acceptance Criteria

- AC-1 (THE gate): `excel-mcp\.venv\Scripts\python.exe -m pytest excel-mcp/tests/ -q`, cwd repo root → exit 0; every headless test passes including the 4 new realmodel tests; `needs_excel` items pass on this box and skip (never error) where Excel is absent.
- AC-2 (frozen schema): `golden_ib.json` top-level keys are exactly `fills, number_formats, fonts, borders, layout` (prove: load JSON, assert key-set equality); `test_format.py` reads no other top-level keys (prove: grep).
- AC-3 (footer verbatim): fixture and any `apply_ib_table` output carry `oddFooter.center.text == "&F | &A"` with left/right `""`, so concatenating the three sections equals `"&F | &A"` (prove: test assert).
- AC-4 (no leaks): repo holds no `audit.log`, no `*.versions/`, no `.bak` outside intended fixtures (prove: `git status --short` shows only intended files).

## Dependencies

Requires: phases/phase_0.md green (110/110) with latch `clear` in sentinel/status_sentinel.json. Unblocks: presets/polish phase (`style_save_as_preset`) and Word scoping, which reuse the headed patterns pinned here.

## Rollback / Failure Handling

On a red gate: `git checkout -- excel-mcp/tests excel-mcp/core` for phase-touched files only (never `sentinel/`); live user workbooks stay untouched (tests use `tmp_path` plus committed fixture inputs only); remove any leaked artefacts; re-run the gate. Protocol halt (`npm run block`) trips only via the OP 3-strike breaker per `sentinel/howtowork.md` — phase rollback itself never trips it.

---
[SN] Phase 1 contract, draft for DIR review. Self-check C1–C9 runs live before handoff; retirement into writing-plans needs DIR approval of this file.
