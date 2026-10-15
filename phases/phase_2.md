# Phase 2 — Visible Headed + PowerPoint Server ([SN] contract)

Active spec: spec.md §8 (R6–R9, N5–N7, D19–D22, Q3–Q4) | Plan: TBD (writing-plans after DIR approves this contract)

## Goal

Headed Excel runs on-screen and a dual-mode PowerPoint server ships, proven by the ordered gate block: the excel suite green (AC-1) then the ppt suite green (AC-2), both from the repo root.

## Inputs

- spec.md §8 (R6–R9, N5–N7, D19–D22, Q3–Q4)
- phases/phase_1.md (R3/D16 headed patterns, fixture conventions)
- excel-mcp/core/headed.py, excel-mcp/tests/test_headed.py
- excel-mcp/server.py, excel-mcp/core/guard.py, excel-mcp/core/versions.py, excel-mcp/core/audit.py, excel-mcp/core/engine.py (patterns PPT mirrors)
- sentinel/howtowork.md (DS check-per-save, OP watch protocol, 3-strike breaker)

## Scope — In / Out

In (each traces to an R):

- Headed-visible flip for Excel + stubbed and integration tests (R6)
- `ppt-mcp/` dual-engine server with V1 deck ops and the Excel paranoid flow (R7)
- PPT headed visible with retry/fallback parity, headless default (R8)
- Golden deck fixture + IB theme asserts vs frozen golden (R9)

Out (each traces to an N):

- Word (N5)
- PPT SmartArt/animations/slideshow/macros (N6)
- Shared-venv restructure (N7)
- `sentinel/` untouched (spec §1 non-goals)

## Steps

P2.1 files first (`core/headed.py` one-line flip, `tests/test_headed.py` visible asserts), DS runs `npm run check` per save with OP watching per `sentinel/howtowork.md`, then P2.2 files (`ppt-mcp/` scaffold → core → server → tests/fixtures in dependency order). AC-1 decides P2.1, AC-2 decides P2.2; the phase passes only with both green, in order.

## Deliverables

- excel-mcp/core/headed.py: `_excel()` sets `Visible=True` (no interface change) (R6)
- excel-mcp/tests/test_headed.py: `test_headed_visible_default` (COM-stubbed `Visible is True`, `DisplayAlerts is False`) (R6)
- ppt-mcp/pyproject.toml, ppt-mcp/.gitignore, ppt-mcp/config.py (`ALLOWED_ROOTS: list[str]`, same default trio), ppt-mcp/styles/ib_deck.py + custom.json (R7)
- ppt-mcp/core/guard.py: `is_path_allowed`, `make_ops_hash`, `check_approval` (same shapes as Excel) (R7)
- ppt-mcp/core/versions.py + audit.py: sidecar backups 20 FIFO + append-only log (same shapes) (R7)
- ppt-mcp/core/headless.py: `list_slides`, `read_text`, `add_slide`, `set_text`, `format_shape`, `add_table`, `add_chart`, `add_image`, `apply_theme`, `preview_diff` (R7)
- ppt-mcp/core/headed.py + engine.py: `is_powerpoint_available`, `headed_apply`, `dispatch(tool, params)` with visible headed + `headed_unavailable_fell_back` (R8)
- ppt-mcp/core/shots.py: `thumbnail`, `full_shot` (R7)
- ppt-mcp/server.py: 4 reads (`list_slides`, `read_text`, `preview_diff`, `screenshot_slide`) + 8 gated (`add_slide`, `set_text`, `format_shape`, `add_table`, `add_chart`, `add_image`, `apply_theme`, `rollback`) (R7)
- ppt-mcp/tests/fixtures/make_deck.py: `build(path: str) -> None`; `ib_deck.pptx`; `golden_deck.json` (frozen); `test_deck.py::test_full_ib_grade_vs_golden`; `test_stress_ppt.py` headed tests (R8, R9)

## Acceptance Criteria

- AC-1 (P2.1 gate): `excel-mcp\.venv\Scripts\python.exe -m pytest excel-mcp/tests/ -q`, cwd repo root → exit 0 including `test_headed_visible_default`; `needs_excel` items pass on this box, skip (never error) elsewhere.
- AC-2 (P2.2 gate): `ppt-mcp\.venv\Scripts\python.exe -m pytest ppt-mcp/tests -q`, cwd repo root → exit 0; `needs_powerpoint` items (if so marked) pass here, skip (never error) elsewhere.
- AC-3 (deck golden): `golden_deck.json` schema frozen as pinned in tests; fixture + `apply_theme` output assert navy/white headings, body font, 16:9 size (prove: tests green).
- AC-4 (no leaks): repo holds no `audit.log`, no `*.versions/`, no `.bak` outside intended fixtures (prove: `git status --short` shows only intended files).

## Dependencies

Requires: phases/phase_1.md approved and green (headed patterns, latch `clear` in sentinel/status_sentinel.json). Unblocks: presets/polish phase and Word scoping.

## Rollback / Failure Handling

On a red gate: `git checkout -- excel-mcp/core excel-mcp/tests ppt-mcp` for phase-touched files only (never `sentinel/`); live user decks/workbooks untouched (tests use `tmp_path` plus committed fixture inputs only); remove leaked artefacts; re-run the gate block from AC-1. Protocol halt (`npm run block`) trips only via the OP 3-strike breaker per `sentinel/howtowork.md` — phase rollback itself never trips it.

---
[SN] Phase 2 contract, draft for DIR review. Self-check C1–C9 runs live before handoff; building starts only after DIR approves this file and writing-plans delivers the task plan.

---

## DIR Executive Decision — 2026-10-15: whitelist purged, allow-by-default denylist

DIR overrules the `ALLOWED_ROOTS` whitelist wherever this contract names it
(Deliverables `config.py` / `core/guard.py` lines above, via spec §8 R7):
both servers now allow every path except `DENIED_PREFIXES` (system dirs,
other users' profiles, drive roots, `sentinel/`; `..` resolves before the
check). Rationale: Claude Desktop could not write under any allowed root.
Approval gate, backups, and audit are unchanged. Implemented in `1896f1f`
(excel suite 123 passed, ppt suite 65 passed). Existing sections above are
left verbatim as build history; `spec.md` D9/D17/R7 carry the amended wording.
