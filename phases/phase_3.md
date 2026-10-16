# Phase 3 — Word MCP Server ([SN] contract)

Active spec: spec.md §9 (R10–R12, N8–N9, D23–D25, Q5–Q6) | Plan: TBD (writing-plans after DIR approves this contract)

## Goal

A dual-mode Word server ships beside Excel and PowerPoint, proven by ONE gate: `word-mcp\.venv\Scripts\python.exe -m pytest word-mcp/tests -q` green from the repo root (`needs_word` passes where Word is present, skips cleanly elsewhere, never errors).

## Inputs

- spec.md §9 (R10–R12, N8–N9, D23–D25, Q5–Q6)
- phases/phase_2.md (P2.2 PPT server shape this mirrors: scaffold → core → server → fixture)
- ppt-mcp/server.py, ppt-mcp/core/guard.py, ppt-mcp/core/versions.py, ppt-mcp/core/audit.py, ppt-mcp/core/engine.py (patterns Word mirrors)
- sentinel/howtowork.md (DS check-per-save, OP watch protocol, 3-strike breaker)

## Scope — In / Out

In (each traces to an R):

- `word-mcp/` dual-engine server with V1 memo ops and the shared paranoid flow (R10)
- Word headed visible with retry/fallback parity, headless default (R11)
- Golden memo fixture + IB theme asserts vs frozen golden (R12)

Out (each traces to an N):

- TOC auto-generation, track-changes/comments, mail merge, macros/VBA, legacy `.doc` (N8)
- PDF export (N9)
- Shared-venv restructure — own venv per N7 as extended (N7)
- `sentinel/` untouched (spec §1 non-goals)

## Steps

Scaffold first (`pyproject`, `.gitignore`, `config.py`, `styles/`, guard/versions/audit + their tests), DS runs `npm run check` per save with OP watching per `sentinel/howtowork.md`, then headless core + format, then headed + engine + shots + server wiring, then fixture + golden + stress. The single Acceptance gate invocation decides for the whole phase.

## Deliverables

- word-mcp/pyproject.toml (`python-docx` pinned, `mcp`, `pillow`, `pywin32`, `pytest`), word-mcp/.gitignore, word-mcp/config.py (`DENIED_PREFIXES: list[str]`, same denylist semantics), word-mcp/opencode_mcp.json.example (R10)
- word-mcp/styles/ib_doc.py: `IB_DOC_DEFAULTS: dict`, `get_doc_style(key: str) -> dict | str`; word-mcp/styles/custom.json (R10)
- word-mcp/core/guard.py: `is_path_allowed`, `make_ops_hash`, `check_approval` (denylist shapes per amended D9) (R10)
- word-mcp/core/versions.py + audit.py: sidecar backups 20 FIFO + append-only log (R10)
- word-mcp/core/headless.py: `list_structure`, `read_text`, `add_paragraph`, `set_heading`, `format_text`, `add_table`, `add_image`, `add_header_footer`, `apply_theme`, `set_page_setup`, `preview_diff` (R10)
- word-mcp/core/headed.py + engine.py: `is_word_available`, `headed_apply`, `dispatch(tool, params)` with visible headed + `headed_unavailable_fell_back` (R11)
- word-mcp/core/shots.py: `thumbnail`, `full_shot` (headless best-effort + headed capture) (R10)
- word-mcp/server.py: 4 reads (`list_structure`, `read_text`, `preview_diff`, `screenshot_doc`) + 9 gated (`add_paragraph`, `set_heading`, `format_text`, `add_table`, `add_image`, `add_header_footer`, `apply_theme`, `set_page_setup`, `rollback`) (R10)
- word-mcp/tests/fixtures/make_memo.py: `build(path: str) -> None`; `ib_memo.docx`; `golden_doc.json` (frozen); `test_doc.py::test_full_ib_grade_vs_golden`; `test_stress_doc.py` headed tests (R11, R12)

## Acceptance Criteria

- AC-1 (gate): `word-mcp\.venv\Scripts\python.exe -m pytest word-mcp/tests -q`, cwd repo root → exit 0; `needs_word` items pass on this box, skip (never error) where Word is absent.
- AC-2 (golden): `golden_doc.json` schema frozen as pinned in tests; fixture + `apply_theme` output assert navy/white headings, Arial body, Letter setup (prove: tests green).
- AC-3 (paranoid parity): write without approval → `needs_approval` + no write; approved write → `backup_id`; system-dir path → rejected; `rollback` restores byte-identical (prove: tests green).
- AC-4 (no leaks): repo holds no `audit.log`, no `*.versions/`, no `.bak` outside intended fixtures (prove: `git status --short` shows only intended files).

## Dependencies

Requires: phases/phase_2.md green (PPT patterns reused; latch `clear` in sentinel/status_sentinel.json). Unblocks: presets/polish phase (`style_save_as_preset` across all three servers), registry/packaging work.

## Rollback / Failure Handling

On a red gate: `git checkout -- word-mcp` for phase-touched files only (never `sentinel/`); live user documents untouched (tests use `tmp_path` plus committed fixture inputs only); remove leaked artefacts; re-run the gate. Protocol halt (`npm run block`) trips only via the OP 3-strike breaker per `sentinel/howtowork.md` — phase rollback itself never trips it.

---
[SN] Phase 3 contract, draft for DIR review. Self-check C1–C9 runs live before handoff; building starts only after DIR approves this file and writing-plans delivers the task plan.
