# IBSuite MCP — IB-specialised Excel MCP

Local MCP server (`excel-mcp/`) that lets the assistant build IB-ready `.xlsx` from chat — 3-statement models, comps, schedules — headless (file lib, no Excel) + headed (COM driving Excel Desktop on the same Windows box).

IB-first, not generic Excel: navy headers, accounting formats, frozen panes, print-fit pitchbook pages by default, with house-style presets + per-call overrides so you never hand-format decimals again.

Single user, local stdio, paranoid-first: nothing writes without whitelist + explicit approval + backup + audit entry. Works with Claude, Opencode, Codex, Antigravity and any MCP client.

## Features

- Dual engine: headless via `openpyxl` (no Excel needed), headed via `Excel.Application` COM with 3x retry + fallback warning.
- 13 tools: 4 always-allowed reads (`list_sheets`, `read_range`, `preview_diff`, `screenshot_sheet`) + 9 gated writes (`write_range`, `set_formula`, `add_sheet`, `format_range`, `add_chart`, `add_pivot`, `add_image`, `macro_run_gated`, `rollback`).
- Gated flow per call: `is_path_allowed` → `check_approval` (exact `ops_hash`) → `backup` → `dispatch` → `audit.append` → `{ok, backup_id, thumbnail}`.
- IB presets + custom: `styles/ib.py` house defaults verbatim (navy `002060` white bold 10, body `Arial 10`, inputs `D9E1F2`, int `_-* #,##0_-;-* #,##0_-;-* "-"_-;_-@_-`, pct `%_0.0%`, thin `D9D9D9` grid, widths 14-18, row height 15, `FitToPagesWide=1 Tall=0` landscape footer `&F | &A`) → `styles/custom.json` overrides without code edits → per-call `style_override` wins. `apply_ib_table` one-shots a range IB-ready; `format_range` covers font/fill/border/number_format/alignment/freeze/print/table/validation/conditional.
- Backups: sidecar `<Stem>.versions/<ts>_<tool>_<label>.bak + .json`, 20 FIFO, git-ignored. `rollback` backs up pre-rollback state first, so rollbacks are undoable.
- Limits: 100,000 cells per write (else `rejected split`), `..` resolved before the check, everything allowed except the denylist (system dirs, other users' profiles, drive roots, `sentinel/`).

## Layout

```text
excel-mcp/
  server.py, config.py
  core/engine.py, headless.py, headed.py, guard.py, versions.py, audit.py, shots.py, format_headless.py, advanced_headless.py
  styles/ib.py, styles/custom.json
  tests/test_*.py, tests/golden_ib.json
sentinel/              # dual-agent protocol, never tracked (see .gitignore)
spec.md                # requirements + locked decisions D1-D12
phases/phase_0.md      # Phase 0 contract (P0.1 + P0.2, complete)
docs/superpowers/      # genesis design archive + plan
```

## Install (fresh clone)

```powershell
cd "Office plugin MCP\excel-mcp"
python -m venv .venv
.venv\Scripts\python.exe -m pip install -e .[test]
# Windows Excel Desktop path needs pywin32 (already pinned): included above
```

Requires Python `>=3.11`. Pinned: `mcp==2.3.0`, `openpyxl==3.1.5`, `pillow==12.3.0`, `pywin32==312`.

## Configure (Claude, Opencode, Codex, Antigravity + any MCP client)

Stdio server — same wiring everywhere. Copy `excel-mcp/opencode_mcp.json.example` into your client MCP config and fix paths:

```json
{ "EXCEL_MCP_DENIED_EXTRA": "" }
```

No path whitelist: every path is allowed except `DENIED_PREFIXES` (system dirs, other users' profiles, drive roots, `sentinel/`; extend via `*_DENIED_EXTRA`). Every tool takes `mode: headless|headed` (default `headless`).

## Use

1. Read: `read_range(path, sheet, cell_range)` / `preview_diff(path, ops)` — never writes.
2. First gated call returns `{status: needs_approval, ops_hash, diff}` — approve that exact diff.
3. Retry with `approval_token: <ops_hash>` → writes, returns `{ok, backup_id, thumbnail_b64}`.
4. `screenshot_sheet(full=true)` for full PNG, `rollback(path, backup_id)` (itself gated) to restore byte-identical.

Macros: `macro_run_gated` needs `macro_approved: true` + approval hash + `mode: headed`; only SHA256 logged, never macro text; headless refuses.

## Verify

```powershell
# cwd = repo root
excel-mcp\.venv\Scripts\python.exe -m pytest excel-mcp/tests/ -q -k "not needs_excel"   # headless CI, 107 passed
excel-mcp\.venv\Scripts\python.exe -m pytest excel-mcp/tests/ -q -m needs_excel          # Excel COM present, 3 passed; skipped otherwise
```

## Security notes

- Denylist `DENIED_PREFIXES` enforced in `server.call_tool` before backup/audit. Denied paths → `rejected`, no side effects.
- Audit lines to `audit.log`: `ts | tool | path | sheet:range | backup_id | ok/ERROR`. Errors never half-write; headed restores `.bak` on exception.
- No network listen (stdio only). `audit.log` + `*.versions/` git-ignored.

## Contracts

- Active: `spec.md`, `phases/phase_0.md` (locked, implemented 110/110).
- Archive: `docs/superpowers/specs/2026-10-13-excel-mcp-design.md`, plan `docs/superpowers/plans/2026-10-13-excel-mcp.md`.

## License

MIT — see `LICENSE`.
