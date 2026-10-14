# Excel MCP — Design Spec (SN Genesis Record)
Date: 2026-10-13
Status: Draft for user review — no implementation yet
Path: `excel-mcp/` beside `sentinel/` in `C:\1. (A) Faraz\Office plugin MCP`

## 1. Intent & Success Criteria

**Outcome:** A local MCP server (`excel-mcp`) that lets the assistant create and edit real Excel files from chat prompts — both headless (no Excel open) and headed (driving Excel Desktop).

**Who:** Single user (Faraz), Windows + Excel Desktop, prompting from Opencode here.

**Success:**
- Prompt "build me a 3-statement model with ..." → valid `.xlsx` on disk, formulas compute, opens clean in Excel.
- IB-grade formatting by default (no neon, no misaligned decimals).
- Paranoid-first trust: nothing writes without explicit approval + backup + audit entry.
- "show me" → screenshot of the sheet. "roll back to v3" → byte-identical restore.
- Big/complex models work (multi-sheet, charts, pivots) without half-writes or COM hangs.

**Non-goals (V1):**
- No Excel Online / Mac / Office.js add-in.
- No remote HTTP exposure (architected pluggable, not shipped).
- No unrelated refactor of `sentinel/`.

## 2. Locked Decisions

| # | Decision | Choice |
|---|----------|--------|
| D1 | Headless meaning | File-level `.xlsx` read/write via lib, no Excel install needed |
| D2 | Head meaning | COM automation of Excel Desktop on same Windows box (`Excel.Application`) |
| D3 | Operations scope | Full power: values, formulas, sheets, formatting, tables, validation, conditional formatting, charts, pivots, images, print setup |
| D4 | Quality bar | IB-grade: Arial/Calibri 9-10, navy/blue headers, accounting formats, thin grid borders, frozen panes, print fit-to-page |
| D5 | Styles customisation | `styles/ib.py` defaults + `styles/custom.json` override, per-call `style_override`, future `style_save_as_preset(name)` |
| D6 | Transport | Local stdio V1 (Opencode / Claude Desktop). `server.py` thin wrapper, `core/` transport-agnostic for later HTTP/SSE |
| D7 | Stack | Python 3.11+, `mcp` SDK, `openpyxl`, `pywin32`, `pillow` |
| D8 | Architecture | Option 1: single Python MCP, dual-engine core (rejected split-server Option 2 as overkill for local-only V1) |
| D9 | Guardrails | Paranoid-first: read-only default, `write/format/macro` returns `needs_approval` until exact diff approved in-turn. Whitelist folders only |
| D10 | Macros/VBA | (B) Read VBA freely, execute only with explicit per-action approval + audit log (text hash logged). Never auto-run on open |
| D11 | Review loop | Diff summary always; auto-thumbnail every change; full PNG only on "show me"; version history with rollback (mixed B+C) |
| D12 | Backups | Sidecar `.versions/` per workbook, copy-before-write, keep 20 FIFO, `.bak` + `.json` diff sidecar, git-ignored |

## 3. Architecture

```
excel-mcp/
  server.py              # stdio wrapper, tool schemas, no Excel logic
  core/
    engine.py            # Engine(mode) interface, dispatches headless/headed
    headless.py          # openpyxl implementation
    headed.py            # pywin32 COM, single-thread marshalled
    guard.py             # approval gate + whitelist
    versions.py          # backup + rollback
    audit.py             # append-only audit.log
    shots.py             # PNG render (headless: openpyxl+pillow; headed: CopyPicture)
  styles/
    ib.py                # IB defaults
    custom.json          # user overrides (optional, git-ignored if personal)
  tests/
    test_headless.py
    test_guard.py
    test_versions.py
    test_styles.py
    test_headed.py       # @needs_excel, skipped in CI
  audit.log              # runtime, git-ignored
  *.versions/            # runtime, git-ignored
```

`sentinel/` untouched. New code lives in `excel-mcp/`.

## 4. Tools (MCP surface, 13)

Always-allowed (read-only, 4):
- `list_sheets(path) -> [names, dims, frozen, print]`
- `read_range(path, sheet, range) -> values + formulas + number-formats`
- `preview_diff(path, ops) -> before/after summary, no write`
- `screenshot_sheet(path, sheet, range?) -> thumbnail always, full PNG on demand`

Gated (require approval token for exact ops hash, 9):
- `write_range`, `set_formula`, `add_sheet`, `format_range`, `add_chart`, `add_pivot`, `add_image`, `macro_run_gated`, `rollback`
- Each gated call without valid approval returns `{status:"needs_approval", diff}` — never writes.
- `rollback` itself needs approval and backups the pre-rollback state first.

Engine param: `mode: "headless" | "headed"` per call, default `"headless"`.

## 5. Data Flow (happy path)

1. User: "fix EBITDA margin + make it IB-grade"
2. Assistant calls `preview_diff` → returns value/format delta, no FS touch.
3. User approves (explicit "yes do it" for that diff).
4. Server: `guard.check(approval)` → `versions.backup()` (SaveCopyAs / file copy) → `engine.apply()` batched by range → `audit.append()` → return `{ok, backup_id, diff_applied, thumbnail_b64}`.
5. User: "show me" → `screenshot_sheet(full=True)` → PNG.
6. User: "roll back to v3" → `rollback` → live file replaced, `audit.log` entry, thumbnail of restored state.

Batching: range writes coalesced, max 100k cells per call (reject with "split" guidance above).

## 6. IB Styles (`styles/ib.py`)

- Header fill `002060` (navy), font white bold 10; body `Arial 10`; inputs light blue `D9E1F2`, hardcoded greys, formulas black.
- Numbers: `_-* #,##0_-;-* #,##0_-;-* "-"_-;_-@_-` for integers, `_-* #,##0.0_-...` for 1dp, `%_0.0%` for margins/multiples.
- Borders: thin `D9D9D9` grid, medium bottom on subtotals, double on totals.
- Layout: freeze top row + first col on schedules, widths 14-18, row height 15, `FitToPagesWide=1, FitToPagesTall=0`, landscape, footer `&F | &A`.
- `custom.json` keys mirror these; missing keys fall back to defaults. No code edit to re-theme.

## 7. Backups, Audit, Rollback

- Sidecar: `<Name>.versions/<ts>_<tool>_<range>.bak` + `.json` (ops, approval id, user ts).
- Copy-before-write enforced in `versions.backup()`; headed path does `Workbook.SaveCopyAs(bak)` before mutation, then `Save`.
- Retention 20 FIFO per workbook. `audit.log` line: `ts | tool | path | sheet:range | backup_id | macro_sha256? | result`.
- Rollback verifies byte-restore + re-reads dims as smoke check.
- Both `audit.log` and `*.versions/` in `.gitignore`.

## 8. Error Handling

- File locked / Excel busy (`RPC_E_CALL_REJECTED`, `PermissionError`): retry 3x exp backoff (200/500/1000ms), then `{status:"needs_retry", hint}` — no partial write (ops applied atomically per backup; on exception restore `.bak` automatically).
- COM detached: re-`GetActiveObject`/`Dispatch`, else fallback to headless with `warning:"headed_unavailable_fell_back"`.
- Oversize / bad range / unknown sheet: `{status:"rejected", reason}` with fix hint.
- Macro failure: log stderr, leave file at pre-macro `.bak` state, return error + backup id.
- All errors append to `audit.log` with `ERROR` flag.

## 9. Testing

- `pytest`: headless round-trip (write→reopen→values+formats equal), golden-file style assert vs `tests/golden_ib.json`, guard tests (write without approval blocked; macro without `macro_approved:true` blocked), versions tests (backup exists, rollback byte-identical, FIFO eviction at 21), shots tests (thumbnail <40KB, full PNG > thumbnail).
- `test_headed.py` gated `@pytest.mark.needs_excel`, skipped if `win32com` or `Excel.Application` unavailable. CI runs headless-only.
- Manual IB check: open golden file in Excel, visual pass on fonts/colors/print preview + screenshot diff.

## 10. Security Notes

- Whitelist enforcement in `guard.py`: `ALLOWED_ROOTS` env/config, `..` traversal rejected, absolute paths must prefix-match.
- Macro text never logged verbatim (SHA256 + first 200 chars redacted) to avoid secret leak; full text only in ephemeral approval request, not persisted.
- No network listen in V1 (stdio only). Future HTTP wrapper must add token auth + TLS before ship — out of scope now.

## 11. Open Questions for Implementation Plan

- `ALLOWED_ROOTS` default: just `C:\1. (A) Faraz\Office plugin MCP` + Documents, or whole drive with per-write confirm?
- Golden IB sample: provide one real (sanitised) model to lock styles against?
- Headed default workbook behaviour: attach to active workbook vs always open by path?

---
[SN] Genesis spec written from brainstorming 2026-10-13. Awaiting DIR review before writing-plans.
