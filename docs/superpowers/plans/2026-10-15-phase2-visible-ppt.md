# Phase 2 Visible-Headed + PowerPoint Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Headed Excel runs on-screen and a dual-mode PowerPoint server ships beside it with the same paranoid flow.

**Architecture:** P2.1 first (one-line visible flip + asserts, excel suite decides), then P2.2 (`ppt-mcp/` scaffold → guard/versions/audit mirrors → headless → headed/engine/shots → server wiring → deck fixture, each with its own tests). Ordered gate block: excel suite green, then ppt suite green.

**Tech Stack:** Python >=3.11, `python-pptx` (pinned, new), `pywin32==312` (`PowerPoint.Application`), `mcp==2.3.0`, `pillow==12.3.0`, `pytest>=8`.

**Spec:** `spec.md` §8 (R6–R9, N5–N7, D19–D22) + `phases/phase_2.md` (AC-1–AC-4). Requires Phase 1 approved + green. Executors read spec + phase contract + this plan.

## Global Constraints

- Python `>=3.11`. Excel pins unchanged; PPT adds `python-pptx` pinned in `ppt-mcp/pyproject.toml`.
- New code ONLY in `excel-mcp/` (P2.1 flip + asserts) and `ppt-mcp/` (P2.2). Do NOT touch `sentinel/`.
- `ppt-mcp/` mirrors Excel shapes; no cross-imports between servers; each keeps its own venv and suite (N7).
- `ppt-mcp/core/` transport-agnostic (no MCP imports); `server.py` thin stdio only.
- Default engine `mode="headless"` on every tool; headed always `Visible=True` (D19); `needs_excel`/`needs_powerpoint` tests skip (never error) where COM is absent.
- Same paranoid flow: whitelist → exact-`ops_hash` approval → copy-before-write backup (20 FIFO, git-ignored) → dispatch → audit append → `{ok, backup_id, thumbnail}`.
- `audit.log` and `*.versions/` git-ignored. Sentinel runtime: DS `npm run check` per save, OP watches per `sentinel/howtowork.md`.

## Review Focus

- Visible flip breaks headless CI (Excel window attempted without COM) — expect headless suite fully green, headed tests skip cleanly. Pinned in Task 1 gate run.
- PPT whitelist narrower/wider than Excel's trio — expect identical default (repo root + Documents + Downloads). Pinned in Task 2 guard test.
- Chart call succeeds headless but file won't open in PowerPoint — expect roundtrip open + headed save clean. Pinned in Task 3 chart test.
- Headed PowerPoint disconnect (`0x80010108` class, seen on Excel 2026-10-15) crashes instead of fallback — expect retry-then-`headed_unavailable_fell_back`. Pinned in Task 4 stress test.
- Demo-shaped files or `audit.log`/`.versions` committed — expect `git status` clean apart from intended files. Pinned in Task 5.

---

### Task 1: P2.1 visible-headed flip (Excel)

**Files:**
- Modify: `excel-mcp/core/headed.py` (one line)
- Modify: `excel-mcp/tests/test_headed.py`
- Test: `excel-mcp/tests/test_headed.py::test_headed_visible_default`

**Interfaces:**
- Consumes: `core/headed._excel()` context manager
- Produces: `_excel()` yields app with `Visible is True`, `DisplayAlerts is False` (no signature change)

- [ ] **Step 1: Write the failing test**

```python
def test_headed_visible_default(monkeypatch):
    seen = {}
    monkeypatch.setattr(headed, "_launch", lambda: FakeApp(seen))
    with headed._excel():
        pass
    assert seen["Visible"] is True and seen["DisplayAlerts"] is False
```

- [ ] **Step 2: Run test to verify it fails**

Run: `excel-mcp\.venv\Scripts\python.exe -m pytest excel-mcp/tests/test_headed.py::test_headed_visible_default -v` (cwd repo root)
Expected: FAIL (asserts not met: `Visible` is `False`)

- [ ] **Step 3: Flip `Visible=False` to `Visible=True` in `_excel()` in `excel-mcp/core/headed.py`**

One line. Nothing else in the file changes.

- [ ] **Step 4: Run excel suite to verify AC-1**

Run: `excel-mcp\.venv\Scripts\python.exe -m pytest excel-mcp/tests/ -q` (cwd repo root)
Expected: exit 0; headless pass; `needs_excel` pass here / skip elsewhere, never error

- [ ] **Step 5: Commit**

```bash
git add excel-mcp/core/headed.py excel-mcp/tests/test_headed.py
git commit -m "feat: headed Excel runs visible"
```

### Task 2: ppt-mcp scaffold + trust primitives

**Files:**
- Create: `ppt-mcp/pyproject.toml`, `ppt-mcp/.gitignore`, `ppt-mcp/config.py`, `ppt-mcp/opencode_mcp.json.example`
- Create: `ppt-mcp/styles/ib_deck.py`, `ppt-mcp/styles/custom.json`
- Create: `ppt-mcp/core/guard.py`, `ppt-mcp/core/versions.py`, `ppt-mcp/core/audit.py`
- Test: `ppt-mcp/tests/test_deck_styles.py`, `ppt-mcp/tests/test_ppt_guard.py`, `ppt-mcp/tests/test_ppt_versions.py`

**Interfaces:**
- Consumes: nothing (first PPT task)
- Produces:
  - `styles/ib_deck.py: IB_DECK_DEFAULTS: dict`, `get_deck_style(key: str) -> dict | str`
  - `config.py: ALLOWED_ROOTS: list[str]` (repo root + Documents + Downloads)
  - `core/guard.py: is_path_allowed(path: str) -> bool`, `make_ops_hash(ops: dict) -> str`, `check_approval(ops, token) -> dict`
  - `core/versions.py: backup(path, tool, label) -> str`, `list_versions(path) -> list[str]`, `rollback(path, backup_id) -> dict`
  - `core/audit.py: append(tool, path, detail, backup_id, status) -> None`

- [ ] **Step 1: Write the failing tests**

```python
def test_deck_defaults():
    assert get_deck_style("header_fill") == "002060"
    assert any("IBsuite" in r or "Office plugin MCP" in r for r in ALLOWED_ROOTS)

def test_ppt_approval_gate():
    denied = check_approval({"tool": "add_slide", "path": "d.pptx"}, None)
    assert denied["status"] == "needs_approval" and denied["allowed"] is False

def test_ppt_backup_fifo(tmp_path):
    f = tmp_path / "d.pptx"; write_minimal_pptx(str(f))
    for i in range(21):
        backup(str(f), "t", f"r{i}")
    assert len(list_versions(str(f))) == 20
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `ppt-mcp\.venv\Scripts\python.exe -m pytest ppt-mcp/tests/test_deck_styles.py ppt-mcp/tests/test_ppt_guard.py ppt-mcp/tests/test_ppt_versions.py -v` (cwd repo root; create venv via `python -m venv ppt-mcp/.venv` + `pip install -e .[test]` first)
Expected: FAIL with "not defined / file not found"

- [ ] **Step 3: Implement scaffold + styles + guard/versions/audit mirroring the Excel shapes**

`IB_DECK_DEFAULTS` holds navy/white/body/border deck values; `custom.json` starts `{}`; `check_approval`: token `== make_ops_hash(ops)` → allowed; sidecar `<Stem>.versions/`, 20 FIFO.

- [ ] **Step 4: Run tests to verify they pass**

Run: same pytest invocation as Step 2
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add ppt-mcp/pyproject.toml ppt-mcp/.gitignore ppt-mcp/config.py ppt-mcp/opencode_mcp.json.example ppt-mcp/styles/ib_deck.py ppt-mcp/styles/custom.json ppt-mcp/core/guard.py ppt-mcp/core/versions.py ppt-mcp/core/audit.py ppt-mcp/tests/test_deck_styles.py ppt-mcp/tests/test_ppt_guard.py ppt-mcp/tests/test_ppt_versions.py
git commit -m "feat: scaffold ppt-mcp with deck styles and trust primitives"
```

### Task 3: PPT headless core (slides, text, tables, charts, images, theme)

**Files:**
- Create: `ppt-mcp/core/headless.py`, `ppt-mcp/core/format_headless.py`
- Test: `ppt-mcp/tests/test_ppt_headless.py`, `ppt-mcp/tests/test_deck_format.py`

**Interfaces:**
- Consumes: Task 2 guard/versions/audit (caller orchestrates), `styles/ib_deck`
- Produces:
  - `list_slides(path: str) -> list[dict]`, `read_text(path: str, slide: int) -> dict`
  - `add_slide(path: str, layout: str, index: int | None) -> dict`, `set_text(path: str, slide: int, shape: str, text: str) -> dict`
  - `add_table(path, slide, rows, cols, values) -> dict`, `add_chart(path, slide, chart_type, data, anchor) -> dict`, `add_image(path, slide, image_path, anchor) -> dict`
  - `apply_theme(path: str) -> dict`, `preview_diff(path: str, ops: dict) -> dict` (never writes)

- [ ] **Step 1: Write the failing tests**

```python
def test_slide_text_table_roundtrip(tmp_path):
    f = tmp_path / "d.pptx"; make_deck(str(f), slides=2)
    add_slide(str(f), "blank", None)
    set_text(str(f), 3, "title", "Comps")
    assert read_text(str(f), 3)["title"] == "Comps"

def test_chart_opens_in_powerpoint_shape(tmp_path):
    f = tmp_path / "d.pptx"; make_deck(str(f), slides=1)
    add_chart(str(f), 1, "column", {"cats": ["A"], "series": {"Rev": [1]}}, "anchor": "D5")
    assert has_chart(str(f), 1) is True

def test_theme_navy_golden(tmp_path):
    f = tmp_path / "d.pptx"; make_deck(str(f), slides=1)
    apply_theme(str(f))
    assert theme_heading_fill(str(f), 1) == "002060"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `ppt-mcp\.venv\Scripts\python.exe -m pytest ppt-mcp/tests/test_ppt_headless.py ppt-mcp/tests/test_deck_format.py -v` (cwd repo root)
Expected: FAIL with "not defined"

- [ ] **Step 3: Implement headless with `python-pptx` + theme via `resolve` overrides**

Unknown slide → `ValueError("unknown slide ...")`; bad shape → `ValueError`; `preview_diff` reads current state and returns computed after-image without saving.

- [ ] **Step 4: Run tests to verify they pass**

Run: same pytest invocation as Step 2
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add ppt-mcp/core/headless.py ppt-mcp/core/format_headless.py ppt-mcp/tests/test_ppt_headless.py ppt-mcp/tests/test_deck_format.py
git commit -m "feat: add ppt headless slides text tables charts images theme"
```

### Task 4: PPT headed + engine + shots + server wiring

**Files:**
- Create: `ppt-mcp/core/headed.py`, `ppt-mcp/core/engine.py`, `ppt-mcp/core/shots.py`, `ppt-mcp/server.py`
- Test: `ppt-mcp/tests/test_ppt_headed.py`, `ppt-mcp/tests/test_ppt_engine.py`, `ppt-mcp/tests/test_ppt_shots.py`, `ppt-mcp/tests/test_ppt_server.py`

**Interfaces:**
- Consumes: Task 2 + 3 (headless fns, guard, versions, audit)
- Produces:
  - `headed.py: is_powerpoint_available() -> bool`, `headed_apply(path: str, ops: dict) -> dict`
  - `engine.py: dispatch(tool: str, params: dict) -> dict` (visible headed, fallback `{warning: "headed_unavailable_fell_back"}`)
  - `shots.py: thumbnail(...) -> str`, `full_shot(..., mode) -> str`
  - `server.py` stdio with 4 reads + 8 gated (same paranoid flow as Excel)

- [ ] **Step 1: Write the failing tests**

```python
@pytest.mark.needs_powerpoint
def test_headed_unavailable_falls_back(tmp_path):
    f = tmp_path / "d.pptx"; make_deck(str(f), slides=1)
    out = dispatch("read_text", {"path": str(f), "slide": 1, "mode": "headed"})
    assert "title" in out or out.get("warning") == "headed_unavailable_fell_back"

def test_write_needs_approval_then_succeeds(tmp_path):
    f = tmp_path / "d.pptx"; make_deck(str(f), slides=1)
    first = call_tool("set_text", {"path": str(f), "slide": 1, "shape": "title", "text": "X"})
    assert first["status"] == "needs_approval"
    second = call_tool("set_text", {"path": str(f), "slide": 1, "shape": "title", "text": "X", "approval_token": first["ops_hash"]})
    assert second["ok"] is True and second["backup_id"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `ppt-mcp\.venv\Scripts\python.exe -m pytest ppt-mcp/tests/test_ppt_headed.py ppt-mcp/tests/test_ppt_engine.py ppt-mcp/tests/test_ppt_shots.py ppt-mcp/tests/test_ppt_server.py -v` (cwd repo root)
Expected: FAIL with "not defined"

- [ ] **Step 3: Implement headed (visible COM, retry 200/500/1000ms, `.bak` restore) + dispatcher + shots + thin server**

`headed.py` does no guard/backup itself — `server.py` does. Detached → `HeadedUnavailable` caught by `dispatch` for fallback.

- [ ] **Step 4: Run tests to verify they pass**

Run: `ppt-mcp\.venv\Scripts\python.exe -m pytest ppt-mcp/tests -q -k "not needs_powerpoint"` then `-m needs_powerpoint` (cwd repo root)
Expected: non-COM PASS; COM-marked SKIP or PASS, never ERROR

- [ ] **Step 5: Commit**

```bash
git add ppt-mcp/core/headed.py ppt-mcp/core/engine.py ppt-mcp/core/shots.py ppt-mcp/server.py ppt-mcp/tests/test_ppt_headed.py ppt-mcp/tests/test_ppt_engine.py ppt-mcp/tests/test_ppt_shots.py ppt-mcp/tests/test_ppt_server.py
git commit -m "feat: add ppt headed engine shots and server wiring"
```

### Task 5: Golden deck fixture + ordered gate block green

**Files:**
- Create: `ppt-mcp/tests/fixtures/make_deck.py`, `ppt-mcp/tests/fixtures/ib_deck.pptx`, `ppt-mcp/tests/golden_deck.json`, `ppt-mcp/tests/test_deck.py`, `ppt-mcp/tests/test_stress_ppt.py`

**Interfaces:**
- Consumes: Tasks 2–4
- Produces: `make_deck.build(path: str) -> None`, `test_full_ib_grade_vs_golden`, headed stress tests, AC-1–AC-4 evidence

- [ ] **Step 1: Write the failing tests**

```python
def test_full_ib_grade_vs_golden():
    assert deck_theme_vs_golden(FIXTURE_DECK, GOLDEN_DECK) is True

@pytest.mark.needs_powerpoint
def test_headed_deck_roundtrip(tmp_path):
    out = dispatch("read_text", {"path": str(tmp_path / "d.pptx"), "slide": 1, "mode": "headed"})
    assert "title" in out or out.get("warning") == "headed_unavailable_fell_back"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `ppt-mcp\.venv\Scripts\python.exe -m pytest ppt-mcp/tests/test_deck.py ppt-mcp/tests/test_stress_ppt.py -v` (cwd repo root)
Expected: FAIL (fixture/golden not present)

- [ ] **Step 3: Implement builder (title, 3-statement summary, comps table, chart slide, 16:9), emit binary, freeze `golden_deck.json`, wire asserts**

- [ ] **Step 4: Run the ordered gate block**

Run AC-1: `excel-mcp\.venv\Scripts\python.exe -m pytest excel-mcp/tests/ -q` then AC-2: `ppt-mcp\.venv\Scripts\python.exe -m pytest ppt-mcp/tests -q` (cwd repo root)
Expected: both exit 0; COM-marked pass here / skip elsewhere, never error

- [ ] **Step 5: Commit**

```bash
git add ppt-mcp/tests/fixtures/make_deck.py ppt-mcp/tests/fixtures/ib_deck.pptx ppt-mcp/tests/golden_deck.json ppt-mcp/tests/test_deck.py ppt-mcp/tests/test_stress_ppt.py
git commit -m "feat: add golden deck fixture and gate evidence"
```
