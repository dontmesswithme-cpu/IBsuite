"""Task 2 tests: whitelist enforcement and the paranoid approval gate."""

import config
from core.guard import check_approval, is_path_allowed, make_ops_hash

OPS = {
    "tool": "write_range",
    "path": "a.xlsx",
    "sheet": "S",
    "range": "A1:A2",
    "values": [[1], [2]],
}


def test_traversal_rejected_no_side_effects():
    assert is_path_allowed("..\\..\\secret.xlsx") is False
    assert is_path_allowed("C:\\Windows\\foo.xlsx") is False


def test_rejected_paths_create_no_side_effects(tmp_path, monkeypatch):
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(config, "ALLOWED_ROOTS", [str(allowed)])

    assert is_path_allowed("..\\..\\secret.xlsx") is False
    assert is_path_allowed("C:\\Windows\\System32\\drivers\\etc\\hosts") is False
    assert sorted(p.name for p in tmp_path.iterdir()) == ["allowed"]


def test_documents_root_is_allowed():
    docs = next(r for r in config.ALLOWED_ROOTS if "Documents" in r)
    assert is_path_allowed(f"{docs}\\model.xlsx") is True


def test_project_root_is_allowed():
    root = next(r for r in config.ALLOWED_ROOTS if "Office plugin MCP" in r)
    assert is_path_allowed(f"{root}\\excel-mcp\\model.xlsx") is True


def test_absolute_path_outside_roots_rejected():
    assert is_path_allowed("C:\\Windows\\System32\\drivers\\etc\\hosts") is False


def test_empty_path_rejected():
    assert is_path_allowed("") is False
    assert is_path_allowed("   ") is False


def test_root_itself_is_not_a_target():
    assert is_path_allowed(config.ALLOWED_ROOTS[0]) is False


def test_approval_gate():
    denied = check_approval(OPS, None)
    assert denied["status"] == "needs_approval" and denied["allowed"] is False
    assert denied["ops_hash"] == make_ops_hash(OPS)

    ok = check_approval(OPS, make_ops_hash(OPS))
    assert ok["allowed"] is True


def test_stale_or_forged_token_rejected():
    stale = dict(OPS, values=[[9], [9]])
    forged = make_ops_hash(stale)

    denied = check_approval(OPS, forged)
    assert denied["allowed"] is False and denied["status"] == "needs_approval"


def test_ops_hash_is_deterministic_and_key_order_insensitive():
    first = make_ops_hash(OPS)
    reordered = {k: OPS[k] for k in reversed(list(OPS))}

    assert first == make_ops_hash(OPS) == make_ops_hash(reordered)
    assert isinstance(first, str) and len(first) == 64


def test_ops_hash_is_hex_sha256():
    int(make_ops_hash(OPS), 16)
