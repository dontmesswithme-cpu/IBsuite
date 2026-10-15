"""Guard tests: denylist path enforcement and the paranoid approval gate."""

import os
from pathlib import Path

import pytest

import config
from core.guard import check_approval, is_path_allowed, make_ops_hash

OPS = {
    "tool": "write_range",
    "path": "a.xlsx",
    "sheet": "S",
    "range": "A1:A2",
    "values": [[1], [2]],
}


def test_system_dirs_denied():
    assert is_path_allowed("C:\\Windows\\foo.xlsx") is False
    assert is_path_allowed("C:\\Windows\\System32\\drivers\\etc\\hosts") is False


def test_denied_prefixes_cover_system_and_sentinel():
    lowered = [d.lower() for d in config.DENIED_PREFIXES]
    assert any("windows" in d for d in lowered)
    assert any("sentinel" in d for d in lowered)


def test_program_files_denied():
    program_files = os.environ.get("ProgramFiles")
    if not program_files:
        pytest.skip("ProgramFiles env var not set on this box")
    assert is_path_allowed(str(Path(program_files) / "app.xlsx")) is False


def test_everyday_paths_allowed(tmp_path):
    assert is_path_allowed(str(tmp_path / "m.xlsx")) is True
    assert is_path_allowed("relative_model.xlsx") is True


def test_dotdot_resolves_before_decision(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    assert is_path_allowed("..\\outside_but_safe.xlsx") is True
    assert is_path_allowed("C:\\Windows\\..\\Windows\\foo.xlsx") is False


def test_denied_paths_create_no_side_effects(tmp_path):
    before = sorted(p.name for p in tmp_path.iterdir())
    assert is_path_allowed("C:\\Windows\\System32\\drivers\\etc\\hosts") is False
    assert sorted(p.name for p in tmp_path.iterdir()) == before


def test_sentinel_tree_denied():
    sentinel_file = Path(config.PROJECT_ROOT) / "sentinel" / "inbox_ds.md"
    assert is_path_allowed(str(sentinel_file)) is False


def test_drive_root_denied():
    assert is_path_allowed("C:\\") is False


def test_empty_path_rejected():
    assert is_path_allowed("") is False
    assert is_path_allowed("   ") is False


def test_existing_directory_is_not_a_target(tmp_path):
    assert is_path_allowed(str(tmp_path)) is False


def test_denied_prefix_override(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "DENIED_PREFIXES", [str(tmp_path)])
    assert is_path_allowed(str(tmp_path / "m.xlsx")) is False


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
