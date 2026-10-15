import hashlib
import json
from pathlib import Path

import config

_SUMMARY_KEYS = ("tool", "path", "sheet", "range", "values", "macro")


def make_ops_hash(ops: dict) -> str:
    canonical = json.dumps(
        ops,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=True,
        default=str,
    )
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def _is_denied(resolved: Path) -> bool:
    for denied in config.DENIED_PREFIXES:
        if resolved.is_relative_to(Path(denied)):
            return True
    return resolved.parent == resolved  # a drive root itself is never a target


def is_path_allowed(path: str) -> bool:
    if not isinstance(path, str) or not path.strip():
        return False

    raw = Path(path.strip())
    resolved = raw.resolve() if raw.is_absolute() else (Path.cwd() / raw).resolve()
    if resolved.is_dir():
        return False
    return not _is_denied(resolved)


def check_approval(ops: dict, approval_token: str | None) -> dict:
    ops_hash = make_ops_hash(ops)

    if approval_token is not None and approval_token == ops_hash:
        return {"allowed": True, "status": "approved", "ops_hash": ops_hash}

    return {
        "allowed": False,
        "status": "needs_approval",
        "ops_hash": ops_hash,
        "diff": {key: ops[key] for key in _SUMMARY_KEYS if key in ops},
    }
