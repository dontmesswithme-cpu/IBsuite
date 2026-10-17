import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path

MAX_VERSIONS = 20

_TS_FMT = "%Y%m%dT%H%M%S.%f"
_UNSAFE_CHARS = '<>:"/\\|?*'


def _versions_dir(path: str) -> Path:
    target = Path(path)
    return target.parent / (target.stem + ".versions")


def _safe_token(value: str) -> str:
    cleaned = "".join("_" if ch in _UNSAFE_CHARS else ch for ch in str(value))
    cleaned = cleaned.strip().strip(".")
    return cleaned or "unlabelled"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _next_backup_slot(version_dir: Path, tool: str, label: str) -> tuple[str, Path]:
    stamp = datetime.now(timezone.utc).strftime(_TS_FMT)
    stem = f"{stamp}_{_safe_token(tool)}_{_safe_token(label)}"
    candidate = version_dir / f"{stem}.bak"
    seq = 0
    while candidate.exists():
        seq += 1
        candidate = version_dir / f"{stem}_{seq}.bak"
    return candidate.name, candidate


def _write_backup(source: Path, version_dir: Path, tool: str, label: str) -> str:
    version_dir.mkdir(parents=True, exist_ok=True)
    backup_id, destination = _next_backup_slot(version_dir, tool, label)
    shutil.copy2(source, destination)
    meta = {
        "backup_id": backup_id,
        "tool": tool,
        "label": label,
        "path": str(source),
        "created_at": datetime.now(timezone.utc).isoformat(),
        "size": destination.stat().st_size,
        "sha256": _sha256(destination),
    }
    (version_dir / (backup_id[:-4] + ".json")).write_text(
        json.dumps(meta, indent=2), encoding="utf-8"
    )
    return backup_id


def _enforce_retention(version_dir: Path) -> None:
    baks = sorted(version_dir.glob("*.bak"), key=lambda p: p.name)
    for stale in baks[:-MAX_VERSIONS]:
        stale.unlink(missing_ok=True)
        stale.with_suffix(".json").unlink(missing_ok=True)


def backup(path: str, tool: str, label: str) -> str:
    source = Path(path)
    if not source.is_file():
        raise FileNotFoundError(f"cannot back up missing file: {source}")
    version_dir = _versions_dir(path)
    backup_id = _write_backup(source, version_dir, tool, label)
    _enforce_retention(version_dir)
    return backup_id


def list_versions(path: str) -> list[str]:
    version_dir = _versions_dir(path)
    if not version_dir.is_dir():
        return []
    return sorted((entry.name for entry in version_dir.glob("*.bak")))


def rollback(path: str, backup_id: str) -> dict:
    source = Path(path)
    version_dir = _versions_dir(path)
    restore_from = version_dir / Path(backup_id).name

    if not restore_from.is_file():
        return {"ok": False, "reason": f"unknown backup_id: {backup_id}"}

    payload = restore_from.read_bytes()

    pre_rollback_id = None
    if source.is_file():
        pre_rollback_id = _write_backup(source, version_dir, "rollback", "pre_rollback")

    source.parent.mkdir(parents=True, exist_ok=True)
    source.write_bytes(payload)

    _enforce_retention(version_dir)

    return {
        "ok": True,
        "path": str(source),
        "backup_id": backup_id,
        "pre_rollback_id": pre_rollback_id,
    }
