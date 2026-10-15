import os
from datetime import datetime, timezone
from pathlib import Path

_DEFAULT_LOG = Path(__file__).resolve().parent.parent / "audit.log"

AUDIT_LOG_PATH = Path(os.environ.get("PPT_MCP_AUDIT_LOG", _DEFAULT_LOG)).resolve()


def _timestamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def append(tool: str, path: str, detail: str, backup_id: str | None, status: str) -> None:
    line = " | ".join(
        [
            _timestamp(),
            str(tool),
            str(path),
            str(detail),
            str(backup_id) if backup_id else "-",
            str(status),
        ]
    )
    AUDIT_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(AUDIT_LOG_PATH, "a", encoding="utf-8") as handle:
        handle.write(line + "\n")
