import json
import os
from pathlib import Path

from styles.ib import get_style

BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent

CUSTOM_STYLES_PATH = BASE_DIR / "styles" / "custom.json"


def _resolve_allowed_roots() -> list[str]:
    env_roots = os.environ.get("EXCEL_MCP_ROOTS", "")
    if env_roots.strip():
        return [str(Path(part).expanduser()) for part in env_roots.split(";") if part.strip()]
    return [str(PROJECT_ROOT), str(Path.home() / "Documents")]


ALLOWED_ROOTS: list[str] = _resolve_allowed_roots()


def load_custom_overrides() -> dict:
    try:
        raw = CUSTOM_STYLES_PATH.read_text(encoding="utf-8-sig")
    except OSError:
        return {}
    try:
        data = json.loads(raw)
    except ValueError:
        return {}
    return data if isinstance(data, dict) else {}


def resolve_style(key: str, override: dict | None) -> dict | str:
    if override is not None and key in override:
        return override[key]
    custom = load_custom_overrides()
    if key in custom:
        return custom[key]
    return get_style(key)
