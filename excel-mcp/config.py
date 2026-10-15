import json
import os
from pathlib import Path

from styles.ib import get_style

BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent

CUSTOM_STYLES_PATH = BASE_DIR / "styles" / "custom.json"


def _default_denied_prefixes() -> list[str]:
    denied: list[str] = []
    windir = os.environ.get("SystemRoot") or os.environ.get("windir") or r"C:\Windows"
    denied.append(windir)
    for var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramData"):
        val = os.environ.get(var)
        if val and val.strip():
            denied.append(val)
    # Other users' profiles (siblings of this user's home under C:\Users).
    try:
        home = Path.home()
        if home.parent.name.lower() == "users":
            for child in home.parent.iterdir():
                if child.is_dir() and child != home:
                    denied.append(str(child))
    except OSError:
        pass
    # Protocol files: the sentinel runtime is never a workbook target.
    denied.append(str(PROJECT_ROOT / "sentinel"))
    # Operator escape hatch: extra semicolon-separated prefixes to deny.
    extra = os.environ.get("EXCEL_MCP_DENIED_EXTRA", "")
    denied.extend(part for part in extra.split(";") if part.strip())
    return [str(Path(part.strip()).expanduser()) for part in denied if part.strip()]


DENIED_PREFIXES: list[str] = _default_denied_prefixes()


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
