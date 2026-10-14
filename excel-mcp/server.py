import hashlib
from pathlib import Path

from mcp.server.mcpserver import MCPServer

from core import audit, engine, guard, versions
from core.guard import check_approval, make_ops_hash
from core.shots import full_shot, thumbnail

mcp = MCPServer(name="excel-mcp")

READ_TOOLS = ("list_sheets", "read_range", "preview_diff", "screenshot_sheet")
GATED_TOOLS = (
    "write_range",
    "set_formula",
    "add_sheet",
    "format_range",
    "add_chart",
    "add_pivot",
    "add_image",
    "rollback",
)

_HASH_EXCLUDE = ("approval_token", "mode")
MACRO_LOG: list[dict] = []


def _canonical(tool: str, params: dict) -> dict:
    return {"tool": tool, **{k: params[k] for k in params if k not in _HASH_EXCLUDE}}


def ops_hash(tool: str, params: dict) -> str:
    return make_ops_hash(_canonical(tool, params))


def executed_flag() -> bool:
    return bool(MACRO_LOG)


def _sanitize(text: str) -> str:
    return str(text).replace("|", "/").replace("\n", " ")


def _backup_label(tool: str, params: dict) -> str:
    if tool in ("write_range", "format_range"):
        return str(params.get("cell_range", "op"))
    if tool == "set_formula":
        return str(params.get("cell", "op"))
    if tool == "add_sheet":
        return str(params.get("name", "op"))
    if tool == "add_chart":
        return str(params.get("anchor", "op"))
    if tool == "add_pivot":
        return str(params.get("dest", "op"))
    if tool == "add_image":
        return str(params.get("anchor", "op"))
    return "op"


def _audit_detail(tool: str, params: dict) -> str:
    sheet = params.get("sheet", "")
    if tool == "write_range":
        return _sanitize(f"{sheet}:{params.get('cell_range', '')}")
    if tool == "set_formula":
        return _sanitize(f"{sheet}:{params.get('cell', '')}")
    if tool == "add_sheet":
        return _sanitize(str(params.get("name", "")))
    if tool == "format_range":
        return _sanitize(f"{sheet}:{params.get('cell_range', '')}")
    if tool == "add_chart":
        return _sanitize(f"{sheet}:{params.get('data_range', '')}@{params.get('anchor', '')}")
    if tool == "add_pivot":
        return _sanitize(f"{sheet}:{params.get('source_range', '')}->{params.get('dest', '')}")
    if tool == "add_image":
        image = Path(str(params.get("image_path", ""))).name
        return _sanitize(f"{sheet}:{params.get('anchor', '')}:{image}")
    if tool == "rollback":
        return _sanitize(f"rollback:{params.get('backup_id', '')}")
    return _sanitize(tool)


def _best_thumb(params: dict):
    try:
        return thumbnail(
            params["path"], params.get("sheet", "Sheet1"), params.get("cell_range")
        )
    except Exception:
        return None


def _screenshot(params: dict) -> dict:
    path = params["path"]
    sheet = params.get("sheet", "Sheet1")
    cell_range = params.get("cell_range")
    full = bool(params.get("full", False))
    mode = params.get("mode", "headless")
    out = {"thumbnail_b64": thumbnail(path, sheet, cell_range), "full_b64": None}
    if full:
        out["full_b64"] = full_shot(path, sheet, cell_range, mode)
    return out


def _call_read(tool: str, params: dict) -> dict:
    try:
        if tool == "screenshot_sheet":
            data = _screenshot(params)
        else:
            data = engine.dispatch(tool, params)
    except (ValueError, FileNotFoundError) as exc:
        return {"ok": False, "status": "error", "reason": str(exc)}
    if isinstance(data, dict):
        return {"ok": True, "status": "ok", **data}
    return {"ok": True, "status": "ok", "result": data}


def _call_gated(tool: str, params: dict) -> dict:
    check = check_approval(_canonical(tool, params), params.get("approval_token"))
    if not check["allowed"]:
        return {
            "ok": False,
            "status": "needs_approval",
            "ops_hash": check["ops_hash"],
            "diff": check["diff"],
        }

    path = params["path"]
    backup_id = None
    if tool != "rollback":
        try:
            backup_id = versions.backup(path, tool, _backup_label(tool, params))
        except (ValueError, FileNotFoundError, OSError) as exc:
            return {"ok": False, "status": "error", "reason": str(exc)}

    try:
        result = engine.dispatch(tool, params)
    except Exception as exc:
        audit.append(tool, path, _audit_detail(tool, params), backup_id, "ERROR")
        return {"ok": False, "status": "error", "reason": str(exc), "backup_id": backup_id}

    if isinstance(result, dict) and result.get("ok") is False:
        audit.append(tool, path, _audit_detail(tool, params), backup_id, "ERROR")
        merged = {"ok": False, "status": "error", "backup_id": backup_id}
        merged.update({k: v for k, v in result.items() if k != "ok"})
        if "reason" not in merged:
            merged["reason"] = "operation failed"
        return merged

    if tool == "rollback" and isinstance(result, dict):
        backup_id = result.get("pre_rollback_id")

    audit.append(tool, path, _audit_detail(tool, params), backup_id, "ok")
    out: dict = {"ok": True, "status": "ok", "backup_id": backup_id, "thumbnail_b64": _best_thumb(params)}
    if isinstance(result, dict):
        for key, value in result.items():
            out.setdefault(key, value)
    return out


def _macro_sha(macro: str) -> str:
    return hashlib.sha256(str(macro).encode("utf-8")).hexdigest()


def _run_macro_com(path: str, macro_name: str) -> dict:
    try:
        import win32com.client
    except ImportError as exc:
        raise engine.headed.HeadedUnavailable("pywin32 not installed") from exc
    try:
        app = win32com.client.Dispatch("Excel.Application")
    except Exception as exc:
        raise engine.headed.HeadedUnavailable(f"Excel.Application unavailable: {exc}") from exc
    try:
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(str(Path(path).resolve()))
        try:
            app.Run(macro_name)
            wb.Save()
            return {"ran": macro_name}
        finally:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
    finally:
        try:
            app.Quit()
        except Exception:
            pass


def _call_macro(params: dict) -> dict:
    path = params["path"]
    macro = params.get("macro", "")
    approved = params.get("macro_approved") is True
    sha = _macro_sha(macro)
    detail = f"sha256:{sha[:16]}"

    check = check_approval(_canonical("macro_run_gated", params), params.get("approval_token"))
    if not approved or not check["allowed"]:
        audit.append("macro_run_gated", path, detail, None, "needs_approval")
        return {
            "ok": False,
            "status": "needs_approval",
            "ops_hash": check["ops_hash"],
            "diff": {"macro_sha256": sha},
        }

    if params.get("mode", "headless") != "headed":
        audit.append("macro_run_gated", path, detail, None, "rejected_headless")
        return {
            "ok": False,
            "status": "rejected",
            "reason": "macro execution requires headed mode (Excel Desktop); headless cannot run VBA",
        }

    try:
        backup_id = versions.backup(path, "macro_run_gated", "macro")
    except (ValueError, FileNotFoundError, OSError) as exc:
        return {"ok": False, "status": "error", "reason": str(exc)}
    try:
        result = _run_macro_com(path, str(macro))
    except Exception as exc:
        audit.append("macro_run_gated", path, detail, backup_id, "ERROR")
        return {"ok": False, "status": "error", "reason": str(exc), "backup_id": backup_id}

    MACRO_LOG.append({"path": path, "sha256": sha})
    audit.append("macro_run_gated", path, detail, backup_id, "ok")
    return {"ok": True, "status": "ok", "macro_sha256": sha, "backup_id": backup_id, **result}


def call_tool(tool: str, params: dict) -> dict:
    params = dict(params or {})
    path = params.get("path", "")
    if not guard.is_path_allowed(path):
        return {"ok": False, "status": "rejected", "reason": f"path outside ALLOWED_ROOTS: {path!r}"}
    if tool in READ_TOOLS:
        return _call_read(tool, params)
    if tool == "macro_run_gated":
        return _call_macro(params)
    if tool in GATED_TOOLS:
        return _call_gated(tool, params)
    return {"ok": False, "status": "rejected", "reason": f"unknown tool: {tool!r}"}


@mcp.tool()
def list_sheets(path: str) -> dict:
    """List worksheets with dimensions, freeze panes and print setup. Always allowed."""
    return call_tool("list_sheets", {"path": path})


@mcp.tool()
def read_range(path: str, sheet: str, cell_range: str, mode: str = "headless") -> dict:
    """Read values, formulas and number formats from a range. Always allowed. Mode headless|headed."""
    return call_tool("read_range", {"path": path, "sheet": sheet, "cell_range": cell_range, "mode": mode})


@mcp.tool()
def preview_diff(path: str, ops: dict) -> dict:
    """Show before/after for a write op without touching disk. Always allowed, never writes."""
    return call_tool("preview_diff", {"path": path, "ops": ops})


@mcp.tool()
def screenshot_sheet(
    path: str,
    sheet: str,
    cell_range: str | None = None,
    full: bool = False,
    mode: str = "headless",
) -> dict:
    """Thumbnail (always) plus full PNG when full=true. Always allowed."""
    return call_tool(
        "screenshot_sheet",
        {"path": path, "sheet": sheet, "cell_range": cell_range, "full": full, "mode": mode},
    )


@mcp.tool()
def write_range(
    path: str,
    sheet: str,
    cell_range: str,
    values: list,
    approval_token: str | None = None,
    mode: str = "headless",
) -> dict:
    """Write a values grid. Gated: returns needs_approval + ops_hash until the exact hash is approved."""
    return call_tool(
        "write_range",
        {"path": path, "sheet": sheet, "cell_range": cell_range, "values": values, "approval_token": approval_token, "mode": mode},
    )


@mcp.tool()
def set_formula(
    path: str,
    sheet: str,
    cell: str,
    formula: str,
    approval_token: str | None = None,
    mode: str = "headless",
) -> dict:
    """Set one formula cell (must start with =). Gated like write_range."""
    return call_tool(
        "set_formula",
        {"path": path, "sheet": sheet, "cell": cell, "formula": formula, "approval_token": approval_token, "mode": mode},
    )


@mcp.tool()
def add_sheet(
    path: str, name: str, approval_token: str | None = None, mode: str = "headless"
) -> dict:
    """Add a worksheet. Gated like write_range."""
    return call_tool(
        "add_sheet", {"path": path, "name": name, "approval_token": approval_token, "mode": mode}
    )


@mcp.tool()
def format_range(
    path: str,
    sheet: str,
    cell_range: str,
    style: dict,
    approval_token: str | None = None,
    mode: str = "headless",
) -> dict:
    """Apply formatting (font/fill/border/number_format/widths/freeze/print/table/validation/conditional). Gated."""
    return call_tool(
        "format_range",
        {"path": path, "sheet": sheet, "cell_range": cell_range, "style": style, "approval_token": approval_token, "mode": mode},
    )


@mcp.tool()
def add_chart(
    path: str,
    sheet: str,
    data_range: str,
    chart_type: str,
    anchor: str,
    approval_token: str | None = None,
    mode: str = "headless",
) -> dict:
    """Add a line/bar/column/pie/area chart. Gated."""
    return call_tool(
        "add_chart",
        {"path": path, "sheet": sheet, "data_range": data_range, "chart_type": chart_type, "anchor": anchor, "approval_token": approval_token, "mode": mode},
    )


@mcp.tool()
def add_pivot(
    path: str,
    sheet: str,
    source_range: str,
    dest: str,
    rows: list,
    values: list,
    approval_token: str | None = None,
    mode: str = "headless",
) -> dict:
    """Build a SUMIFS summary sheet (native pivots cannot be authored headless). Gated."""
    return call_tool(
        "add_pivot",
        {"path": path, "sheet": sheet, "source_range": source_range, "dest": dest, "rows": rows, "values": values, "approval_token": approval_token, "mode": mode},
    )


@mcp.tool()
def add_image(
    path: str,
    sheet: str,
    image_path: str,
    anchor: str,
    approval_token: str | None = None,
    mode: str = "headless",
) -> dict:
    """Embed an image at an anchor cell. Gated."""
    return call_tool(
        "add_image",
        {"path": path, "sheet": sheet, "image_path": image_path, "anchor": anchor, "approval_token": approval_token, "mode": mode},
    )


@mcp.tool()
def macro_run_gated(
    path: str,
    macro: str,
    macro_approved: bool = False,
    approval_token: str | None = None,
    mode: str = "headless",
) -> dict:
    """Run a named workbook macro. Requires macro_approved=true plus the approval hash, headed mode only. Only the SHA256 is logged, never the macro text."""
    return call_tool(
        "macro_run_gated",
        {"path": path, "macro": macro, "macro_approved": macro_approved, "approval_token": approval_token, "mode": mode},
    )


@mcp.tool()
def rollback(path: str, backup_id: str, approval_token: str | None = None, mode: str = "headless") -> dict:
    """Restore a backup (pre-rollback state is backed up first). Itself gated."""
    return call_tool(
        "rollback", {"path": path, "backup_id": backup_id, "approval_token": approval_token, "mode": mode}
    )


if __name__ == "__main__":
    mcp.run()
