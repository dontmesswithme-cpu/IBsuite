import hashlib
from pathlib import Path

from mcp.server.mcpserver import MCPServer

from core import audit, engine, guard, versions
from core.guard import check_approval, make_ops_hash
from core.shots import full_shot, thumbnail

mcp = MCPServer(name="ppt-mcp")

READ_TOOLS = ("list_slides", "read_text", "preview_diff", "screenshot_slide")
GATED_TOOLS = (
    "add_slide",
    "set_text",
    "format_shape",
    "add_table",
    "add_chart",
    "add_image",
    "apply_theme",
    "rollback",
)

_HASH_EXCLUDE = ("approval_token", "mode")


def _canonical(tool: str, params: dict) -> dict:
    return {"tool": tool, **{k: params[k] for k in params if k not in _HASH_EXCLUDE}}


def ops_hash(tool: str, params: dict) -> str:
    return make_ops_hash(_canonical(tool, params))


def _sanitize(text: str) -> str:
    return str(text).replace("|", "/").replace("\n", " ")


def _backup_label(tool: str, params: dict) -> str:
    if tool in ("set_text", "format_shape"):
        return str(params.get("shape", "shape"))
    if tool == "add_slide":
        return str(params.get("layout", "slide"))
    if tool in ("add_table", "add_chart", "add_image"):
        return str(params.get("slide", "slide"))
    if tool == "rollback":
        return str(params.get("backup_id", "rollback"))
    return "op"


def _audit_detail(tool: str, params: dict) -> str:
    slide = params.get("slide", "")
    if tool == "set_text":
        return _sanitize(f"slide:{slide}:{params.get('shape', '')}")
    if tool == "format_shape":
        return _sanitize(f"slide:{slide}:{params.get('shape', '')}")
    if tool == "add_slide":
        return _sanitize(str(params.get("layout", "blank")))
    if tool == "add_table":
        return _sanitize(f"slide:{slide}:{params.get('rows', '')}x{params.get('cols', '')}")
    if tool == "add_chart":
        return _sanitize(f"slide:{slide}:{params.get('chart_type', '')}")
    if tool == "add_image":
        image = Path(str(params.get("image_path", ""))).name
        return _sanitize(f"slide:{slide}:{image}")
    if tool == "apply_theme":
        return _sanitize("ib_deck_theme")
    if tool == "rollback":
        return _sanitize(f"rollback:{params.get('backup_id', '')}")
    return _sanitize(tool)


def _best_thumb(params: dict):
    try:
        return thumbnail(params["path"], slide=int(params.get("slide", 1)))
    except Exception:
        return None


def _screenshot(params: dict) -> dict:
    path = params["path"]
    slide = int(params.get("slide", 1))
    full = bool(params.get("full", False))
    mode = params.get("mode", "headless")
    out = {"thumbnail_b64": thumbnail(path, slide), "full_b64": None}
    if full:
        out["full_b64"] = full_shot(path, slide, mode)
    return out


def call_tool(name: str, arguments: dict | None = None) -> dict:
    params = dict(arguments or {})
    path = params.get("path")

    if not path or not guard.is_path_allowed(path):
        return {
            "ok": False,
            "status": "rejected",
            "reason": f"denied path: {path}",
        }

    if name == "screenshot_slide":
        try:
            return {"ok": True, **_screenshot(params)}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    if name in READ_TOOLS:
        try:
            res = engine.dispatch(name, params)
            return {"ok": True, **res}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    if name not in GATED_TOOLS:
        return {"ok": False, "status": "rejected", "reason": f"unknown tool: {name}"}

    canonical_ops = _canonical(name, params)
    approval_token = params.get("approval_token")
    approval = check_approval(canonical_ops, approval_token)

    if not approval.get("allowed"):
        return {
            "ok": False,
            "status": "needs_approval",
            "allowed": False,
            "ops_hash": approval["ops_hash"],
            "diff": approval.get("diff"),
        }

    label = _backup_label(name, params)
    detail = _audit_detail(name, params)
    backup_id = None

    try:
        if Path(path).is_file():
            backup_id = versions.backup(path, name, label)
    except Exception as exc:
        audit.append(name, path, detail, None, "ERROR")
        return {"ok": False, "error": f"backup failed: {exc}"}

    try:
        result = engine.dispatch(name, params)
        status = "ok" if result.get("ok", True) else "ERROR"
        audit.append(name, path, detail, backup_id, status)
        thumb = _best_thumb(params)
        return {
            "ok": True,
            "backup_id": backup_id,
            "thumbnail_b64": thumb,
            **result,
        }
    except Exception as exc:
        audit.append(name, path, detail, backup_id, "ERROR")
        return {"ok": False, "error": str(exc)}


@mcp.tool(name="list_slides")
def tool_list_slides(path: str) -> dict:
    return call_tool("list_slides", {"path": path})


@mcp.tool(name="read_text")
def tool_read_text(path: str, slide: int = 1, mode: str = "headless") -> dict:
    return call_tool("read_text", {"path": path, "slide": slide, "mode": mode})


@mcp.tool(name="preview_diff")
def tool_preview_diff(path: str, ops: dict) -> dict:
    return call_tool("preview_diff", {"path": path, "ops": ops})


@mcp.tool(name="screenshot_slide")
def tool_screenshot_slide(path: str, slide: int = 1, full: bool = False, mode: str = "headless") -> dict:
    return call_tool("screenshot_slide", {"path": path, "slide": slide, "full": full, "mode": mode})


@mcp.tool(name="add_slide")
def tool_add_slide(path: str, layout: str = "blank", index: int | None = None, mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("add_slide", {"path": path, "layout": layout, "index": index, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="set_text")
def tool_set_text(path: str, slide: int, shape: str, text: str, mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("set_text", {"path": path, "slide": slide, "shape": shape, "text": text, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="format_shape")
def tool_format_shape(path: str, slide: int, shape: str, font_name: str | None = None, font_size: int | None = None, font_color: str | None = None, fill_color: str | None = None, bold: bool | None = None, approval_token: str | None = None) -> dict:
    return call_tool("format_shape", {"path": path, "slide": slide, "shape": shape, "font_name": font_name, "font_size": font_size, "font_color": font_color, "fill_color": fill_color, "bold": bold, "approval_token": approval_token})


@mcp.tool(name="add_table")
def tool_add_table(path: str, slide: int, rows: int, cols: int, values: list[list] | None = None, anchor: str | None = None, approval_token: str | None = None) -> dict:
    return call_tool("add_table", {"path": path, "slide": slide, "rows": rows, "cols": cols, "values": values, "anchor": anchor, "approval_token": approval_token})


@mcp.tool(name="add_chart")
def tool_add_chart(path: str, slide: int, chart_type: str, data: dict, anchor: str | None = None, approval_token: str | None = None) -> dict:
    return call_tool("add_chart", {"path": path, "slide": slide, "chart_type": chart_type, "data": data, "anchor": anchor, "approval_token": approval_token})


@mcp.tool(name="add_image")
def tool_add_image(path: str, slide: int, image_path: str, anchor: str | None = None, approval_token: str | None = None) -> dict:
    return call_tool("add_image", {"path": path, "slide": slide, "image_path": image_path, "anchor": anchor, "approval_token": approval_token})


@mcp.tool(name="apply_theme")
def tool_apply_theme(path: str, approval_token: str | None = None) -> dict:
    return call_tool("apply_theme", {"path": path, "approval_token": approval_token})


@mcp.tool(name="rollback")
def tool_rollback(path: str, backup_id: str, approval_token: str | None = None) -> dict:
    return call_tool("rollback", {"path": path, "backup_id": backup_id, "approval_token": approval_token})


if __name__ == "__main__":
    mcp.run()
