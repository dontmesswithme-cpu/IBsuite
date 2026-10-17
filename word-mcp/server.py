from pathlib import Path

from mcp.server.mcpserver import MCPServer

from core import audit, engine, guard, versions
from core.guard import check_approval, make_ops_hash
from core.shots import full_shot, thumbnail

mcp = MCPServer(name="word-mcp")

READ_TOOLS = ("list_structure", "read_text", "preview_diff", "screenshot_doc")
GATED_TOOLS = (
    "add_paragraph",
    "set_heading",
    "format_text",
    "add_table",
    "add_image",
    "add_header_footer",
    "apply_theme",
    "set_page_setup",
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
    if tool in ("add_paragraph", "set_heading", "format_text"):
        return str(params.get("style", params.get("text", "op")))[:20]
    if tool in ("add_table", "add_image"):
        return str(params.get("anchor", "doc"))
    if tool == "rollback":
        return str(params.get("backup_id", "rollback"))
    return "op"


def _audit_detail(tool: str, params: dict) -> str:
    if tool == "add_paragraph":
        return _sanitize(f"paragraph:{params.get('text', '')[:40]}")
    if tool == "set_heading":
        return _sanitize(f"heading:{params.get('index', '')}:{params.get('level', '')}")
    if tool == "format_text":
        return _sanitize(f"format:{params.get('start', '')}-{params.get('end', '')}:{params.get('style', '')}")
    if tool == "add_table":
        return _sanitize(f"table:{params.get('rows', '')}x{params.get('cols', '')}")
    if tool == "add_image":
        image = Path(str(params.get("image_path", ""))).name
        return _sanitize(f"image:{image}")
    if tool == "add_header_footer":
        return _sanitize(f"{params.get('kind', '')}:{params.get('text', '')[:40]}")
    if tool == "apply_theme":
        return _sanitize("ib_doc_theme")
    if tool == "set_page_setup":
        return _sanitize(f"pagesetup:{params.get('size', '')}:{params.get('orientation', '')}")
    if tool == "rollback":
        return _sanitize(f"rollback:{params.get('backup_id', '')}")
    return _sanitize(tool)


def _best_thumb(params: dict):
    try:
        return thumbnail(params["path"])
    except Exception:
        return None


def _screenshot(params: dict) -> dict:
    path = params["path"]
    page = int(params.get("page", params.get("slide", 1)))
    full = bool(params.get("full", False))
    mode = params.get("mode", "headless")
    out = {"thumbnail_b64": thumbnail(path, page=page), "full_b64": None}
    if full:
        out["full_b64"] = full_shot(path, page=page, mode=mode)
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

    if name == "screenshot_doc":
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


@mcp.tool(name="list_structure")
def tool_list_structure(path: str) -> dict:
    return call_tool("list_structure", {"path": path})


@mcp.tool(name="read_text")
def tool_read_text(path: str, start: int = 0, end: int | None = None, mode: str = "headless") -> dict:
    args = {"path": path, "start": start, "mode": mode}
    if end is not None:
        args["end"] = end
    return call_tool("read_text", args)


@mcp.tool(name="preview_diff")
def tool_preview_diff(path: str, ops: dict) -> dict:
    return call_tool("preview_diff", {"path": path, "ops": ops})


@mcp.tool(name="screenshot_doc")
def tool_screenshot_doc(path: str, page: int = 1, full: bool = False, mode: str = "headless") -> dict:
    return call_tool("screenshot_doc", {"path": path, "page": page, "full": full, "mode": mode})


@mcp.tool(name="add_paragraph")
def tool_add_paragraph(path: str, text: str = "", style: str = "Normal", mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("add_paragraph", {"path": path, "text": text, "style": style, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="set_heading")
def tool_set_heading(path: str, index: int, level: int, text: str, mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("set_heading", {"path": path, "index": index, "level": level, "text": text, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="format_text")
def tool_format_text(path: str, start: int, end: int, style: str, mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("format_text", {"path": path, "start": start, "end": end, "style": style, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="add_table")
def tool_add_table(path: str, anchor: str = "end", rows: int = 1, cols: int = 1, values: list | None = None, mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("add_table", {"path": path, "anchor": anchor, "rows": rows, "cols": cols, "values": values, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="add_image")
def tool_add_image(path: str, anchor: str = "end", image_path: str = "", mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("add_image", {"path": path, "anchor": anchor, "image_path": image_path, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="add_header_footer")
def tool_add_header_footer(path: str, kind: str, text: str, mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("add_header_footer", {"path": path, "kind": kind, "text": text, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="apply_theme")
def tool_apply_theme(path: str, mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("apply_theme", {"path": path, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="set_page_setup")
def tool_set_page_setup(path: str, size: str = "Letter", margins = None, orientation: str = "portrait", mode: str = "headless", approval_token: str | None = None) -> dict:
    return call_tool("set_page_setup", {"path": path, "size": size, "margins": margins, "orientation": orientation, "mode": mode, "approval_token": approval_token})


@mcp.tool(name="rollback")
def tool_rollback(path: str, backup_id: str, approval_token: str | None = None) -> dict:
    return call_tool("rollback", {"path": path, "backup_id": backup_id, "approval_token": approval_token})


if __name__ == "__main__":
    mcp.run()
