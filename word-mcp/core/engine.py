import core.format_headless as fmt
import core.headless as headless
import core.versions as versions
from core import headed

FALLBACK_WARNING = "headed_unavailable_fell_back"


def _need(tool: str, params: dict, *keys: str) -> list:
    missing = [key for key in keys if key not in params]
    if missing:
        raise ValueError(f"{tool} missing params: {missing}")
    return [params[key] for key in keys]


def _dispatch_headless(tool: str, params: dict) -> dict:
    if tool == "list_structure":
        (path,) = _need(tool, params, "path")
        return {"paragraphs": headless.list_structure(path)}
    if tool == "read_text":
        path = params["path"]
        start = int(params.get("start", 0))
        end = params.get("end", None)
        if end is None:
            # read all
            doc_paras = headless.list_structure(path)
            end = len(doc_paras) - 1 if doc_paras else 0
        return headless.read_text(path, start, int(end))
    if tool == "add_paragraph":
        path = params["path"]
        text = params.get("text", "")
        style = params.get("style", "Normal")
        return headless.add_paragraph(path, text, style)
    if tool == "set_heading":
        path, index, level, text = _need(tool, params, "path", "index", "level", "text")
        return headless.set_heading(path, int(index), int(level), text)
    if tool == "format_text":
        path, start, end, style = _need(tool, params, "path", "start", "end", "style")
        return headless.format_text(path, int(start), int(end), style)
    if tool == "add_table":
        path = params["path"]
        anchor = params.get("anchor", "end")
        rows = int(params.get("rows", 1))
        cols = int(params.get("cols", 1))
        values = params.get("values")
        return headless.add_table(path, anchor, rows, cols, values)
    if tool == "add_image":
        path = params["path"]
        anchor = params.get("anchor", "end")
        image_path = params["image_path"]
        return headless.add_image(path, anchor, image_path)
    if tool == "add_header_footer":
        path, kind, text = _need(tool, params, "path", "kind", "text")
        return headless.add_header_footer(path, kind, text)
    if tool == "apply_theme":
        (path,) = _need(tool, params, "path")
        return fmt.apply_theme(path)
    if tool == "set_page_setup":
        path = params["path"]
        size = params.get("size", "Letter")
        margins = params.get("margins")
        orientation = params.get("orientation", "portrait")
        return headless.set_page_setup(path, size, margins, orientation)
    if tool == "preview_diff":
        path, ops = _need(tool, params, "path", "ops")
        return headless.preview_diff(path, ops)
    if tool == "rollback":
        path, backup_id = _need(tool, params, "path", "backup_id")
        return versions.rollback(path, backup_id)
    if tool == "list_versions":
        (path,) = _need(tool, params, "path")
        return {"versions": versions.list_versions(path)}
    raise ValueError(f"unknown tool: {tool!r}")


def _dispatch_headed(tool: str, params: dict) -> dict:
    if tool == "read_text":
        path = params["path"]
        start = int(params.get("start", 0))
        end = params.get("end", None)
        if end is None:
            # attempt all
            end = None
        return headed.headed_read(path, start, end)
    if tool in ("add_paragraph", "set_heading", "apply_theme", "add_table", "add_image", "format_text", "add_header_footer", "set_page_setup"):
        path = params["path"]
        return headed.headed_apply(path, {"tool": tool, **params})
    raise headed.HeadedUnavailable(f"headed mode does not implement {tool!r}")


def dispatch(tool: str, params: dict) -> dict:
    if not isinstance(params, dict):
        raise ValueError(f"dispatch params must be a dict for {tool!r}")
    mode = params.get("mode", "headless")
    clean = {key: value for key, value in params.items() if key != "mode"}

    if mode == "headless":
        return _dispatch_headless(tool, clean)
    if mode == "headed":
        try:
            return _dispatch_headed(tool, clean)
        except headed.HeadedUnavailable as exc:
            result = _dispatch_headless(tool, clean)
            if isinstance(result, dict):
                result["warning"] = FALLBACK_WARNING
                result["headed_error"] = str(exc)
            return result
    raise ValueError(f"unknown mode: {mode!r} (expected 'headless' or 'headed')")
