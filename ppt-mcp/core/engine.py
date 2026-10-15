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
    if tool == "list_slides":
        (path,) = _need(tool, params, "path")
        return {"slides": headless.list_slides(path)}
    if tool == "read_text":
        path, slide = _need(tool, params, "path", "slide")
        return headless.read_text(path, int(slide))
    if tool == "add_slide":
        (path,) = _need(tool, params, "path")
        layout = params.get("layout", "blank")
        index = params.get("index")
        return headless.add_slide(path, layout, index)
    if tool == "set_text":
        path, slide, shape, text = _need(tool, params, "path", "slide", "shape", "text")
        return headless.set_text(path, int(slide), shape, text)
    if tool == "format_shape":
        path, slide, shape = _need(tool, params, "path", "slide", "shape")
        return fmt.format_shape(
            path,
            int(slide),
            shape,
            font_name=params.get("font_name"),
            font_size=params.get("font_size"),
            font_color=params.get("font_color"),
            fill_color=params.get("fill_color"),
            bold=params.get("bold"),
        )
    if tool == "add_table":
        path, slide, rows, cols = _need(tool, params, "path", "slide", "rows", "cols")
        values = params.get("values")
        anchor = params.get("anchor")
        return headless.add_table(path, int(slide), int(rows), int(cols), values, anchor)
    if tool == "add_chart":
        path, slide, chart_type, data = _need(tool, params, "path", "slide", "chart_type", "data")
        anchor = params.get("anchor")
        return headless.add_chart(path, int(slide), chart_type, data, anchor)
    if tool == "add_image":
        path, slide, image_path = _need(tool, params, "path", "slide", "image_path")
        anchor = params.get("anchor")
        return headless.add_image(path, int(slide), image_path, anchor)
    if tool == "apply_theme":
        (path,) = _need(tool, params, "path")
        return fmt.apply_theme(path)
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
        slide = int(params.get("slide", 1))
        return headed.headed_read(path, slide)
    if tool in ("set_text", "add_slide"):
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
