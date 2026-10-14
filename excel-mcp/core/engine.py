import core.advanced_headless as advanced
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
    if tool == "list_sheets":
        (path,) = _need(tool, params, "path")
        return {"sheets": headless.list_sheets(path)}
    if tool == "read_range":
        path, sheet, cell_range = _need(tool, params, "path", "sheet", "cell_range")
        return headless.read_range(path, sheet, cell_range)
    if tool == "write_range":
        path, sheet, cell_range, values = _need(tool, params, "path", "sheet", "cell_range", "values")
        return headless.write_range(path, sheet, cell_range, values)
    if tool == "set_formula":
        path, sheet, cell, formula = _need(tool, params, "path", "sheet", "cell", "formula")
        return headless.set_formula(path, sheet, cell, formula)
    if tool == "add_sheet":
        path, name = _need(tool, params, "path", "name")
        return headless.add_sheet(path, name)
    if tool == "preview_diff":
        path, ops = _need(tool, params, "path", "ops")
        return headless.preview_diff(path, ops)
    if tool == "format_range":
        path, sheet, cell_range, style = _need(tool, params, "path", "sheet", "cell_range", "style")
        return fmt.format_range(path, sheet, cell_range, style)
    if tool == "apply_ib_table":
        path, sheet, cell_range, has_totals = _need(tool, params, "path", "sheet", "cell_range", "has_totals")
        return fmt.apply_ib_table(path, sheet, cell_range, has_totals)
    if tool == "add_chart":
        path, sheet, data_range, chart_type, anchor = _need(tool, params, "path", "sheet", "data_range", "chart_type", "anchor")
        return advanced.add_chart(path, sheet, data_range, chart_type, anchor)
    if tool == "add_pivot":
        path, sheet, source_range, dest, rows, values = _need(tool, params, "path", "sheet", "source_range", "dest", "rows", "values")
        return advanced.add_pivot(path, sheet, source_range, dest, rows, values)
    if tool == "add_image":
        path, sheet, image_path, anchor = _need(tool, params, "path", "sheet", "image_path", "anchor")
        return advanced.add_image(path, sheet, image_path, anchor)
    if tool == "rollback":
        path, backup_id = _need(tool, params, "path", "backup_id")
        return versions.rollback(path, backup_id)
    if tool == "list_versions":
        (path,) = _need(tool, params, "path")
        return {"versions": versions.list_versions(path)}
    if tool == "macro_run_gated":
        raise ValueError("macro_run_gated is server-only; engine.dispatch refuses to run macros")
    raise ValueError(f"unknown tool: {tool!r}")


def _dispatch_headed(tool: str, params: dict) -> dict:
    if tool == "read_range":
        path, sheet, cell_range = _need(tool, params, "path", "sheet", "cell_range")
        return headed.headed_read(path, sheet, cell_range)
    if tool in ("write_range", "set_formula", "add_sheet"):
        (path,) = _need(tool, params, "path")
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
