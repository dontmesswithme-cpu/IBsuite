IB_DEFAULTS: dict = {
    "header_fill": "002060",
    "header_font": "white bold 10",
    "body_font": "Arial 10",
    "input_fill": "D9E1F2",
    "hardcode_font": "grey 10",
    "formula_font": "black 10",
    "number_format_int": '_-* #,##0_-;-* #,##0_-;-* "-"_-;_-@_-',
    "number_format_1dp": '_-* #,##0.0_-;-* #,##0.0_-;-* "-"_-;_-@_-',
    "number_format_pct": "%_0.0%",
    "border_grid": "D9D9D9",
    "border_grid_style": "thin",
    "border_subtotal_style": "medium",
    "border_total_style": "double",
    "freeze_panes": "B2",
    "col_width": 16,
    "col_width_min": 14,
    "col_width_max": 18,
    "row_height": 15,
    "print_setup": {
        "fit_to_pages_wide": 1,
        "fit_to_pages_tall": 0,
        "orientation": "landscape",
        "footer": "&F | &A",
    },
}


def get_style(key: str) -> dict | str:
    if key not in IB_DEFAULTS:
        raise KeyError(f"unknown style key: {key!r}")
    return IB_DEFAULTS[key]
