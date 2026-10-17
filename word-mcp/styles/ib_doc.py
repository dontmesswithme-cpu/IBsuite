IB_DOC_DEFAULTS: dict = {
    "heading_fill": "002060",
    "heading_font_color": "FFFFFF",
    "heading_font_bold": True,
    "heading_font_name": "Arial",
    "heading_font_size": 14,
    "body_font": "Arial 10",
    "body_font_name": "Arial",
    "body_font_size": 10,
    "body_font_color": "000000",
    "table_grid_color": "D9D9D9",
    "table_header_fill": "002060",
    "table_header_color": "FFFFFF",
    "page_size": "Letter",
}


def get_doc_style(key: str) -> dict | str | float | tuple:
    if key not in IB_DOC_DEFAULTS:
        raise KeyError(f"unknown style key: {key!r}")
    return IB_DOC_DEFAULTS[key]
