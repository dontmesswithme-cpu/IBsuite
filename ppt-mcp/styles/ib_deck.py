IB_DECK_DEFAULTS: dict = {
    "header_fill": "002060",
    "header_font_color": "FFFFFF",
    "header_font_bold": True,
    "header_font_size": 18,
    "body_font_name": "Arial",
    "body_font_size": 12,
    "body_font_color": "000000",
    "table_grid_color": "D9D9D9",
    "slide_width_16x9_in": (13.33, 7.5),
    "slide_size_16x9": "16:9",
}


def get_deck_style(key: str) -> dict | str | float | tuple:
    if key not in IB_DECK_DEFAULTS:
        raise KeyError(f"unknown style key: {key!r}")
    return IB_DECK_DEFAULTS[key]
