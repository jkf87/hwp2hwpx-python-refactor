"""Value conversion maps for HWP binary values to HWPX XML enum strings."""


def color_from_int(val):
    """Convert HWP color int (0xBBGGRR or -1 for none) to HWPX hex string."""
    if val is None or val == -1 or val == 0xFFFFFFFF or val == 4294967295:
        return "none"
    r = val & 0xFF
    g = (val >> 8) & 0xFF
    b = (val >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


def color_from_int_with_alpha(val):
    """Convert with full 4-byte handling, return (color, alpha)."""
    if val is None or val == -1 or val == 0xFFFFFFFF:
        return "none", "0"
    r = val & 0xFF
    g = (val >> 8) & 0xFF
    b = (val >> 16) & 0xFF
    a = (val >> 24) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}", str(a)


# Gradation type
GRADATION_TYPE_MAP = {
    1: "LINEAR",
    2: "RADIAL",
    3: "CONICAL",
    4: "SQUARE",
}


# Line/border stroke type (stroke_flags lower 4 bits)
STROKE_TYPE_MAP = {
    0: "NONE",
    1: "SOLID",
    2: "DASH",
    3: "DOT",
    4: "DASH_DOT",
    5: "DASH_DOT_DOT",
    6: "LONG_DASH",
    7: "CIRCLE",
    8: "DOUBLE_SLIM",
    9: "SLIM_THICK",
    10: "THICK_SLIM",
    11: "SLIM_THICK_SLIM",
    12: "WAVE",
    13: "DOUBLE_WAVE",
    14: "THICK_3D",
    15: "THICK_3D_REVERSE",
    16: "3D",
    17: "3D_REVERSE",
}

# Border width from width_flags
BORDER_WIDTH_MAP = {
    0: "0.1 mm",
    1: "0.12 mm",
    2: "0.15 mm",
    3: "0.2 mm",
    4: "0.25 mm",
    5: "0.3 mm",
    6: "0.4 mm",
    7: "0.5 mm",
    8: "0.6 mm",
    9: "0.7 mm",
    10: "1.0 mm",
    11: "1.5 mm",
    12: "2.0 mm",
    13: "3.0 mm",
    14: "4.0 mm",
    15: "5.0 mm",
}

# Horizontal alignment
HALIGN_MAP = {
    0: "JUSTIFY",
    1: "LEFT",
    2: "RIGHT",
    3: "CENTER",
    4: "DISTRIBUTE",
    5: "DISTRIBUTE_SPACE",
}

# Vertical alignment
VALIGN_MAP = {
    0: "BASELINE",
    1: "TOP",
    2: "CENTER",
    3: "BOTTOM",
}

# Line spacing type
LINESPACING_TYPE_MAP = {
    0: "PERCENT",
    1: "FIXED",
    2: "BETWEEN_LINES",
    3: "AT_LEAST",
}

# Heading type
HEADING_TYPE_MAP = {
    0: "NONE",
    1: "OUTLINE",
    2: "NUMBER",
    3: "BULLET",
}

# Break latin word
BREAK_LATIN_WORD_MAP = {
    0: "KEEP_WORD",
    1: "HYPHENATION",
    2: "BREAK_WORD",
}

# Break non-latin word
BREAK_NON_LATIN_WORD_MAP = {
    0: "KEEP_WORD",
    1: "BREAK_WORD",
}

# Line wrap
LINE_WRAP_MAP = {
    0: "BREAK",
    1: "SQUEEZE",
    2: "KEEP",
}

# Style type
STYLE_TYPE_MAP = {
    0: "PARA",
    1: "CHAR",
}

# Underline type
UNDERLINE_TYPE_MAP = {
    0: "NONE",
    1: "BOTTOM",
    2: "CENTER",
    3: "TOP",
}

# Underline/strikeout shape
LINE_SHAPE_MAP = {
    0: "SOLID",
    1: "DASH",
    2: "DOT",
    3: "DASH_DOT",
    4: "DASH_DOT_DOT",
    5: "LONG_DASH",
    6: "CIRCLE",
    7: "DOUBLE_SLIM",
    8: "SLIM_THICK",
    9: "THICK_SLIM",
    10: "SLIM_THICK_SLIM",
}

# Strikeout type
STRIKEOUT_TYPE_MAP = {
    0: "NONE",
    1: "CONTINUOUS",
    2: "SUPERFIX",
    3: "SUBFIX",
}

# Outline type
OUTLINE_TYPE_MAP = {
    0: "NONE",
    1: "SOLID",
    2: "DASH",
    3: "DOT",
    4: "DASH_DOT",
    5: "DASH_DOT_DOT",
    6: "LONG_DASH",
    7: "CIRCLE",
    8: "DOUBLE_SLIM",
    9: "SLIM_THICK",
    10: "THICK_SLIM",
    11: "SLIM_THICK_SLIM",
}

# Shadow type
SHADOW_TYPE_MAP = {
    0: "NONE",
    1: "DROP",
    2: "CONTINUOUS",
}

# Emphasis mark
SYM_MARK_MAP = {
    0: "NONE",
    1: "DOT_ABOVE",
    2: "RING_ABOVE",
    3: "TILDE",
    4: "CARON",
    5: "SIDE_DOT",
    6: "COLON",
    7: "GRAVE_ACCENT",
    8: "ACUTE_ACCENT",
    9: "CIRCUMFLEX",
    10: "MACRON",
    11: "HOOK_ABOVE",
    12: "HORN",
}

# Numbering / format enums
NUM_FORMAT_MAP = {
    0: "DIGIT",
    1: "CIRCLED_DIGIT",
    2: "ROMAN_CAPITAL",
    3: "ROMAN_SMALL",
    4: "LATIN_CAPITAL",
    5: "LATIN_SMALL",
    6: "CIRCLED_LATIN_CAPITAL",
    7: "CIRCLED_LATIN_SMALL",
    8: "HANGUL_SYLLABLE",
    9: "CIRCLED_HANGUL_SYLLABLE",
    10: "HANGUL_JAMO",
    11: "CIRCLED_HANGUL_JAMO",
    12: "HANGUL_PHONETIC",
    13: "IDEOGRAPH",
    14: "CIRCLED_IDEOGRAPH",
}

# Text offset type
TEXT_OFFSET_TYPE_MAP = {
    0: "PERCENT",
    1: "HWPUNIT",
}

# Page starts on
PAGE_STARTS_ON_MAP = {
    0: "BOTH",
    1: "EVEN",
    2: "ODD",
}

# Footnote numbering type
FOOTNOTE_NUMBERING_MAP = {
    0: "CONTINUOUS",
    1: "ON_SECTION",
    2: "ON_PAGE",
}

# Endnote numbering type
ENDNOTE_NUMBERING_MAP = {
    0: "CONTINUOUS",
    1: "ON_SECTION",
}

# Footnote/Endnote placement
FOOTNOTE_PLACE_MAP = {
    0: "EACH_COLUMN",
    1: "RIGHT_MOST_COLUMN",
    2: "BENEATH_TEXT",
}

ENDNOTE_PLACE_MAP = {
    0: "END_OF_DOCUMENT",
    1: "END_OF_SECTION",
}

# Page break type for tables
PAGE_BREAK_MAP = {
    0: "NONE",
    1: "CELL",
    2: "TABLE",
}

# Text wrap
TEXT_WRAP_MAP = {
    0: "SQUARE",
    1: "TIGHT",
    2: "THROUGH",
    3: "TOP_AND_BOTTOM",
    4: "BEHIND_TEXT",
    5: "IN_FRONT_OF_TEXT",
}

# Text flow
TEXT_FLOW_MAP = {
    0: "BOTH_SIDES",
    1: "LEFT_ONLY",
    2: "RIGHT_ONLY",
    3: "LARGEST_ONLY",
}

# Vert/Horz relative to
VERT_REL_TO_MAP = {
    0: "PAPER",
    1: "PAGE",
    2: "PARA",
}

HORZ_REL_TO_MAP = {
    0: "PAPER",
    1: "PAGE",
    2: "COLUMN",
    3: "PARA",
}

VERT_ALIGN_MAP = {
    0: "TOP",
    1: "CENTER",
    2: "BOTTOM",
    3: "INSIDE",
    4: "OUTSIDE",
}

HORZ_ALIGN_MAP = {
    0: "LEFT",
    1: "CENTER",
    2: "RIGHT",
    3: "INSIDE",
    4: "OUTSIDE",
}

# Column type
COLUMN_TYPE_MAP = {
    0: "NEWSPAPER",
    1: "BALANCED_NEWSPAPER",
    2: "PARALLEL",
}

# Column layout
COLUMN_LAYOUT_MAP = {
    0: "LEFT",
    1: "RIGHT",
    2: "MIRROR",
}

# Text direction
TEXT_DIRECTION_MAP = {
    0: "HORIZONTAL",
    1: "VERTICAL",
}

# Cell vertical alignment (listflags bits 8-9)
CELL_VERT_ALIGN_MAP = {
    0: "TOP",
    1: "CENTER",
    2: "BOTTOM",
}

# Font family type (from PANOSE)
FAMILY_TYPE_MAP = {
    0: "FCAT_UNKNOWN",
    1: "FCAT_MYEONGJO",
    2: "FCAT_GOTHIC",
    3: "FCAT_GRAPHIC",
    4: "FCAT_CURSIVE",
    5: "FCAT_DECORATIVE",
    6: "FCAT_NON_LATIN",
}

# Language code for fontface groups
FONT_LANG_KEYS = ["ko", "en", "cn", "jp", "other", "symbol", "user"]
FONT_LANG_NAMES = ["HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"]
FONT_LANG_ATTR_NAMES = ["hangul", "latin", "hanja", "japanese", "other", "symbol", "user"]

# Numbering format suffix char
SUFFIX_CHAR_MAP = {
    0: "",
    41: ")",
    46: ".",
}

# Gutter type
GUTTER_TYPE_MAP = {
    0: "LEFT_ONLY",
    1: "LEFT_RIGHT",
    2: "TOP_BOTTOM",
}

# Fill area
FILL_AREA_MAP = {
    0: "PAPER",
    1: "PAGE",
    2: "BORDER",
}

# Border fill type
BORDER_FILL_TYPE_MAP = {
    0: "BOTH",
    1: "EVEN",
    2: "ODD",
}

# Compatible document target
COMPATIBLE_TARGET_MAP = {
    0: "HWP201X",
    1: "HWP200X",
    2: "MS_WORD",
}

# Image type from binary data extension
IMAGE_TYPE_MAP = {
    "bmp": ("image/bmp", "BMP"),
    "emf": ("image/x-emf", "EMF"),
    "gif": ("image/gif", "GIF"),
    "jpg": ("image/jpeg", "JPG"),
    "jpeg": ("image/jpeg", "JPG"),
    "png": ("image/png", "PNG"),
    "svg": ("image/svg+xml", "SVG"),
    "tif": ("image/tiff", "TIF"),
    "tiff": ("image/tiff", "TIF"),
    "wmf": ("image/x-wmf", "WMF"),
    "ole": ("application/x-ole", "OLE"),
}


# Picture effect type
PICTURE_EFFECT_MAP = {
    0: "REAL_PIC",
    1: "GRAY_SCALE",
    2: "BLACK_WHITE",
    3: "PATTERN8x8",
}


def extract_charshape_flags(flags):
    """Extract charshape boolean flags from the flags integer."""
    return {
        "bold": bool(flags & 0x01),
        "italic": bool(flags & 0x02),
        "underline_type": (flags >> 2) & 0x07,
        "underline_shape": (flags >> 5) & 0x0F,
        "outline_type": (flags >> 9) & 0x07,
        "shadow_type": (flags >> 12) & 0x03,
        "emboss": bool(flags & (1 << 14)),
        "engrave": bool(flags & (1 << 15)),
        "superscript": bool(flags & (1 << 16)),
        "subscript": bool(flags & (1 << 17)),
        "strikeout_type": (flags >> 18) & 0x07,
        "sym_mark": (flags >> 21) & 0x0F,
        "use_font_space": bool(flags & (1 << 25)),
        "use_kerning": bool(flags & (1 << 26)),
    }


def extract_parashape_flags(flags):
    """Extract parashape flags from the flags integer."""
    return {
        "halign": (flags >> 2) & 0x07,
        "break_latin_word": (flags >> 5) & 0x03,
        "break_non_latin_word": bool(flags & (1 << 7)),
        "snap_to_grid": bool(flags & (1 << 8)),
        "condense": (flags >> 9) & 0x7F,
        "widow_orphan": bool(flags & (1 << 16)),
        "keep_with_next": bool(flags & (1 << 17)),
        "keep_lines": bool(flags & (1 << 18)),
        "page_break_before": bool(flags & (1 << 19)),
        "valign": (flags >> 20) & 0x03,
        "font_line_height": bool(flags & (1 << 22)),
        "heading_type": (flags >> 23) & 0x03,
        "heading_level": (flags >> 25) & 0x07,
        "line_wrap": (flags >> 28) & 0x03,
        "auto_spacing_ea_eng": bool(flags & (1 << 30)),
        "auto_spacing_ea_num": bool(flags & (1 << 31)) if flags >= 0 else False,
    }


def extract_parashape_flags2(flags2):
    """Extract parashape flags2."""
    return {
        "suppress_line_numbers": bool(flags2 & 0x01) if flags2 else False,
    }


def extract_linespacing_type(flags):
    """Extract line spacing type from parashape flags."""
    return (flags >> 0) & 0x03


def border_info(border_dict):
    """Convert HWP border dict to HWPX attributes."""
    if not border_dict:
        return {"type": "NONE", "width": "0.1 mm", "color": "#000000"}
    stroke = border_dict.get("stroke_flags", 0) & 0x1F
    width = border_dict.get("width_flags", 0)
    color = border_dict.get("color", 0)
    return {
        "type": STROKE_TYPE_MAP.get(stroke, "NONE"),
        "width": BORDER_WIDTH_MAP.get(width, "0.1 mm"),
        "color": color_from_int(color),
    }


def extract_page_border_fill_flags(flags):
    """Extract page border fill type from flags.

    The Java hwp2hwpx uses a separate index-based approach:
    pageBorderFill models appear in order: BOTH(0), EVEN(1), ODD(2).
    The 'type' is determined by the order of appearance, not from flags.
    """
    text_border = (flags >> 2) & 0x01
    header_inside = (flags >> 3) & 0x01
    footer_inside = (flags >> 4) & 0x01
    fill_area = (flags >> 5) & 0x03
    return {
        "textBorder": "BODY" if text_border else "PAPER",
        "headerInside": str(header_inside),
        "footerInside": str(footer_inside),
        "fillArea": FILL_AREA_MAP.get(fill_area, "PAPER"),
    }
