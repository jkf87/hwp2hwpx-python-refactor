"""Convert HWP DocInfo to HWPX header.xml."""

from .xml_builder import root_element, sub, make_tag
from . import value_maps as vm


def build_header_xml(reader):
    """Build the header.xml element tree from HWP docinfo."""
    props = reader.get_document_properties()
    section_count = props.get("section_count", 1)

    head = root_element("hh", "head")
    head.set("version", "1.4")
    head.set("secCnt", str(section_count))

    # beginNum
    begin_num = sub(head, "hh", "beginNum")
    begin_num.set("page", str(props.get("page_startnum", 1)))
    begin_num.set("footnote", str(props.get("footnote_startnum", 1)))
    begin_num.set("endnote", str(props.get("endnote_startnum", 1)))
    begin_num.set("pic", str(props.get("picture_startnum", 1)))
    begin_num.set("tbl", str(props.get("table_startnum", 1)))
    begin_num.set("equation", str(props.get("math_startnum", 1)))

    # refList
    ref_list = sub(head, "hh", "refList")

    _build_fontfaces(ref_list, reader)
    _build_border_fills(ref_list, reader)
    _build_char_properties(ref_list, reader)
    _build_tab_properties(ref_list, reader)
    _build_numberings(ref_list, reader)
    _build_para_properties(ref_list, reader)
    _build_styles(ref_list, reader)

    # compatibleDocument
    compat = reader.get_compatible_document()
    target = compat.get("target", 0) if compat else 0
    target_str = vm.COMPATIBLE_TARGET_MAP.get(target, "HWP201X")
    compat_elem = sub(head, "hh", "compatibleDocument")
    compat_elem.set("targetProgram", target_str)
    sub(compat_elem, "hh", "layoutCompatibility")

    # docOption
    doc_option = sub(head, "hh", "docOption")
    link_info = sub(doc_option, "hh", "linkinfo")
    link_info.set("path", "")
    link_info.set("pageInherit", "0")
    link_info.set("footnoteInherit", "0")

    # trackchangeConfig
    track = sub(head, "hh", "trackchageConfig")
    track.set("flags", "56")

    return head


def _build_fontfaces(ref_list, reader):
    """Build fontfaces section from HWPTAG_FACE_NAME models."""
    id_mappings = reader.get_id_mappings()
    face_names = reader.get_face_names()

    # Font counts per language group
    lang_counts = [
        id_mappings.get("ko_fonts", 0),
        id_mappings.get("en_fonts", 0),
        id_mappings.get("cn_fonts", 0),
        id_mappings.get("jp_fonts", 0),
        id_mappings.get("other_fonts", 0),
        id_mappings.get("symbol_fonts", 0),
        id_mappings.get("user_fonts", 0),
    ]

    fontfaces = sub(ref_list, "hh", "fontfaces")
    fontfaces.set("itemCnt", str(len(vm.FONT_LANG_NAMES)))

    idx = 0
    for lang_idx, lang_name in enumerate(vm.FONT_LANG_NAMES):
        count = lang_counts[lang_idx] if lang_idx < len(lang_counts) else 0
        ff = sub(fontfaces, "hh", "fontface")
        ff.set("lang", lang_name)
        ff.set("fontCnt", str(count))

        for font_idx in range(count):
            if idx >= len(face_names):
                break
            fn = face_names[idx]
            idx += 1

            font_elem = sub(ff, "hh", "font")
            font_elem.set("id", str(font_idx))
            font_elem.set("face", fn.get("name", ""))
            font_elem.set("type", "TTF")
            font_elem.set("isEmbedded", "0")

            # typeInfo from panose1
            panose = fn.get("panose1", {})
            if panose:
                ti = sub(font_elem, "hh", "typeInfo")
                family = panose.get("family_type", 0)
                ti.set("familyType", vm.FAMILY_TYPE_MAP.get(family, "FCAT_UNKNOWN"))
                ti.set("weight", str(panose.get("weight", 0)))
                ti.set("proportion", str(panose.get("proportion", 0)))
                ti.set("contrast", str(panose.get("contrast", 0)))
                ti.set("strokeVariation", str(panose.get("stroke_variation", 0)))
                ti.set("armStyle", str(panose.get("arm_style", 0)))
                ti.set("letterform", str(panose.get("letterform", 0)))
                ti.set("midline", str(panose.get("midline", 0)))
                ti.set("xHeight", str(panose.get("x_height", 0)))


def _build_border_fills(ref_list, reader):
    """Build borderFills section."""
    border_fills = reader.get_border_fills()
    bfs = sub(ref_list, "hh", "borderFills")
    bfs.set("itemCnt", str(len(border_fills)))

    for i, bf in enumerate(border_fills):
        bf_elem = sub(bfs, "hh", "borderFill")
        bf_elem.set("id", str(i + 1))

        bflags = bf.get("borderflags", 0)
        bf_elem.set("threeD", str((bflags >> 0) & 1))
        bf_elem.set("shadow", str((bflags >> 1) & 1))
        bf_elem.set("centerLine", "NONE")
        bf_elem.set("breakCellSeparateLine", str((bflags >> 2) & 1))

        # slash/backSlash
        slash = sub(bf_elem, "hh", "slash")
        slash.set("type", "NONE")
        slash.set("Crooked", "0")
        slash.set("isCounter", "0")
        bslash = sub(bf_elem, "hh", "backSlash")
        bslash.set("type", "NONE")
        bslash.set("Crooked", "0")
        bslash.set("isCounter", "0")

        # Borders
        for side in ["left", "right", "top", "bottom"]:
            border_data = bf.get(side, {})
            bi = vm.border_info(border_data)
            be = sub(bf_elem, "hh", f"{side}Border")
            be.set("type", bi["type"])
            be.set("width", bi["width"])
            be.set("color", bi["color"])

        # Diagonal
        diag_data = bf.get("diagonal", {})
        di = vm.border_info(diag_data)
        de = sub(bf_elem, "hh", "diagonal")
        de.set("type", di["type"])
        de.set("width", di["width"])
        de.set("color", di["color"])

        # Fill
        fill_flags = bf.get("fillflags", 0)
        if fill_flags:
            _build_fill_brush(bf_elem, bf)


def _build_fill_brush(parent, bf):
    """Build fillBrush element for a borderFill."""
    fb = sub(parent, "hc", "fillBrush")
    fill_flags = bf.get("fillflags", 0)

    # Solid color fill (bit 0)
    if fill_flags & 0x01:
        cp = bf.get("fill_colorpattern", {})
        bg_color = cp.get("background_color", -1)
        pat_color = cp.get("pattern_color", -1)
        wb = sub(fb, "hc", "winBrush")
        wb.set("faceColor", vm.color_from_int(bg_color))
        wb.set("hatchColor", vm.color_from_int(pat_color))
        wb.set("alpha", "0")

    # Gradation fill (bit 2)
    if fill_flags & 0x04:
        grad = bf.get("fill_gradation", {})
        if grad:
            ge = sub(fb, "hc", "gradation")
            grad_type = grad.get("type", 1)
            ge.set("type", vm.GRADATION_TYPE_MAP.get(grad_type, "LINEAR"))
            ge.set("angle", str(grad.get("shear", 0)))
            center = grad.get("center", {})
            ge.set("centerX", str(center.get("x", 50)))
            ge.set("centerY", str(center.get("y", 50)))
            ge.set("step", str(grad.get("blur", 50)))
            ge.set("colorNum", str(len(grad.get("colors", []))))
            # Color stops
            colors = grad.get("colors", [])
            for ci, c in enumerate(colors):
                cs_elem = sub(ge, "hc", "color")
                cs_elem.set("value", vm.color_from_int(c))


def _build_char_properties(ref_list, reader):
    """Build charProperties section."""
    char_shapes = reader.get_char_shapes()
    cps = sub(ref_list, "hh", "charProperties")
    cps.set("itemCnt", str(len(char_shapes)))

    for i, cs in enumerate(char_shapes):
        cp = sub(cps, "hh", "charPr")
        cp.set("id", str(i))

        height = cs.get("basesize", 1000)
        cp.set("height", str(height))
        cp.set("textColor", vm.color_from_int(cs.get("text_color", 0)))
        cp.set("shadeColor", vm.color_from_int(cs.get("shade_color", -1)))

        flags = vm.extract_charshape_flags(cs.get("charshapeflags", 0))
        cp.set("useFontSpace", "1" if flags["use_font_space"] else "0")
        cp.set("useKerning", "1" if flags["use_kerning"] else "0")
        cp.set("symMark", vm.SYM_MARK_MAP.get(flags["sym_mark"], "NONE"))
        cp.set("borderFillIDRef", str(cs.get("borderfill_id", 0) + 1 if "borderfill_id" in cs else 2))

        if flags["bold"]:
            cp.set("bold", "1")
        if flags["italic"]:
            cp.set("italic", "1")
        if flags["emboss"]:
            cp.set("emboss", "1")
        if flags["engrave"]:
            cp.set("engrave", "1")
        if flags["superscript"]:
            cp.set("superscript", "1")
        if flags["subscript"]:
            cp.set("subscript", "1")

        # fontRef
        font_face = cs.get("font_face", {})
        fr = sub(cp, "hh", "fontRef")
        for attr_name, key in zip(vm.FONT_LANG_ATTR_NAMES, vm.FONT_LANG_KEYS):
            fr.set(attr_name, str(font_face.get(key, 0)))

        # ratio
        ratio = cs.get("letter_width_expansion", {})
        r = sub(cp, "hh", "ratio")
        for attr_name, key in zip(vm.FONT_LANG_ATTR_NAMES, vm.FONT_LANG_KEYS):
            r.set(attr_name, str(ratio.get(key, 100)))

        # spacing
        spacing = cs.get("letter_spacing", {})
        s = sub(cp, "hh", "spacing")
        for attr_name, key in zip(vm.FONT_LANG_ATTR_NAMES, vm.FONT_LANG_KEYS):
            s.set(attr_name, str(spacing.get(key, 0)))

        # relSz
        rel_sz = cs.get("relative_size", {})
        rs = sub(cp, "hh", "relSz")
        for attr_name, key in zip(vm.FONT_LANG_ATTR_NAMES, vm.FONT_LANG_KEYS):
            rs.set(attr_name, str(rel_sz.get(key, 100)))

        # offset
        position = cs.get("position", {})
        off = sub(cp, "hh", "offset")
        for attr_name, key in zip(vm.FONT_LANG_ATTR_NAMES, vm.FONT_LANG_KEYS):
            off.set(attr_name, str(position.get(key, 0)))

        # underline
        ul = sub(cp, "hh", "underline")
        ul.set("type", vm.UNDERLINE_TYPE_MAP.get(flags["underline_type"], "NONE"))
        ul.set("shape", vm.LINE_SHAPE_MAP.get(flags["underline_shape"], "SOLID"))
        ul.set("color", vm.color_from_int(cs.get("underline_color", 0)))

        # strikeout
        so = sub(cp, "hh", "strikeout")
        so.set("shape", vm.STRIKEOUT_TYPE_MAP.get(flags["strikeout_type"], "NONE"))
        so.set("color", vm.color_from_int(cs.get("text_color", 0)))

        # outline
        ol = sub(cp, "hh", "outline")
        ol.set("type", vm.OUTLINE_TYPE_MAP.get(flags["outline_type"], "NONE"))

        # shadow
        sh = sub(cp, "hh", "shadow")
        sh.set("type", vm.SHADOW_TYPE_MAP.get(flags["shadow_type"], "NONE"))
        sh.set("color", vm.color_from_int(cs.get("shadow_color", 12632256)))
        shadow_space = cs.get("shadow_space", {"x": 10, "y": 10})
        sh.set("offsetX", str(shadow_space.get("x", 10)))
        sh.set("offsetY", str(shadow_space.get("y", 10)))


def _build_tab_properties(ref_list, reader):
    """Build tabProperties section."""
    tab_defs = reader.get_tab_defs()
    tps = sub(ref_list, "hh", "tabProperties")
    tps.set("itemCnt", str(len(tab_defs)))

    for i, td in enumerate(tab_defs):
        tp = sub(tps, "hh", "tabPr")
        tp.set("id", str(i))
        flags = td.get("flags", 0)
        tp.set("autoTabLeft", str(flags & 1))
        tp.set("autoTabRight", str((flags >> 1) & 1))

        # Tab items
        tab_items = td.get("tabs", [])
        for item in tab_items:
            ti = sub(tp, "hh", "tabItem")
            ti.set("pos", str(item.get("pos", 0)))
            ti.set("type", str(item.get("type", 0)))
            ti.set("leader", str(item.get("leader", 0)))


def _build_numberings(ref_list, reader):
    """Build numberings section."""
    numberings = reader.get_numberings()
    ns_elem = sub(ref_list, "hh", "numberings")
    ns_elem.set("itemCnt", str(len(numberings)))

    for i, num in enumerate(numberings):
        n = sub(ns_elem, "hh", "numbering")
        n.set("id", str(i + 1))
        n.set("start", str(num.get("start", 0)))

        # paraHead items
        levels = num.get("levels", [])
        for lvl_idx, lvl in enumerate(levels):
            ph = sub(n, "hh", "paraHead")
            ph.set("start", str(lvl.get("start", 1)))
            ph.set("level", str(lvl_idx + 1))
            ph.set("align", vm.HALIGN_MAP.get(lvl.get("align", 1), "LEFT"))
            ph.set("useInstWidth", str(lvl.get("use_inst_width", 1)))
            ph.set("autoIndent", str(lvl.get("auto_indent", 1)))
            ph.set("widthAdjust", str(lvl.get("width_adjust", 0)))
            ph.set("textOffsetType", vm.TEXT_OFFSET_TYPE_MAP.get(lvl.get("text_offset_type", 0), "PERCENT"))
            ph.set("textOffset", str(lvl.get("text_offset", 50)))
            ph.set("numFormat", vm.NUM_FORMAT_MAP.get(lvl.get("num_format", 0), "DIGIT"))
            ph.set("charPrIDRef", str(lvl.get("charshape_id", 4294967295)))
            ph.set("checkable", str(lvl.get("checkable", 0)))
            ph.text = lvl.get("format_string", "")


def _build_para_properties(ref_list, reader):
    """Build paraProperties section."""
    para_shapes = reader.get_para_shapes()
    pps = sub(ref_list, "hh", "paraProperties")
    pps.set("itemCnt", str(len(para_shapes)))

    for i, ps in enumerate(para_shapes):
        pp = sub(pps, "hh", "paraPr")
        pp.set("id", str(i))

        flags = vm.extract_parashape_flags(ps.get("parashapeflags", 0))
        flags2 = vm.extract_parashape_flags2(ps.get("flags2", 0))

        pp.set("tabPrIDRef", str(ps.get("tabdef_id", 0)))
        pp.set("condense", str(flags.get("condense", 0)))
        pp.set("fontLineHeight", "1" if flags["font_line_height"] else "0")
        pp.set("snapToGrid", "1" if flags["snap_to_grid"] else "0")
        pp.set("suppressLineNumbers", "1" if flags2["suppress_line_numbers"] else "0")
        pp.set("checked", "0")

        # align
        align = sub(pp, "hh", "align")
        align.set("horizontal", vm.HALIGN_MAP.get(flags["halign"], "JUSTIFY"))
        align.set("vertical", vm.VALIGN_MAP.get(flags["valign"], "BASELINE"))

        # heading
        heading = sub(pp, "hh", "heading")
        heading.set("type", vm.HEADING_TYPE_MAP.get(flags["heading_type"], "NONE"))
        heading.set("idRef", str(ps.get("numbering_bullet_id", 0)))
        heading.set("level", str(flags["heading_level"]))

        # breakSetting
        brk = sub(pp, "hh", "breakSetting")
        brk.set("breakLatinWord", vm.BREAK_LATIN_WORD_MAP.get(flags["break_latin_word"], "KEEP_WORD"))
        brk.set("breakNonLatinWord", "BREAK_WORD" if flags["break_non_latin_word"] else "KEEP_WORD")
        brk.set("widowOrphan", "1" if flags["widow_orphan"] else "0")
        brk.set("keepWithNext", "1" if flags["keep_with_next"] else "0")
        brk.set("keepLines", "1" if flags["keep_lines"] else "0")
        brk.set("pageBreakBefore", "1" if flags["page_break_before"] else "0")
        brk.set("lineWrap", vm.LINE_WRAP_MAP.get(flags["line_wrap"], "BREAK"))

        # autoSpacing
        auto_sp = sub(pp, "hh", "autoSpacing")
        auto_sp.set("eAsianEng", "1" if flags["auto_spacing_ea_eng"] else "0")
        auto_sp.set("eAsianNum", "1" if flags["auto_spacing_ea_num"] else "0")

        # margin and lineSpacing (using hp:switch for hwpunitchar)
        _build_para_margin_linespacing(pp, ps)

        # border
        border = sub(pp, "hh", "border")
        border.set("borderFillIDRef", str(ps.get("borderfill_id", 1)))
        border.set("offsetLeft", str(ps.get("border_left", 0)))
        border.set("offsetRight", str(ps.get("border_right", 0)))
        border.set("offsetTop", str(ps.get("border_top", 0)))
        border.set("offsetBottom", str(ps.get("border_bottom", 0)))
        border.set("connect", "0")
        border.set("ignoreMargin", "0")


def _build_para_margin_linespacing(pp, ps):
    """Build margin and lineSpacing with hp:switch structure."""
    flags = ps.get("parashapeflags", 0)
    ls_type = vm.extract_linespacing_type(flags)
    ls_type_str = vm.LINESPACING_TYPE_MAP.get(ls_type, "PERCENT")

    # margin values: HWP stores doubled values for some
    indent = ps.get("indent", 0)
    margin_left = ps.get("doubled_margin_left", 0)
    margin_right = ps.get("doubled_margin_right", 0)
    margin_top = ps.get("doubled_margin_top", 0)
    margin_bottom = ps.get("doubled_margin_bottom", 0)
    linespacing = ps.get("linespacing", 160)

    # hp:switch
    switch = sub(pp, "hp", "switch")
    case = sub(switch, "hp", "case")
    case.set(make_tag("hp", "required-namespace"), "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar")

    # Case: half values (HwpUnitChar) - pyhwp returns doubled values, so halve them
    margin_case = sub(case, "hh", "margin")
    sub(margin_case, "hc", "intent", {"value": str(indent // 2), "unit": "HWPUNIT"})
    sub(margin_case, "hc", "left", {"value": str(margin_left // 2), "unit": "HWPUNIT"})
    sub(margin_case, "hc", "right", {"value": str(margin_right // 2), "unit": "HWPUNIT"})
    sub(margin_case, "hc", "prev", {"value": str(margin_top // 2), "unit": "HWPUNIT"})
    sub(margin_case, "hc", "next", {"value": str(margin_bottom // 2), "unit": "HWPUNIT"})
    ls_case = sub(case, "hh", "lineSpacing")
    ls_case.set("type", ls_type_str)
    ls_case.set("value", str(linespacing))
    ls_case.set("unit", "HWPUNIT")

    # Default: use raw doubled values as-is
    default = sub(switch, "hp", "default")
    margin_def = sub(default, "hh", "margin")
    sub(margin_def, "hc", "intent", {"value": str(indent), "unit": "HWPUNIT"})
    sub(margin_def, "hc", "left", {"value": str(margin_left), "unit": "HWPUNIT"})
    sub(margin_def, "hc", "right", {"value": str(margin_right), "unit": "HWPUNIT"})
    sub(margin_def, "hc", "prev", {"value": str(margin_top), "unit": "HWPUNIT"})
    sub(margin_def, "hc", "next", {"value": str(margin_bottom), "unit": "HWPUNIT"})
    ls_def = sub(default, "hh", "lineSpacing")
    ls_def.set("type", ls_type_str)
    ls_def.set("value", str(linespacing))
    ls_def.set("unit", "HWPUNIT")


def _build_styles(ref_list, reader):
    """Build styles section."""
    styles = reader.get_styles()
    ss = sub(ref_list, "hh", "styles")
    ss.set("itemCnt", str(len(styles)))

    for i, st in enumerate(styles):
        s = sub(ss, "hh", "style")
        s.set("id", str(i))
        flags = st.get("flags", 0)
        style_type = (flags >> 0) & 0x07
        s.set("type", vm.STYLE_TYPE_MAP.get(style_type, "PARA"))
        s.set("name", st.get("local_name", ""))
        s.set("engName", st.get("name", ""))
        s.set("paraPrIDRef", str(st.get("parashape_id", 0)))
        s.set("charPrIDRef", str(st.get("charshape_id", 0)))
        s.set("nextStyleIDRef", str(st.get("next_style_id", 0)))
        s.set("langID", str(st.get("lang_id", 1042)))
        s.set("lockForm", "0")
