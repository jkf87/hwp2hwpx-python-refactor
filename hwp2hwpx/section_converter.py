"""Convert HWP BodyText sections to HWPX section XML files."""

import struct
from .xml_builder import root_element, sub, make_tag
from . import value_maps as vm


def _arrow_style(val):
    """Convert arrow head/tail style integer to HWPX string."""
    _ARROW_MAP = {
        0: "NORMAL",
        1: "ARROW",
        2: "SPEAR",
        3: "CONCAVE_ARROW",
        4: "EMPTY_DIAMOND",
        5: "EMPTY_CIRCLE",
        6: "EMPTY_BOX",
        7: "FILLED_DIAMOND",
        8: "FILLED_CIRCLE",
        9: "FILLED_BOX",
    }
    return _ARROW_MAP.get(val, "NORMAL")


def _compute_final_dimensions(sc_content):
    """Compute final shape dimensions by applying all scalerotation transforms.

    HWP shapes have initial_width/height and a chain of scalerotation transforms.
    The SHAPE_COMPONENT width/height only captures the first level.
    We need to apply ALL transforms to get the true rendered dimensions.
    """
    if not sc_content:
        return None, None
    scalerotations = sc_content.get("scalerotations", [])
    if not scalerotations:
        return sc_content.get("width"), sc_content.get("height")

    w = sc_content.get("initial_width", 0)
    h = sc_content.get("initial_height", 0)
    if w == 0 and h == 0:
        return sc_content.get("width"), sc_content.get("height")

    for sr in scalerotations:
        s = sr.get("scaler", {})
        a = s.get("a", 1.0)
        b = s.get("b", 0.0)
        c = s.get("c", 0.0)
        d = s.get("d", 1.0)
        new_w = abs(a * w + c * h)
        new_h = abs(b * w + d * h)
        w, h = new_w, new_h

    return int(round(w)), int(round(h))


def _transform_point(x, y, sc_content):
    """Transform a point through the scalerotation chain (scaling only, no translation).

    For HWPX, line endpoints are in the shape's local coordinate space.
    We apply only the scaling/rotation part of each transform, not the translation,
    since translation positions the shape within its parent.
    """
    if not sc_content:
        return x, y
    scalerotations = sc_content.get("scalerotations", [])
    if not scalerotations:
        return x, y

    fx, fy = float(x), float(y)
    for sr in scalerotations:
        s = sr.get("scaler", {})
        a = s.get("a", 1.0)
        b = s.get("b", 0.0)
        c = s.get("c", 0.0)
        d = s.get("d", 1.0)
        # Exclude translation (e, f) - positioning is handled by pos element
        new_x = a * fx + c * fy
        new_y = b * fx + d * fy
        fx, fy = new_x, new_y

    return int(round(fx)), int(round(fy))


def build_section_xml(reader, section_idx, bindata_id_map=None):
    """Build section XML element from HWP section models."""
    models = reader.get_section_models(section_idx)
    sec = root_element("hs", "sec")

    ctx = ConversionContext(reader, models, bindata_id_map)
    ctx.process_section(sec)
    return sec


class ConversionContext:
    """Manages state during section model conversion."""

    def __init__(self, reader, models, bindata_id_map=None):
        self.reader = reader
        self.models = models
        self.pos = 0
        # Maps 1-based BIN_DATA list index → "imageN" label for entries
        # with actual binary data.  Used to emit valid binaryItemIDRef.
        self._bindata_id_map = bindata_id_map or {}

    def advance(self):
        self.pos += 1

    def process_section(self, sec):
        """Process all models for a section."""
        while self.pos < len(self.models):
            model = self.models[self.pos]
            tagname = model.get("tagname", "")
            level = model.get("level", 0)

            if tagname == "HWPTAG_PARA_HEADER" and level == 0:
                self._process_paragraph(sec)
            else:
                self.advance()

    def _find_end_of_children(self, start_pos, parent_level):
        """Find the position after all children of a model at parent_level."""
        pos = start_pos
        while pos < len(self.models):
            if self.models[pos].get("level", 0) <= parent_level:
                break
            pos += 1
        return pos

    def _process_paragraph(self, parent):
        """Process a paragraph and its children.

        Key insight: controls (section def, column def, tables) appear as both
        control chars in PARA_TEXT and as CTRL_HEADER models. The CTRL_HEADER
        models with their children need to be placed inline within the paragraph's
        run elements, at the position of their control char.
        """
        model = self.models[self.pos]
        content = model.get("content", {})
        para_level = model.get("level", 0)

        p = sub(parent, "hp", "p")
        p.set("id", str(content.get("instance_id", 0)))
        p.set("paraPrIDRef", str(content.get("parashape_id", 0)))
        p.set("styleIDRef", str(content.get("style_id", 0)))
        p.set("pageBreak", "0")
        p.set("columnBreak", "0")
        p.set("merged", "0")

        self.advance()

        # Scan children: collect text/charshape/lineseg, and record ctrl positions
        text_chunks = []
        char_shapes = []
        line_segs = []
        ctrl_start_positions = []  # (chid, model_start_pos, model_end_pos)

        scan_pos = self.pos
        while scan_pos < len(self.models):
            child = self.models[scan_pos]
            child_tag = child.get("tagname", "")
            child_level = child.get("level", 0)

            if child_level <= para_level:
                break

            if child_tag == "HWPTAG_PARA_TEXT" and child_level == para_level + 1:
                text_chunks = child.get("content", {}).get("chunks", [])
                scan_pos += 1
            elif child_tag == "HWPTAG_PARA_CHAR_SHAPE" and child_level == para_level + 1:
                char_shapes = child.get("content", {}).get("charshapes", [])
                scan_pos += 1
            elif child_tag == "HWPTAG_PARA_LINE_SEG" and child_level == para_level + 1:
                line_segs = child.get("content", {}).get("linesegs", [])
                scan_pos += 1
            elif child_tag == "HWPTAG_CTRL_HEADER" and child_level == para_level + 1:
                chid = child.get("content", {}).get("chid", "")
                ctrl_start = scan_pos
                scan_pos += 1
                # Find end of this control's children
                while scan_pos < len(self.models):
                    if self.models[scan_pos].get("level", 0) <= para_level + 1:
                        break
                    scan_pos += 1
                ctrl_start_positions.append((chid, ctrl_start, scan_pos))
            else:
                scan_pos += 1

        # Set self.pos to after all paragraph children
        self.pos = scan_pos

        # Build runs with inline controls
        self._build_runs_with_controls(p, text_chunks, char_shapes, ctrl_start_positions, para_level)

        # Build linesegarray
        if line_segs:
            self._build_line_segs(p, line_segs)

    def _build_runs_with_controls(self, p, text_chunks, char_shapes, ctrl_positions, para_level):
        """Build run elements with inline controls."""
        if not text_chunks and not ctrl_positions:
            run = sub(p, "hp", "run")
            run.set("charPrIDRef", str(char_shapes[0][1] if char_shapes else 0))
            return

        def get_charshape_id(pos):
            result = 0
            for cs_pos, cs_id in char_shapes:
                if cs_pos <= pos:
                    result = cs_id
                else:
                    break
            return result

        current_run = None
        current_cs_id = None
        text_buffer = []
        ctrl_idx = 0

        def flush_text():
            nonlocal text_buffer
            if current_run is not None and text_buffer:
                t = sub(current_run, "hp", "t")
                t.text = "".join(text_buffer)
                text_buffer = []

        def ensure_run(cs_id):
            nonlocal current_run, current_cs_id
            if cs_id != current_cs_id or current_run is None:
                flush_text()
                current_cs_id = cs_id
                current_run = sub(p, "hp", "run")
                current_run.set("charPrIDRef", str(cs_id))

        for chunk in text_chunks:
            if len(chunk) < 2:
                continue
            pos_range, value = chunk[0], chunk[1]
            char_pos = pos_range[0] if isinstance(pos_range, tuple) else 0

            if isinstance(value, str):
                # Split text at char shape boundaries
                text = value
                text_start = char_pos
                idx = 0
                while idx < len(text):
                    cs_id = get_charshape_id(text_start + idx)
                    ensure_run(cs_id)
                    # Find how far this char shape extends
                    seg_start = idx
                    while idx < len(text):
                        next_cs = get_charshape_id(text_start + idx)
                        if next_cs != cs_id:
                            break
                        idx += 1
                    text_buffer.append(text[seg_start:idx])
            elif isinstance(value, dict):
                code = value.get("code", 0)
                cs_id = get_charshape_id(char_pos)
                ensure_run(cs_id)

                if code == 13:
                    flush_text()
                elif code == 10:
                    flush_text()
                    sub(current_run, "hp", "lineBreak")
                elif code == 9:
                    flush_text()
                    sub(current_run, "hp", "tab")
                elif code == 24:
                    text_buffer.append("-")
                elif code == 30:
                    text_buffer.append("\u00A0")
                elif code == 31:
                    text_buffer.append("\u3000")
                elif code == 2:
                    # Section/Column control - insert inline
                    flush_text()
                    if ctrl_idx < len(ctrl_positions):
                        chid, start, end = ctrl_positions[ctrl_idx]
                        ctrl_idx += 1
                        ctrl_model = self.models[start]
                        if chid == "secd":
                            self._build_section_def_inline(current_run, ctrl_model, start + 1, end)
                        elif chid == "cold":
                            self._build_column_def_inline(current_run, ctrl_model, start + 1, end)
                elif code == 11:
                    # Extended control
                    flush_text()
                    if ctrl_idx < len(ctrl_positions):
                        chid, start, end = ctrl_positions[ctrl_idx]
                        ctrl_idx += 1
                        ctrl_model = self.models[start]
                        if chid.strip() == "tbl":
                            self._build_table_inline(current_run, ctrl_model, start + 1, end)
                        elif chid.strip() == "gso":
                            self._build_gso_inline(current_run, ctrl_model, start + 1, end)
                        elif chid.strip() == "eqed":
                            self._build_equation_inline(current_run, ctrl_model, start + 1, end)
                elif code == 16:
                    # Header/Footer control
                    flush_text()
                    if ctrl_idx < len(ctrl_positions):
                        chid, start, end = ctrl_positions[ctrl_idx]
                        ctrl_idx += 1
                        ctrl_model = self.models[start]
                        if chid == "head":
                            self._build_header_footer_inline(
                                current_run, ctrl_model, start + 1, end, "header"
                            )
                        elif chid == "foot":
                            self._build_header_footer_inline(
                                current_run, ctrl_model, start + 1, end, "footer"
                            )
                elif code == 17:
                    # Footnote/Endnote reference
                    flush_text()
                    if ctrl_idx < len(ctrl_positions):
                        chid, start, end = ctrl_positions[ctrl_idx]
                        ctrl_idx += 1
                        ctrl_model = self.models[start]
                        if chid.strip() == "fn":
                            self._build_footnote_endnote_inline(
                                current_run, ctrl_model, start + 1, end, "footNote", "FOOTNOTE"
                            )
                        elif chid.strip() == "en":
                            self._build_footnote_endnote_inline(
                                current_run, ctrl_model, start + 1, end, "endNote", "ENDNOTE"
                            )
                elif code == 18:
                    # Auto number (atno) - used inside footnote/endnote body
                    flush_text()
                    if ctrl_idx < len(ctrl_positions):
                        chid, start, end = ctrl_positions[ctrl_idx]
                        ctrl_idx += 1
                        ctrl_model = self.models[start]
                        self._build_auto_num_inline(current_run, ctrl_model)
                elif code == 21:
                    # Autonomous inline control (pgnp, pghd, nwno, tcps)
                    flush_text()
                    if ctrl_idx < len(ctrl_positions):
                        chid, start, end = ctrl_positions[ctrl_idx]
                        ctrl_idx += 1
                        ctrl_model = self.models[start]
                        self._build_autonomous_ctrl_inline(current_run, chid.strip(), ctrl_model)
                elif code == 23:
                    # Table cell paragraph shape (tcps) - consume ctrl
                    flush_text()
                    if ctrl_idx < len(ctrl_positions):
                        ctrl_idx += 1
                elif code == 3:
                    # Field begin - has a CTRL_HEADER
                    flush_text()
                    if ctrl_idx < len(ctrl_positions):
                        chid, start, end = ctrl_positions[ctrl_idx]
                        ctrl_idx += 1
                        ctrl_model = self.models[start]
                        self._build_field_begin_inline(current_run, ctrl_model)
                elif code == 4:
                    # Field end - NO CTRL_HEADER, info is in chunk param bytes
                    flush_text()
                    self._build_field_end_inline(current_run, value)

        flush_text()

        # Add empty <hp:t/> if the last run has control but no trailing text
        if current_run is not None:
            has_text = any(
                child.tag.endswith("}t") for child in current_run
            )
            # Check if last child is a control (tbl, secPr, ctrl) - need trailing empty t
            children = list(current_run)
            if children:
                last_tag = children[-1].tag.split("}")[-1] if "}" in children[-1].tag else children[-1].tag
                if last_tag in ("tbl", "secPr", "ctrl") or not has_text:
                    t = sub(current_run, "hp", "t")

    def _build_line_segs(self, p, line_segs):
        """Build linesegarray element."""
        lsa = sub(p, "hp", "linesegarray")
        for ls in line_segs:
            lse = sub(lsa, "hp", "lineseg")
            lse.set("textpos", str(ls.get("chpos", 0)))
            lse.set("vertpos", str(ls.get("y", 0)))
            lse.set("vertsize", str(ls.get("height", 1000)))
            lse.set("textheight", str(ls.get("height_text", 1000)))
            lse.set("baseline", str(ls.get("height_baseline", 850)))
            lse.set("spacing", str(ls.get("space_below", 600)))
            lse.set("horzpos", str(ls.get("x", 0)))
            lse.set("horzsize", str(ls.get("width", 0)))
            lse.set("flags", str(ls.get("lineseg_flags", 393216)))

    # ---------- Inline control builders ----------

    def _build_section_def_inline(self, run, ctrl_model, children_start, children_end):
        """Build secPr inside a run element."""
        content = ctrl_model.get("content", {})

        sec_pr = sub(run, "hp", "secPr")
        sec_pr.set("id", "")
        sec_pr.set("textDirection", "HORIZONTAL")
        sec_pr.set("spaceColumns", str(content.get("columnspacing", 1134)))
        sec_pr.set("tabStop", str(content.get("defaultTabStops", 8000)))
        sec_pr.set("tabStopVal", str(content.get("defaultTabStops", 8000) // 2))
        sec_pr.set("tabStopUnit", "HWPUNIT")
        sec_pr.set("outlineShapeIDRef", str(content.get("numbering_shape_id", 1)))
        sec_pr.set("memoShapeIDRef", "0")
        sec_pr.set("textVerticalWidthHead", "0")
        sec_pr.set("masterPageCnt", "0")

        # Grid
        grid = sub(sec_pr, "hp", "grid")
        grid.set("lineGrid", str(content.get("grid_vertical", 0)))
        grid.set("charGrid", str(content.get("grid_horizontal", 0)))
        grid.set("wonggojiFormat", "0")

        # Process child models
        page_def = None
        footnote_shapes = []
        page_border_fills = []

        for i in range(children_start, children_end):
            child = self.models[i]
            child_tag = child.get("tagname", "")
            child_content = child.get("content", {})
            if child_tag == "HWPTAG_PAGE_DEF":
                page_def = child_content
            elif child_tag == "HWPTAG_FOOTNOTE_SHAPE":
                footnote_shapes.append(child_content)
            elif child_tag == "HWPTAG_PAGE_BORDER_FILL":
                page_border_fills.append(child_content)

        # startNum
        start_num = sub(sec_pr, "hp", "startNum")
        start_num.set("pageStartsOn", "BOTH")
        start_num.set("page", str(content.get("starting_pagenum", 0)))
        start_num.set("pic", str(content.get("starting_picturenum", 0)))
        start_num.set("tbl", str(content.get("starting_tablenum", 0)))
        start_num.set("equation", str(content.get("starting_equationnum", 0)))

        # visibility
        vis = sub(sec_pr, "hp", "visibility")
        vis.set("hideFirstHeader", "0")
        vis.set("hideFirstFooter", "0")
        vis.set("hideFirstMasterPage", "0")
        vis.set("border", "SHOW_ALL")
        vis.set("fill", "SHOW_ALL")
        vis.set("hideFirstPageNum", "0")
        vis.set("hideFirstEmptyLine", "0")
        vis.set("showLineNumber", "0")

        # lineNumberShape
        lns = sub(sec_pr, "hp", "lineNumberShape")
        lns.set("restartType", "0")
        lns.set("countBy", "0")
        lns.set("distance", "0")
        lns.set("startNumber", "0")

        # pagePr
        if page_def:
            self._build_page_pr(sec_pr, page_def)

        # footNotePr, endNotePr
        if len(footnote_shapes) >= 1:
            self._build_footnote_pr(sec_pr, footnote_shapes[0], "footNotePr")
        if len(footnote_shapes) >= 2:
            self._build_footnote_pr(sec_pr, footnote_shapes[1], "endNotePr")

        # pageBorderFill
        pbf_types = ["BOTH", "EVEN", "ODD"]
        for idx, pbf in enumerate(page_border_fills):
            pbf_type = pbf_types[idx] if idx < len(pbf_types) else "BOTH"
            self._build_page_border_fill(sec_pr, pbf, pbf_type)

    def _build_column_def_inline(self, run, ctrl_model, children_start, children_end):
        """Build colPr ctrl inside a run element."""
        content = ctrl_model.get("content", {})

        ctrl = sub(run, "hp", "ctrl")
        col_pr = sub(ctrl, "hp", "colPr")
        col_pr.set("id", "")

        # Column property bit layout:
        # bits 0-1: type, bits 2-9: colCount, bits 10-11: layout,
        # bit 12: sameSz, bit 13: sameGap
        flags = content.get("flags", 0)
        col_type = flags & 0x03
        col_count = (flags >> 2) & 0xFF
        col_layout = (flags >> 10) & 0x03
        same_sz = (flags >> 12) & 0x01

        col_pr.set("type", vm.COLUMN_TYPE_MAP.get(col_type, "NEWSPAPER"))
        col_pr.set("layout", vm.COLUMN_LAYOUT_MAP.get(col_layout, "LEFT"))
        col_pr.set("colCount", str(col_count if col_count > 0 else 1))
        col_pr.set("sameSz", str(same_sz))
        col_pr.set("sameGap", str(content.get("spacing", 0)))

    def _build_header_footer_inline(self, run, ctrl_model, children_start, children_end, tag_name):
        """Build header or footer element inside a run element.

        tag_name is either 'header' or 'footer'.
        HWP structure: CTRL_HEADER(head/foot) → LIST_HEADER → paragraphs
        HWPX: <hp:ctrl><hp:header|footer><hp:subList><hp:p>...</hp:p></hp:subList></hp:header|footer></hp:ctrl>
        """
        content = ctrl_model.get("content", {})
        ctrl_level = ctrl_model.get("level", 1)

        # Flags bits 0-1: places (BOTH_PAGES=0, EVEN_PAGE=1, ODD_PAGE=2)
        flags = content.get("flags", 0)
        places = flags & 0x03
        places_map = {0: "BOTH", 1: "EVEN", 2: "ODD"}

        ctrl = sub(run, "hp", "ctrl")
        hf_elem = sub(ctrl, "hp", tag_name)
        hf_elem.set("id", "")
        hf_elem.set("applyPageType", places_map.get(places, "BOTH"))

        # Find LIST_HEADER child
        list_header = None
        para_start = children_start
        for i in range(children_start, children_end):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_LIST_HEADER" and m.get("level", 0) == ctrl_level + 1:
                list_header = m.get("content", {})
                para_start = i + 1
                break

        # Build subList
        sub_list = sub(hf_elem, "hp", "subList")
        sub_list.set("id", "")
        sub_list.set("textDirection", "HORIZONTAL")
        sub_list.set("lineWrap", "BREAK")

        if list_header:
            lf = list_header.get("listflags", 0)
            valign_val = (lf >> 5) & 0x03
            valign_map = {0: "TOP", 1: "CENTER", 2: "BOTTOM"}
            sub_list.set("vertAlign", valign_map.get(valign_val, "TOP"))
            sub_list.set("linkListIDRef", "0")
            sub_list.set("linkListNextIDRef", "0")
            sub_list.set("textWidth", str(list_header.get("width", 0)))
            sub_list.set("textHeight", str(list_header.get("height", 0)))
            sub_list.set("hasTextRef", str(list_header.get("textrefsbitmap", 0)))
            sub_list.set("hasNumRef", str(list_header.get("numberrefsbitmap", 0)))
        else:
            sub_list.set("vertAlign", "TOP")
            sub_list.set("linkListIDRef", "0")
            sub_list.set("linkListNextIDRef", "0")
            sub_list.set("textWidth", "0")
            sub_list.set("textHeight", "0")
            sub_list.set("hasTextRef", "0")
            sub_list.set("hasNumRef", "0")

        # Process paragraphs within the header/footer
        saved_pos = self.pos
        self.pos = para_start
        while self.pos < children_end:
            m = self.models[self.pos]
            if m.get("tagname") == "HWPTAG_PARA_HEADER" and m.get("level", 0) == ctrl_level + 1:
                self._process_paragraph(sub_list)
            else:
                self.pos += 1
        self.pos = saved_pos

    def _build_footnote_endnote_inline(self, run, ctrl_model, children_start, children_end, tag_name, num_type):
        """Build footNote or endNote element inside a run element.

        tag_name is either 'footNote' or 'endNote'.
        num_type is 'FOOTNOTE' or 'ENDNOTE'.
        HWP structure: CTRL_HEADER(fn/en) → LIST_HEADER → paragraphs
        HWPX: <hp:ctrl><hp:footNote|endNote number=.. suffixChar=.. instId=..>
                 <hp:subList ...><hp:p>...</hp:p></hp:subList>
               </hp:footNote|endNote></hp:ctrl>
        """
        content = ctrl_model.get("content", {})
        ctrl_level = ctrl_model.get("level", 1)

        ctrl = sub(run, "hp", "ctrl")
        note_elem = sub(ctrl, "hp", tag_name)
        note_elem.set("number", str(content.get("number", 1)))
        # suffixChar is stored in upper 16 bits of unknown0
        suffix_char = (content.get("unknown0", 0) >> 16) & 0xFFFF
        note_elem.set("suffixChar", str(suffix_char))
        # instId is (unknown3 << 16) | unknown2
        inst_id = (content.get("unknown3", 0) << 16) | content.get("unknown2", 0)
        note_elem.set("instId", str(inst_id))

        # Find LIST_HEADER child
        list_header = None
        para_start = children_start
        for i in range(children_start, children_end):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_LIST_HEADER" and m.get("level", 0) == ctrl_level + 1:
                list_header = m.get("content", {})
                para_start = i + 1
                break

        # Build subList
        sub_list = sub(note_elem, "hp", "subList")
        sub_list.set("id", "")
        sub_list.set("textDirection", "HORIZONTAL")
        sub_list.set("lineWrap", "BREAK")

        if list_header:
            lf = list_header.get("listflags", 0)
            valign_val = (lf >> 5) & 0x03
            valign_map = {0: "TOP", 1: "CENTER", 2: "BOTTOM"}
            sub_list.set("vertAlign", valign_map.get(valign_val, "TOP"))
            sub_list.set("linkListIDRef", "0")
            sub_list.set("linkListNextIDRef", "0")
            sub_list.set("textWidth", str(list_header.get("width", 0)))
            sub_list.set("textHeight", str(list_header.get("height", 0)))
            sub_list.set("hasTextRef", str(list_header.get("textrefsbitmap", 0)))
            sub_list.set("hasNumRef", str(list_header.get("numberrefsbitmap", 0)))
        else:
            sub_list.set("vertAlign", "TOP")
            sub_list.set("linkListIDRef", "0")
            sub_list.set("linkListNextIDRef", "0")
            sub_list.set("textWidth", "0")
            sub_list.set("textHeight", "0")
            sub_list.set("hasTextRef", "0")
            sub_list.set("hasNumRef", "0")

        # Process paragraphs within the footnote/endnote
        saved_pos = self.pos
        self.pos = para_start
        while self.pos < children_end:
            m = self.models[self.pos]
            if m.get("tagname") == "HWPTAG_PARA_HEADER" and m.get("level", 0) == ctrl_level + 1:
                self._process_paragraph(sub_list)
            else:
                self.pos += 1
        self.pos = saved_pos

    def _build_auto_num_inline(self, run, ctrl_model):
        """Build autoNum element inside a run element.

        HWP: CTRL_HEADER(atno) with flags, number, suffix
        HWPX: <hp:ctrl><hp:autoNum num=.. numType=..>
                 <hp:autoNumFormat type=.. suffixChar=.. .../>
               </hp:autoNum></hp:ctrl>
        """
        content = ctrl_model.get("content", {})
        flags = content.get("flags", 0)

        # flags bits 0-1: numType (1=FOOTNOTE, 2=ENDNOTE, etc.)
        num_type_val = flags & 0x03
        num_type_map = {0: "PAGE", 1: "FOOTNOTE", 2: "ENDNOTE", 3: "PICTURE", 4: "TABLE", 5: "EQUATION"}
        num_type = num_type_map.get(num_type_val, "FOOTNOTE")

        ctrl = sub(run, "hp", "ctrl")
        auto_num = sub(ctrl, "hp", "autoNum")
        auto_num.set("num", str(content.get("number", 1)))
        auto_num.set("numType", num_type)

        # autoNumFormat
        anf = sub(auto_num, "hp", "autoNumFormat")
        # Format type from bits 2-5 of flags
        fmt_type = (flags >> 2) & 0x0F
        anf.set("type", vm.NUM_FORMAT_MAP.get(fmt_type, "DIGIT"))
        anf.set("userChar", "")
        anf.set("prefixChar", "")
        suffix_val = content.get("suffix", 0)
        anf.set("suffixChar", vm.SUFFIX_CHAR_MAP.get(suffix_val, chr(suffix_val) if suffix_val > 0 else ""))
        anf.set("supscript", "0")

    def _build_autonomous_ctrl_inline(self, run, chid, ctrl_model):
        """Build autonomous inline control (pgnp, pghd, nwno, tcps) inside a run.

        HWP: CTRL_HEADER with chid='pgnp'|'pghd'|'nwno'|'tcps'
        HWPX: <hp:ctrl><hp:pageNum .../></hp:ctrl>
              <hp:ctrl><hp:pageHiding .../></hp:ctrl>
              <hp:ctrl><hp:newNum .../></hp:ctrl>
        """
        content = ctrl_model.get("content", {})

        if chid == "pgnp":
            # Page number position
            flags = content.get("flags", 0)
            # flags bits 8-11: position, bits 0-3: format type
            pos_val = (flags >> 8) & 0x0F
            pos_map = {
                0: "NONE",
                1: "TOP_LEFT", 2: "TOP_CENTER", 3: "TOP_RIGHT",
                4: "BOTTOM_LEFT", 5: "BOTTOM_CENTER", 6: "BOTTOM_RIGHT",
                7: "OUTSIDE_TOP", 8: "OUTSIDE_BOTTOM",
                9: "INSIDE_TOP", 10: "INSIDE_BOTTOM",
            }
            fmt_val = flags & 0x0F
            fmt_map = {0: "DIGIT", 1: "CIRCLE_DIGIT", 2: "ROMAN_CAPITAL",
                       3: "ROMAN_SMALL", 4: "HANGUL", 5: "HANJA",
                       6: "HANGUL_SYLLABLE", 7: "HANJA_SYLLABLE"}

            ctrl = sub(run, "hp", "ctrl")
            pg = sub(ctrl, "hp", "pageNum")
            pg.set("pos", pos_map.get(pos_val, "BOTTOM_CENTER"))
            pg.set("formatType", fmt_map.get(fmt_val, "DIGIT"))
            dash = content.get("dash", 0)
            pg.set("sideChar", chr(dash) if 0x20 <= dash < 0x7F else "-")

        elif chid == "pghd":
            # Page hiding definition
            flags = content.get("flags", 0)
            ctrl = sub(run, "hp", "ctrl")
            ph = sub(ctrl, "hp", "pageHiding")
            ph.set("hideHeader", str((flags >> 0) & 0x01))
            ph.set("hideFooter", str((flags >> 1) & 0x01))
            ph.set("hideMasterPage", str((flags >> 2) & 0x01))
            ph.set("hideBorder", str((flags >> 3) & 0x01))
            ph.set("hideFill", str((flags >> 4) & 0x01))
            ph.set("hidePageNum", str((flags >> 5) & 0x01))

        elif chid == "nwno":
            # New numbering - flags bits 0-3 encode numType
            flags = content.get("flags", 0)
            num_type_val = flags & 0x0F
            nwno_type_map = {0: "PAGE", 1: "FOOTNOTE", 2: "ENDNOTE",
                             3: "PICTURE", 4: "TABLE", 5: "EQUATION"}
            ctrl = sub(run, "hp", "ctrl")
            nn = sub(ctrl, "hp", "newNum")
            nn.set("num", str(content.get("number", 1)))
            nn.set("numType", nwno_type_map.get(num_type_val, "PAGE"))

        elif chid == "tcps":
            # Table cell paragraph shape - currently minimal support
            pass
        # else: unknown autonomous control, skip

    # ---------- Field control builders ----------

    # Field type mapping: chid prefix → HWPX type name
    _FIELD_TYPE_MAP = {
        "%clk": "CLICK_HERE",
        "%hyp": "HYPERLINK",
        "%bok": "BOOKMARK",
        "%formula": "FORMULA",
        "%summry": "SUMMARY",
        "%usr": "USERINFO",
        "%hlnk": "HYPERLINK",
        "%date": "DATE",
        "%dtfm": "DATE",
        "%path": "FILEPATH",
        "%bmk": "BOOKMARK",
        "%mmrg": "MAIL_MERGE",
        "%xref": "CROSSREF",
        "%unkn": "UNKNOWN",
    }

    def _build_field_begin_inline(self, run, ctrl_model):
        """Build fieldBegin element inside a run.

        HWP: CTRL_HEADER with chid='%clk'|'%hyp'|'%bok'|etc., flags, extra_attr, command, id
        HWPX: <hp:ctrl><hp:fieldBegin id=.. type=.. name=.. editable=.. dirty=.. zorder=.. fieldid=..>
                 <hp:parameters ...>
               </hp:fieldBegin></hp:ctrl>
        """
        content = ctrl_model.get("content", {})
        chid = content.get("chid", "").strip()
        field_type = self._FIELD_TYPE_MAP.get(chid, "UNKNOWN")

        field_id = content.get("id", 0)

        # Track field begins for LIFO matching with field ends
        if not hasattr(self, "_field_begin_stack"):
            self._field_begin_stack = []
        self._field_begin_stack.append(field_id)

        ctrl = sub(run, "hp", "ctrl")
        fb = sub(ctrl, "hp", "fieldBegin")
        fb.set("id", str(field_id))
        fb.set("type", field_type)
        fb.set("name", "")
        flags = content.get("flags", 0)
        fb.set("editable", str((flags >> 0) & 0x01))
        fb.set("dirty", str((flags >> 15) & 0x01))
        fb.set("zorder", "-1")
        # fieldid = chid bytes as big-endian int
        fieldid = int.from_bytes(chid[:4].encode("ascii", errors="replace"), "big") if len(chid) >= 4 else 0
        fb.set("fieldid", str(fieldid))

        # Build parameters from extra_attr and command
        command = content.get("command", "")
        extra_attr = content.get("extra_attr", 0)
        if command or extra_attr:
            params = sub(fb, "hp", "parameters")
            param_count = 0
            if extra_attr:
                ip = sub(params, "hp", "integerParam")
                ip.set("name", "Prop")
                ip.text = str(extra_attr)
                param_count += 1
            if command:
                sp = sub(params, "hp", "stringParam")
                sp.set("name", "Command")
                sp.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                sp.text = command
                param_count += 1
            params.set("cnt", str(param_count))
            params.set("name", "")

    def _build_field_end_inline(self, run, chunk_value):
        """Build fieldEnd element inside a run.

        Field ends have no CTRL_HEADER. They match field begins in LIFO order
        (fields nest like parentheses). We pop the most recent open field ID.
        HWPX: <hp:ctrl><hp:fieldEnd beginIDRef=.. fieldid=../></hp:ctrl>
        """
        begin_id = 0
        if hasattr(self, "_field_begin_stack") and self._field_begin_stack:
            begin_id = self._field_begin_stack.pop()

        # fieldid = chid bytes (without leading \t) → same field type id as begin
        chid = chunk_value.get("chid", "") if isinstance(chunk_value, dict) else ""
        # Field end chid starts with \t, replace with % to match begin's chid
        chid_clean = "%" + chid[1:] if chid.startswith("\t") else chid
        fieldid = int.from_bytes(chid_clean[:4].encode("ascii", errors="replace"), "big") if len(chid_clean) >= 4 else 0

        ctrl = sub(run, "hp", "ctrl")
        fe = sub(ctrl, "hp", "fieldEnd")
        fe.set("beginIDRef", str(begin_id))
        fe.set("fieldid", str(fieldid))

    # ---------- Equation builder ----------

    def _build_equation_inline(self, run, ctrl_model, children_start, children_end):
        """Build hp:equation element for an equation control.

        HWP: CTRL_HEADER(eqed) → CTRL_EQEDIT
        pyhwp doesn't parse eqed CommonControl fields or CTRL_EQEDIT content,
        so we parse the raw unparsed bytes.
        HWPX: <hp:equation version=.. baseLine=.. textColor=.. baseUnit=.. lineMode=.. font=..>
                 <hp:sz .../><hp:pos .../><hp:outMargin .../>
                 <hp:shapeComment>text</hp:shapeComment>
                 <hp:script>formula</hp:script>
               </hp:equation>
        """
        content = ctrl_model.get("content", {})
        unparsed = ctrl_model.get("unparsed", b"")

        # Parse CommonControl fields from unparsed bytes
        eq_props = self._parse_eqed_ctrl(unparsed)

        # Parse CTRL_EQEDIT payload for equation content
        eq_content = {}
        for i in range(children_start, children_end):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_CTRL_EQEDIT":
                eq_content = self._parse_eqedit(m.get("unparsed", m.get("payload", b"")))
                break

        eq = sub(run, "hp", "equation")
        eq.set("id", str(eq_props.get("instance_id", 0)))
        eq.set("zOrder", str(eq_props.get("z_order", 0)))
        eq.set("numberingType", "EQUATION")

        ctrl_flags = eq_props.get("flags", 0)
        text_wrap_type = (ctrl_flags >> 21) & 0x07
        text_flow = (ctrl_flags >> 24) & 0x03
        eq.set("textWrap", vm.TEXT_WRAP_MAP.get(text_wrap_type, "TOP_AND_BOTTOM"))
        eq.set("textFlow", vm.TEXT_FLOW_MAP.get(text_flow, "BOTH_SIDES"))
        eq.set("lock", "0")
        eq.set("dropcapstyle", "None")

        eq.set("version", eq_content.get("version", ""))
        eq.set("baseLine", str(eq_content.get("baseline", 0)))
        eq.set("textColor", vm.color_from_int(eq_content.get("text_color", 0)))
        eq.set("baseUnit", str(eq_content.get("base_unit", 1000)))
        line_mode = eq_content.get("line_mode", 1)
        eq.set("lineMode", "LINE" if line_mode else "CHAR")
        eq.set("font", eq_content.get("font", "HancomEQN"))

        # hp:sz
        sz = sub(eq, "hp", "sz")
        sz.set("width", str(eq_props.get("width", 0)))
        sz.set("widthRelTo", "COLUMN")
        sz.set("height", str(eq_props.get("height", 0)))
        sz.set("heightRelTo", "ABSOLUTE")
        sz.set("protect", "0")

        # hp:pos
        pos = sub(eq, "hp", "pos")
        pos.set("treatAsChar", str((ctrl_flags >> 0) & 0x01))
        pos.set("affectLSpacing", str((ctrl_flags >> 2) & 0x01))
        pos.set("flowWithText", str((ctrl_flags >> 17) & 0x01))
        pos.set("allowOverlap", str((ctrl_flags >> 18) & 0x01))
        pos.set("holdAnchorAndSO", "0")
        vert_rel = (ctrl_flags >> 3) & 0x03
        horz_rel = (ctrl_flags >> 8) & 0x03
        vert_align = (ctrl_flags >> 10) & 0x07
        horz_align = (ctrl_flags >> 14) & 0x07
        pos.set("vertRelTo", vm.VERT_REL_TO_MAP.get(vert_rel, "PARA"))
        pos.set("horzRelTo", vm.HORZ_REL_TO_MAP.get(horz_rel, "PARA"))
        pos.set("vertAlign", vm.VERT_ALIGN_MAP.get(vert_align, "TOP"))
        pos.set("horzAlign", vm.HORZ_ALIGN_MAP.get(horz_align, "LEFT"))
        pos.set("vertOffset", str(eq_props.get("y", 0)))
        pos.set("horzOffset", str(eq_props.get("x", 0)))

        # hp:outMargin
        om = sub(eq, "hp", "outMargin")
        om.set("left", str(eq_props.get("margin_left", 0)))
        om.set("right", str(eq_props.get("margin_right", 0)))
        om.set("top", str(eq_props.get("margin_top", 0)))
        om.set("bottom", str(eq_props.get("margin_bottom", 0)))

        # hp:shapeComment
        desc = eq_props.get("description", "")
        sc = sub(eq, "hp", "shapeComment")
        if desc:
            sc.text = desc

        # hp:script (equation formula)
        script = eq_content.get("script", "")
        if script:
            scr = sub(eq, "hp", "script")
            scr.text = script

    @staticmethod
    def _parse_eqed_ctrl(unparsed):
        """Parse CommonControl fields for eqed from unparsed bytes."""
        result = {}
        if not unparsed or len(unparsed) < 36:
            return result
        off = 0
        result["flags"] = struct.unpack_from("<I", unparsed, off)[0]; off += 4
        result["y"] = struct.unpack_from("<i", unparsed, off)[0]; off += 4
        result["x"] = struct.unpack_from("<i", unparsed, off)[0]; off += 4
        result["width"] = struct.unpack_from("<I", unparsed, off)[0]; off += 4
        result["height"] = struct.unpack_from("<I", unparsed, off)[0]; off += 4
        result["z_order"] = struct.unpack_from("<i", unparsed, off)[0]; off += 4
        result["margin_left"] = struct.unpack_from("<H", unparsed, off)[0]; off += 2
        result["margin_right"] = struct.unpack_from("<H", unparsed, off)[0]; off += 2
        result["margin_top"] = struct.unpack_from("<H", unparsed, off)[0]; off += 2
        result["margin_bottom"] = struct.unpack_from("<H", unparsed, off)[0]; off += 2
        result["instance_id"] = struct.unpack_from("<I", unparsed, off)[0]; off += 4
        if off + 4 <= len(unparsed):
            result["pagenum"] = struct.unpack_from("<i", unparsed, off)[0]; off += 4
        if off + 2 <= len(unparsed):
            desc_len = struct.unpack_from("<H", unparsed, off)[0]; off += 2
            if off + desc_len * 2 <= len(unparsed):
                result["description"] = unparsed[off:off + desc_len * 2].decode("utf-16-le", errors="replace")
                off += desc_len * 2
        return result

    @staticmethod
    def _parse_eqedit(payload):
        """Parse CTRL_EQEDIT payload for equation script, version, font, etc."""
        result = {}
        if not payload or len(payload) < 6:
            return result
        off = 0
        prop = struct.unpack_from("<I", payload, off)[0]; off += 4
        result["line_mode"] = prop & 1

        script_len = struct.unpack_from("<H", payload, off)[0]; off += 2
        if off + script_len * 2 <= len(payload):
            result["script"] = payload[off:off + script_len * 2].decode("utf-16-le", errors="replace")
            off += script_len * 2

        # base_unit (DWORD), text_color (DWORD), baseline (WORD), unknown (WORD)
        if off + 12 <= len(payload):
            result["base_unit"] = struct.unpack_from("<I", payload, off)[0]; off += 4
            result["text_color"] = struct.unpack_from("<I", payload, off)[0]; off += 4
            result["baseline"] = struct.unpack_from("<H", payload, off)[0]; off += 2
            off += 2  # skip unknown WORD

        # version string (WORD len + WCHAR[])
        if off + 2 <= len(payload):
            ver_len = struct.unpack_from("<H", payload, off)[0]; off += 2
            if ver_len > 0 and off + ver_len * 2 <= len(payload):
                result["version"] = payload[off:off + ver_len * 2].decode("utf-16-le", errors="replace")
                off += ver_len * 2

        # font string (WORD len + WCHAR[]) - may not be present
        if off + 2 <= len(payload):
            font_len = struct.unpack_from("<H", payload, off)[0]; off += 2
            if font_len > 0 and off + font_len * 2 <= len(payload):
                result["font"] = payload[off:off + font_len * 2].decode("utf-16-le", errors="replace")
                off += font_len * 2

        return result

    def _build_table_inline(self, run, ctrl_model, children_start, children_end):
        """Build table element inside a run element."""
        content = ctrl_model.get("content", {})

        # Find TABLE model among children
        table_model = None
        table_pos = children_start
        for i in range(children_start, children_end):
            if self.models[i].get("tagname") == "HWPTAG_TABLE":
                table_model = self.models[i].get("content", {})
                table_pos = i + 1
                break

        if table_model is None:
            return

        tbl = sub(run, "hp", "tbl")
        tbl.set("id", str(content.get("instance_id", 0)))
        tbl.set("zOrder", str(content.get("z_order", 0)))
        tbl.set("numberingType", "TABLE")

        ctrl_flags = content.get("flags", 0)
        text_wrap_type = (ctrl_flags >> 21) & 0x07
        text_flow_type = (ctrl_flags >> 24) & 0x03
        tbl.set("textWrap", vm.TEXT_WRAP_MAP.get(text_wrap_type, "SQUARE"))
        tbl.set("textFlow", vm.TEXT_FLOW_MAP.get(text_flow_type, "BOTH_SIDES"))
        tbl.set("lock", "0")
        tbl.set("dropcapstyle", "None")

        tbl_flags = table_model.get("flags", 0)
        page_break = (tbl_flags >> 0) & 0x03
        tbl.set("pageBreak", vm.PAGE_BREAK_MAP.get(page_break, "CELL"))
        repeat_header = (tbl_flags >> 2) & 0x01
        tbl.set("repeatHeader", str(repeat_header))

        rows = table_model.get("rows", 0)
        cols = table_model.get("cols", 0)
        tbl.set("rowCnt", str(rows))
        tbl.set("colCnt", str(cols))
        tbl.set("cellSpacing", str(table_model.get("cellspacing", 0)))
        tbl.set("borderFillIDRef", str(table_model.get("borderfill_id", 2)))
        tbl.set("noAdjust", "0")

        # sz
        sz = sub(tbl, "hp", "sz")
        sz.set("width", str(content.get("width", 0)))
        sz.set("widthRelTo", "ABSOLUTE")
        sz.set("height", str(content.get("height", 0)))
        sz.set("heightRelTo", "ABSOLUTE")
        sz.set("protect", "0")

        # pos - CommonControl property bit layout:
        # bit 0: treatAsChar, bit 2: affectLSpacing
        # bits 3-4: vertRelTo, bits 8-9: horzRelTo
        # bits 10-12: vertAlign, bits 14-16: horzAlign
        # bit 17: flowWithText, bit 18: allowOverlap
        pos = sub(tbl, "hp", "pos")
        pos.set("treatAsChar", str((ctrl_flags >> 0) & 0x01))
        pos.set("affectLSpacing", str((ctrl_flags >> 2) & 0x01))
        pos.set("flowWithText", str((ctrl_flags >> 17) & 0x01))
        pos.set("allowOverlap", str((ctrl_flags >> 18) & 0x01))
        pos.set("holdAnchorAndSO", "0")

        vert_rel = (ctrl_flags >> 3) & 0x03
        horz_rel = (ctrl_flags >> 8) & 0x03
        vert_align = (ctrl_flags >> 10) & 0x07
        horz_align = (ctrl_flags >> 14) & 0x07

        pos.set("vertRelTo", vm.VERT_REL_TO_MAP.get(vert_rel, "PARA"))
        pos.set("horzRelTo", vm.HORZ_REL_TO_MAP.get(horz_rel, "COLUMN"))
        pos.set("vertAlign", vm.VERT_ALIGN_MAP.get(vert_align, "TOP"))
        pos.set("horzAlign", vm.HORZ_ALIGN_MAP.get(horz_align, "LEFT"))
        pos.set("vertOffset", str(content.get("y", 0)))
        pos.set("horzOffset", str(content.get("x", 0)))

        # outMargin
        margin = content.get("margin", {})
        om = sub(tbl, "hp", "outMargin")
        om.set("left", str(margin.get("left", 283)))
        om.set("right", str(margin.get("right", 283)))
        om.set("top", str(margin.get("top", 283)))
        om.set("bottom", str(margin.get("bottom", 283)))

        # inMargin
        padding = table_model.get("padding", {})
        im = sub(tbl, "hp", "inMargin")
        im.set("left", str(padding.get("left", 510)))
        im.set("right", str(padding.get("right", 510)))
        im.set("top", str(padding.get("top", 141)))
        im.set("bottom", str(padding.get("bottom", 141)))

        # Process cells - need to navigate the model range [table_pos, children_end)
        self._build_table_cells(tbl, table_pos, children_end, ctrl_model.get("level", 1))

    def _build_table_cells(self, tbl, start_pos, end_pos, ctrl_level):
        """Build table rows and cells from model range.

        Cell content runs from after a LIST_HEADER to just before the next
        LIST_HEADER at the same level, or to end_pos. Paragraphs inside cells
        are at the SAME level as LIST_HEADER (not deeper).
        """
        # First, find all LIST_HEADER positions to determine cell boundaries
        cell_level = ctrl_level + 1
        list_header_positions = []
        for i in range(start_pos, end_pos):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_LIST_HEADER" and m.get("level", 0) == cell_level:
                list_header_positions.append(i)

        current_row_idx = -1
        tr = None

        for idx, lh_pos in enumerate(list_header_positions):
            cell_content = self.models[lh_pos].get("content", {})
            cell_row = cell_content.get("row", 0)

            if cell_row != current_row_idx:
                tr = sub(tbl, "hp", "tr")
                current_row_idx = cell_row

            # Cell content: from lh_pos+1 to next LIST_HEADER or end_pos
            cell_start = lh_pos + 1
            if idx + 1 < len(list_header_positions):
                cell_end = list_header_positions[idx + 1]
            else:
                cell_end = end_pos

            self._build_table_cell(tr, cell_content, cell_start, cell_end)

    def _build_table_cell(self, tr, cell_content, children_start, children_end):
        """Build a table cell element."""
        tc = sub(tr, "hp", "tc")
        tc.set("name", "")
        # header flag is at bit 18 of listflags
        list_flags = cell_content.get("listflags", 0)
        tc.set("header", str((list_flags >> 18) & 0x01))
        tc.set("hasMargin", "0")
        tc.set("protect", "0")
        tc.set("editable", "0")
        tc.set("dirty", "0")
        tc.set("borderFillIDRef", str(cell_content.get("borderfill_id", 3)))

        # subList
        sl = sub(tc, "hp", "subList")
        sl.set("id", "")
        text_dir = (list_flags >> 16) & 0x03
        sl.set("textDirection", vm.TEXT_DIRECTION_MAP.get(text_dir, "HORIZONTAL"))
        sl.set("lineWrap", "BREAK")
        vert_align = (list_flags >> 8) & 0x03
        sl.set("vertAlign", vm.CELL_VERT_ALIGN_MAP.get(vert_align, "TOP"))
        sl.set("linkListIDRef", "0")
        sl.set("linkListNextIDRef", "0")
        sl.set("textWidth", str(cell_content.get("width", 0)))
        sl.set("textHeight", "0")
        sl.set("hasTextRef", "0")
        sl.set("hasNumRef", "0")

        # Process paragraphs inside the cell using a sub-context
        old_pos = self.pos
        self.pos = children_start
        while self.pos < children_end:
            model = self.models[self.pos]
            if model.get("tagname") == "HWPTAG_PARA_HEADER":
                self._process_paragraph(sl)
            else:
                self.advance()
        self.pos = old_pos

        # cellAddr
        addr = sub(tc, "hp", "cellAddr")
        addr.set("colAddr", str(cell_content.get("col", 0)))
        addr.set("rowAddr", str(cell_content.get("row", 0)))

        # cellSpan
        span = sub(tc, "hp", "cellSpan")
        span.set("colSpan", str(cell_content.get("colspan", 1)))
        span.set("rowSpan", str(cell_content.get("rowspan", 1)))

        # cellSz
        csz = sub(tc, "hp", "cellSz")
        csz.set("width", str(cell_content.get("width", 0)))
        csz.set("height", str(cell_content.get("height", 0)))

        # cellMargin
        padding = cell_content.get("padding", {})
        cm = sub(tc, "hp", "cellMargin")
        cm.set("left", str(padding.get("left", 510)))
        cm.set("right", str(padding.get("right", 510)))
        cm.set("top", str(padding.get("top", 141)))
        cm.set("bottom", str(padding.get("bottom", 141)))

    # ---------- GSO (Graphic/Shape Object) builders ----------

    def _build_gso_inline(self, run, ctrl_model, children_start, children_end):
        """Build GSO element (picture, rectangle, line, container) inside a run."""
        content = ctrl_model.get("content", {})
        ctrl_level = ctrl_model.get("level", 1)

        # Find the top-level SHAPE_COMPONENT to determine shape type
        for i in range(children_start, children_end):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_SHAPE_COMPONENT" and m.get("level", 0) == ctrl_level + 1:
                sc_content = m.get("content", {})
                chid = sc_content.get("chid", "").strip()

                if chid == "$pic":
                    self._build_picture(run, content, sc_content, i + 1, children_end, ctrl_level + 1)
                elif chid == "$rec":
                    self._build_rectangle(run, content, sc_content, i + 1, children_end, ctrl_level + 1)
                elif chid == "$con":
                    self._build_container(run, content, sc_content, i + 1, children_end, ctrl_level + 1)
                elif chid == "$lin":
                    self._build_line_shape(run, content, sc_content, i + 1, children_end, ctrl_level + 1)
                elif chid == "$ell":
                    self._build_ellipse(run, content, sc_content, i + 1, children_end, ctrl_level + 1)
                # $arc, $pol, $cur, $tat (textart), $ole: not yet implemented
                break

    def _gso_common_attrs(self, elem, ctrl_content, numbering_type="PICTURE", sc_content=None):
        """Set common GSO attributes on an element.

        Hancom HWPX reference order for all GSO elements:
        attrs → hp:offset, hp:orgSz, hp:curSz, hp:flip, hp:rotationInfo,
        hp:renderingInfo, [element-specific], hp:sz, hp:pos, hp:outMargin
        """
        elem.set("id", str(ctrl_content.get("instance_id", 0)))
        elem.set("zOrder", str(ctrl_content.get("z_order", 0)))
        elem.set("numberingType", numbering_type)

        ctrl_flags = ctrl_content.get("flags", 0)
        text_wrap_type = (ctrl_flags >> 21) & 0x07
        text_flow = (ctrl_flags >> 24) & 0x03
        elem.set("textWrap", vm.TEXT_WRAP_MAP.get(text_wrap_type, "TOP_AND_BOTTOM"))
        elem.set("textFlow", vm.TEXT_FLOW_MAP.get(text_flow, "BOTH_SIDES"))
        elem.set("lock", "0")
        elem.set("dropcapstyle", "None")
        elem.set("href", "")
        src = sc_content if sc_content else {}
        elem.set("groupLevel", str(src.get("level_in_group", 0)))
        elem.set("instid", str(ctrl_content.get("instance_id", 0)))

        # hp:offset
        offset = sub(elem, "hp", "offset")
        x_in_group = src.get("x_in_group", 0)
        y_in_group = src.get("y_in_group", 0)
        if x_in_group < 0:
            x_in_group = x_in_group + 4294967296
        if y_in_group < 0:
            y_in_group = y_in_group + 4294967296
        offset.set("x", str(x_in_group))
        offset.set("y", str(y_in_group))

        # hp:orgSz
        org_sz = sub(elem, "hp", "orgSz")
        org_sz.set("width", str(src.get("initial_width", 0)))
        org_sz.set("height", str(src.get("initial_height", 0)))

        # hp:curSz
        cur_sz = sub(elem, "hp", "curSz")
        cur_sz.set("width", str(src.get("width", 0)))
        cur_sz.set("height", str(src.get("height", 0)))

        # hp:flip
        flip = sub(elem, "hp", "flip")
        flip.set("horizontal", "0")
        flip.set("vertical", "0")

        # hp:rotationInfo
        rot_info = sub(elem, "hp", "rotationInfo")
        rot_info.set("angle", str(src.get("angle", 0)))
        rot_center = src.get("rotation_center", {})
        rot_info.set("centerX", str(rot_center.get("x", 0)))
        rot_info.set("centerY", str(rot_center.get("y", 0)))
        rot_info.set("rotateimage", "1")

        # hp:renderingInfo
        rendering = sub(elem, "hp", "renderingInfo")
        trans = src.get("translation", {})
        trans_m = sub(rendering, "hc", "transMatrix")
        trans_m.set("e1", str(trans.get("a", 1.0)))
        trans_m.set("e2", str(trans.get("c", 0.0)))
        trans_m.set("e3", str(trans.get("e", 0.0)))
        trans_m.set("e4", str(trans.get("b", 0.0)))
        trans_m.set("e5", str(trans.get("d", 1.0)))
        trans_m.set("e6", str(trans.get("f", 0.0)))

        scalerotations = src.get("scalerotations", [])
        if scalerotations:
            for sr in scalerotations:
                scaler = sr.get("scaler", {})
                sca_m = sub(rendering, "hc", "scaMatrix")
                sca_m.set("e1", str(scaler.get("a", 1.0)))
                sca_m.set("e2", str(scaler.get("c", 0.0)))
                sca_m.set("e3", str(scaler.get("e", 0.0)))
                sca_m.set("e4", str(scaler.get("b", 0.0)))
                sca_m.set("e5", str(scaler.get("d", 1.0)))
                sca_m.set("e6", str(scaler.get("f", 0.0)))

                rotator = sr.get("rotator", {})
                rot_m = sub(rendering, "hc", "rotMatrix")
                rot_m.set("e1", str(rotator.get("a", 1.0)))
                rot_m.set("e2", str(rotator.get("c", 0.0)))
                rot_m.set("e3", str(rotator.get("e", 0.0)))
                rot_m.set("e4", str(rotator.get("b", 0.0)))
                rot_m.set("e5", str(rotator.get("d", 1.0)))
                rot_m.set("e6", str(rotator.get("f", 0.0)))
        else:
            for mname in ["scaMatrix", "rotMatrix"]:
                mat = sub(rendering, "hc", mname)
                mat.set("e1", "1")
                mat.set("e2", "0")
                mat.set("e3", "0")
                mat.set("e4", "0")
                mat.set("e5", "1")
                mat.set("e6", "0")

        # lineShape from SHAPE_COMPONENT's line/border properties
        src = sc_content if sc_content else {}
        line_props = src.get("line", src.get("border", {}))
        if line_props:
            ls = sub(elem, "hp", "lineShape")
            line_color = line_props.get("color", 0)
            line_width = line_props.get("width", 0)
            line_flags = line_props.get("flags", 0)
            line_stroke = line_flags & 0x1F
            ls.set("color", vm.color_from_int(line_color))
            ls.set("width", str(line_width))
            ls.set("style", vm.STROKE_TYPE_MAP.get(line_stroke, "SOLID"))
            ls.set("endCap", "FLAT")
            head_style = (line_flags >> 10) & 0x0F
            tail_style = (line_flags >> 14) & 0x0F
            ls.set("headStyle", _arrow_style(head_style))
            ls.set("tailStyle", _arrow_style(tail_style))
            ls.set("headfill", "0")
            ls.set("tailfill", "0")
            ls.set("headSz", "MEDIUM_MEDIUM")
            ls.set("tailSz", "MEDIUM_MEDIUM")
            ls.set("outlineStyle", "NORMAL")
            ls.set("alpha", "0")

        # fillBrush from SHAPE_COMPONENT fill data
        fill_flags = src.get("fill_flags", 0)
        if fill_flags:
            fb = None
            # Solid fill (bit 0)
            if fill_flags & 0x01:
                face_color = src.get("fill_face_color", src.get("fill_color", None))
                if face_color is not None:
                    if fb is None:
                        fb = sub(elem, "hc", "fillBrush")
                    wb = sub(fb, "hc", "winBrush")
                    wb.set("faceColor", vm.color_from_int(face_color))
                    hatch_color = src.get("fill_hatch_color", None)
                    wb.set("hatchColor", vm.color_from_int(hatch_color) if hatch_color is not None else "#FF000000")
                    wb.set("alpha", "0")
            # Gradation fill (bit 2)
            if fill_flags & 0x04:
                grad = src.get("fill_gradation", {})
                if grad:
                    if fb is None:
                        fb = sub(elem, "hc", "fillBrush")
                    ge = sub(fb, "hc", "gradation")
                    grad_type = grad.get("type", 1)
                    ge.set("type", vm.GRADATION_TYPE_MAP.get(grad_type, "LINEAR"))
                    ge.set("angle", str(grad.get("shear", 0)))
                    center = grad.get("center", {})
                    ge.set("centerX", str(center.get("x", 50)))
                    ge.set("centerY", str(center.get("y", 50)))
                    ge.set("step", str(grad.get("blur", 50)))
                    colors = grad.get("colors", [])
                    ge.set("colorNum", str(len(colors)))
                    for c in colors:
                        cs_elem = sub(ge, "hc", "color")
                        cs_elem.set("value", vm.color_from_int(c))

    def _gso_common_attrs_tail(self, elem, ctrl_content, sc_content=None):
        """Emit sz, pos, outMargin as trailing children (for pic elements where these come last)."""
        ctrl_flags = ctrl_content.get("flags", 0)

        # hp:sz
        sz = sub(elem, "hp", "sz")
        sz.set("width", str(ctrl_content.get("width", 0)))
        sz.set("widthRelTo", "ABSOLUTE")
        sz.set("height", str(ctrl_content.get("height", 0)))
        sz.set("heightRelTo", "ABSOLUTE")
        sz.set("protect", "0")

        # hp:pos
        pos = sub(elem, "hp", "pos")
        pos.set("treatAsChar", str((ctrl_flags >> 0) & 0x01))
        pos.set("affectLSpacing", str((ctrl_flags >> 2) & 0x01))
        pos.set("flowWithText", str((ctrl_flags >> 17) & 0x01))
        pos.set("allowOverlap", str((ctrl_flags >> 18) & 0x01))
        pos.set("holdAnchorAndSO", "0")

        vert_rel = (ctrl_flags >> 3) & 0x03
        horz_rel = (ctrl_flags >> 8) & 0x03
        vert_align = (ctrl_flags >> 10) & 0x07
        horz_align = (ctrl_flags >> 14) & 0x07

        pos.set("vertRelTo", vm.VERT_REL_TO_MAP.get(vert_rel, "PARA"))
        pos.set("horzRelTo", vm.HORZ_REL_TO_MAP.get(horz_rel, "COLUMN"))
        pos.set("vertAlign", vm.VERT_ALIGN_MAP.get(vert_align, "TOP"))
        pos.set("horzAlign", vm.HORZ_ALIGN_MAP.get(horz_align, "LEFT"))
        pos.set("vertOffset", str(ctrl_content.get("y", 0)))
        pos.set("horzOffset", str(ctrl_content.get("x", 0)))

        # hp:outMargin
        margin = ctrl_content.get("margin", {})
        om = sub(elem, "hp", "outMargin")
        om.set("left", str(margin.get("left", 0)))
        om.set("right", str(margin.get("right", 0)))
        om.set("top", str(margin.get("top", 0)))
        om.set("bottom", str(margin.get("bottom", 0)))

    def _build_picture(self, parent, ctrl_content, sc_content, children_start, children_end, sc_level, in_container=False):
        """Build hp:pic element for an image.

        Hancom HWPX reference element order:
        hp:pic attrs → hp:offset, hp:orgSz, hp:curSz, hp:flip,
        hp:rotationInfo, hp:renderingInfo, hp:imgRect, hp:imgClip,
        hp:inMargin, hp:imgDim, hc:img (leaf), hp:effects,
        hp:sz, hp:pos, hp:outMargin, hp:shapeComment
        """
        pic = sub(parent, "hp", "pic")

        # Common GSO head: attrs + offset, orgSz, curSz, flip, rotationInfo, renderingInfo
        self._gso_common_attrs(pic, ctrl_content, "PICTURE", sc_content)

        # Find SHAPE_COMPONENT_PICTURE child for image-specific data
        pic_found = False
        for i in range(children_start, children_end):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_SHAPE_COMPONENT_PICTURE":
                pic_content = m.get("content", {})
                picture = pic_content.get("picture", {})
                bindata_id = picture.get("bindata_id", 0)

                # Resolve to a label that exists in the manifest/ZIP.
                # bindata_id is a 1-based index into the BIN_DATA list.
                image_label = self._bindata_id_map.get(bindata_id)
                if not image_label:
                    # No actual binary data for this reference — drop the picture
                    parent.remove(pic)
                    return

                pic_found = True

                # hp:imgRect as child of hp:pic
                rect = pic_content.get("rect", {})
                img_rect = sub(pic, "hp", "imgRect")
                for pt_name in ["pt0", "pt1", "pt2", "pt3"]:
                    src_name = pt_name.replace("pt", "p")
                    pt = rect.get(src_name, {"x": 0, "y": 0})
                    sub(img_rect, "hc", pt_name, {"x": str(pt.get("x", 0)), "y": str(pt.get("y", 0))})

                # hp:imgClip as child of hp:pic
                clip = pic_content.get("clip", {})
                img_clip = sub(pic, "hp", "imgClip")
                img_clip.set("left", str(clip.get("left", 0)))
                img_clip.set("right", str(clip.get("right", 0)))
                img_clip.set("top", str(clip.get("top", 0)))
                img_clip.set("bottom", str(clip.get("bottom", 0)))

                # hp:inMargin
                padding = pic_content.get("padding", {})
                in_margin = sub(pic, "hp", "inMargin")
                in_margin.set("left", str(padding.get("left", 0)))
                in_margin.set("right", str(padding.get("right", 0)))
                in_margin.set("top", str(padding.get("top", 0)))
                in_margin.set("bottom", str(padding.get("bottom", 0)))

                # hp:imgDim (use clip dimensions as original image dims)
                img_dim = sub(pic, "hp", "imgDim")
                img_dim.set("dimwidth", str(clip.get("right", 0)))
                img_dim.set("dimheight", str(clip.get("bottom", 0)))

                # hc:img as a LEAF element (no children)
                img = sub(pic, "hc", "img")
                img.set("binaryItemIDRef", image_label)
                img.set("bright", str(picture.get("brightness", 0)))
                img.set("contrast", str(picture.get("contrast", 0)))
                effect = picture.get("effect", 0)
                img.set("effect", vm.PICTURE_EFFECT_MAP.get(effect, "REAL_PIC"))
                img.set("alpha", "0")

                # hp:effects (empty placeholder)
                sub(pic, "hp", "effects")

                break

        if not pic_found:
            parent.remove(pic)
            return

        # hp:sz, hp:pos, hp:outMargin AFTER image data (not for container children)
        if not in_container:
            self._gso_common_attrs_tail(pic, ctrl_content, sc_content)

        # hp:shapeComment (use description from ctrl_content if available)
        comment = sub(pic, "hp", "shapeComment")
        desc = ctrl_content.get("description", "")
        if desc:
            comment.text = desc

    def _build_rectangle(self, parent, ctrl_content, sc_content, children_start, children_end, sc_level, in_container=False):
        """Build hp:rect element for a text box / rectangle shape."""
        rect_elem = sub(parent, "hp", "rect")
        self._gso_common_attrs(rect_elem, ctrl_content, "PICTURE", sc_content)
        # Find SHAPE_COMPONENT_RECTANGLE for coordinates
        for i in range(children_start, children_end):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_SHAPE_COMPONENT_RECTANGLE":
                rc = m.get("content", {})
                coord = sub(rect_elem, "hc", "pt0")
                coord.set("x", str(rc.get("p0", {}).get("x", 0)))
                coord.set("y", str(rc.get("p0", {}).get("y", 0)))
                coord = sub(rect_elem, "hc", "pt1")
                coord.set("x", str(rc.get("p1", {}).get("x", 0)))
                coord.set("y", str(rc.get("p1", {}).get("y", 0)))
                coord = sub(rect_elem, "hc", "pt2")
                coord.set("x", str(rc.get("p2", {}).get("x", 0)))
                coord.set("y", str(rc.get("p2", {}).get("y", 0)))
                coord = sub(rect_elem, "hc", "pt3")
                coord.set("x", str(rc.get("p3", {}).get("x", 0)))
                coord.set("y", str(rc.get("p3", {}).get("y", 0)))
                break

        # Find LIST_HEADER + paragraphs for text content inside rectangle
        for i in range(children_start, children_end):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_LIST_HEADER" and m.get("level", 0) == sc_level + 1:
                lh_content = m.get("content", {})
                sl = sub(rect_elem, "hp", "subList")
                sl.set("id", "")
                sl.set("textDirection", "HORIZONTAL")
                sl.set("lineWrap", "BREAK")
                sl.set("vertAlign", "CENTER")
                sl.set("linkListIDRef", "0")
                sl.set("linkListNextIDRef", "0")
                sl.set("textWidth", str(lh_content.get("maxwidth", 0)))
                sl.set("textHeight", "0")
                sl.set("hasTextRef", "0")
                sl.set("hasNumRef", "0")

                # Process paragraphs inside
                para_start = i + 1
                para_end = children_end
                # Find end of this list header's children
                for j in range(i + 1, children_end):
                    if self.models[j].get("level", 0) <= sc_level + 1:
                        if self.models[j].get("tagname") != "HWPTAG_PARA_HEADER":
                            para_end = j
                            break
                        # If another LIST_HEADER at same level, stop
                        if self.models[j].get("tagname") == "HWPTAG_LIST_HEADER":
                            para_end = j
                            break
                    # Keep going for paragraph children at deeper levels

                # Find actual end - paragraphs are at same level as list header
                para_end = children_end
                for j in range(i + 1, children_end):
                    mj = self.models[j]
                    if mj.get("level", 0) <= sc_level and mj.get("tagname") != "HWPTAG_PARA_HEADER":
                        para_end = j
                        break
                    if mj.get("tagname") == "HWPTAG_SHAPE_COMPONENT_RECTANGLE":
                        para_end = j
                        break

                old_pos = self.pos
                self.pos = para_start
                while self.pos < para_end:
                    model = self.models[self.pos]
                    if model.get("tagname") == "HWPTAG_PARA_HEADER":
                        self._process_paragraph(sl)
                    else:
                        self.advance()
                self.pos = old_pos
                break

        if not in_container:
            self._gso_common_attrs_tail(rect_elem, ctrl_content)

    def _make_child_ctrl_content(self, parent_ctrl_content, child_sc_content):
        """Create a ctrl_content-like dict for a child shape inside a container,
        using the child's own dimensions from its SHAPE_COMPONENT record."""
        child_ctrl = dict(parent_ctrl_content)
        # Compute final dimensions from the full scalerotation chain
        final_w, final_h = _compute_final_dimensions(child_sc_content)
        if final_w is not None:
            child_ctrl["width"] = final_w
        if final_h is not None:
            child_ctrl["height"] = final_h
        child_ctrl["x"] = child_sc_content.get("x_in_group", 0)
        child_ctrl["y"] = child_sc_content.get("y_in_group", 0)
        return child_ctrl

    def _build_container(self, parent, ctrl_content, sc_content, children_start, children_end, sc_level):
        """Build hp:container element for grouped shapes."""
        container = sub(parent, "hp", "container")
        self._gso_common_attrs(container, ctrl_content, "PICTURE", sc_content)

        # Process child SHAPE_COMPONENTs (before sz/pos/outMargin)
        i = children_start
        while i < children_end:
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_SHAPE_COMPONENT" and m.get("level", 0) == sc_level + 1:
                child_sc = m.get("content", {})
                child_chid = child_sc.get("chid", "").strip()

                # Find end of this child shape
                child_end = i + 1
                while child_end < children_end:
                    if self.models[child_end].get("level", 0) <= sc_level + 1:
                        break
                    child_end += 1

                # Build child ctrl_content with child's own dimensions
                child_ctrl = self._make_child_ctrl_content(ctrl_content, child_sc)

                if child_chid == "$pic":
                    self._build_picture(container, child_ctrl, child_sc, i + 1, child_end, sc_level + 1, in_container=True)
                elif child_chid == "$rec":
                    self._build_rectangle(container, child_ctrl, child_sc, i + 1, child_end, sc_level + 1, in_container=True)
                elif child_chid == "$lin":
                    self._build_line_shape(container, child_ctrl, child_sc, i + 1, child_end, sc_level + 1, in_container=True)
                elif child_chid == "$ell":
                    self._build_ellipse(container, child_ctrl, child_sc, i + 1, child_end, sc_level + 1, in_container=True)
                elif child_chid == "$con":
                    self._build_container(container, child_ctrl, child_sc, i + 1, child_end, sc_level + 1)

                i = child_end
            else:
                i += 1

        # sz/pos/outMargin AFTER children (matches Hancom reference)
        self._gso_common_attrs_tail(container, ctrl_content)

    def _build_line_shape(self, parent, ctrl_content, sc_content, children_start, children_end, sc_level, in_container=False):
        """Build hp:line element for a line shape."""
        line = sub(parent, "hp", "line")
        self._gso_common_attrs(line, ctrl_content, "PICTURE", sc_content)

        # Find SHAPE_COMPONENT_LINE child
        for i in range(children_start, children_end):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_SHAPE_COMPONENT_LINE":
                lc = m.get("content", {})
                p0 = lc.get("p0", {"x": 0, "y": 0})
                p1 = lc.get("p1", {"x": 0, "y": 0})
                # Transform through scalerotation chain for actual coordinates
                x0, y0 = _transform_point(p0.get("x", 0), p0.get("y", 0), sc_content)
                x1, y1 = _transform_point(p1.get("x", 0), p1.get("y", 0), sc_content)
                sub(line, "hc", "startPt", {"x": str(x0), "y": str(y0)})
                sub(line, "hc", "endPt", {"x": str(x1), "y": str(y1)})
                break
        if not in_container:
            self._gso_common_attrs_tail(line, ctrl_content)

    def _build_ellipse(self, parent, ctrl_content, sc_content, children_start, children_end, sc_level, in_container=False):
        """Build hp:ellipse element for an ellipse shape."""
        ellipse = sub(parent, "hp", "ellipse")
        self._gso_common_attrs(ellipse, ctrl_content, "PICTURE", sc_content)

        # Find SHAPE_COMPONENT_ELLIPSE child
        for i in range(children_start, children_end):
            m = self.models[i]
            if m.get("tagname") == "HWPTAG_SHAPE_COMPONENT_ELLIPSE":
                ec = m.get("content", {})
                cx = ec.get("cx", 0)
                cy = ec.get("cy", 0)
                rx = ec.get("rx", 0)
                ry = ec.get("ry", 0)
                ellipse.set("intervalDirty", "0")
                ellipse.set("hasArcPr", "0")
                sub(ellipse, "hp", "ax", {"x": str(cx + rx), "y": str(cy)})
                sub(ellipse, "hp", "ay", {"x": str(cx), "y": str(cy + ry)})
                sub(ellipse, "hp", "center", {"x": str(cx), "y": str(cy)})
                break
        if not in_container:
            self._gso_common_attrs_tail(ellipse, ctrl_content)

    # ---------- Helper builders ----------

    def _build_page_pr(self, sec_pr, page_def):
        """Build pagePr element."""
        width = page_def.get("width", 59528)
        height = page_def.get("height", 84188)
        attr_val = page_def.get("attr", 0)

        # NOTE:
        # Some real-world Korean official documents render more correctly in Hancom
        # when page orientation is treated as landscape-first, even when raw HWP page
        # dimensions are portrait-shaped. Prefer explicit attr bit if it exists; fall back
        # to a landscape-first heuristic for document compatibility.
        orientation_flag = attr_val & 0x01
        if orientation_flag:
            landscape = "WIDELY"
        else:
            landscape = "WIDELY" if width <= height else "NARROWLY"

        gutter_type = (attr_val >> 1) & 0x03

        pp = sub(sec_pr, "hp", "pagePr")
        pp.set("landscape", landscape)
        pp.set("width", str(width))
        pp.set("height", str(height))
        pp.set("gutterType", vm.GUTTER_TYPE_MAP.get(gutter_type, "LEFT_ONLY"))

        margin = sub(pp, "hp", "margin")
        margin.set("header", str(page_def.get("header_offset", 4252)))
        margin.set("footer", str(page_def.get("footer_offset", 4252)))
        margin.set("gutter", str(page_def.get("bookbinding_offset", 0)))
        margin.set("left", str(page_def.get("left_offset", 8504)))
        margin.set("right", str(page_def.get("right_offset", 8504)))
        margin.set("top", str(page_def.get("top_offset", 5668)))
        margin.set("bottom", str(page_def.get("bottom_offset", 4252)))

    def _build_footnote_pr(self, sec_pr, fn_shape, tag_name):
        """Build footNotePr or endNotePr element."""
        fnp = sub(sec_pr, "hp", tag_name)

        anf = sub(fnp, "hp", "autoNumFormat")
        anf.set("type", vm.NUM_FORMAT_MAP.get(fn_shape.get("flags", 0) & 0x0F, "DIGIT"))
        anf.set("userChar", "")
        anf.set("prefixChar", "")
        suffix_code = fn_shape.get("suffix", 0)
        anf.set("suffixChar", vm.SUFFIX_CHAR_MAP.get(suffix_code, chr(suffix_code) if suffix_code else ""))
        anf.set("supscript", "0")

        nl = sub(fnp, "hp", "noteLine")
        nl.set("length", str(fn_shape.get("splitter_length", -1)))
        stroke = fn_shape.get("splitter_stroke_type", 1)
        nl.set("type", vm.STROKE_TYPE_MAP.get(stroke, "SOLID"))
        width = fn_shape.get("splitter_width", 1)
        nl.set("width", vm.BORDER_WIDTH_MAP.get(width, "0.12 mm"))
        nl.set("color", vm.color_from_int(fn_shape.get("splitter_color", 0)))

        ns = sub(fnp, "hp", "noteSpacing")
        ns.set("betweenNotes", str(fn_shape.get("notes_spacing", 283)))
        ns.set("belowLine", str(fn_shape.get("splitter_margin_bottom", 567)))
        ns.set("aboveLine", str(fn_shape.get("splitter_margin_top", 850)))

        num = sub(fnp, "hp", "numbering")
        num_type = (fn_shape.get("flags", 0) >> 4) & 0x03
        if "endNote" in tag_name:
            num.set("type", vm.ENDNOTE_NUMBERING_MAP.get(num_type, "CONTINUOUS"))
        else:
            num.set("type", vm.FOOTNOTE_NUMBERING_MAP.get(num_type, "CONTINUOUS"))
        num.set("newNum", str(fn_shape.get("starting_number", 1)))

        plc = sub(fnp, "hp", "placement")
        plc_type = (fn_shape.get("flags", 0) >> 6) & 0x03
        if "endNote" in tag_name:
            plc.set("place", vm.ENDNOTE_PLACE_MAP.get(plc_type, "END_OF_DOCUMENT"))
        else:
            plc.set("place", vm.FOOTNOTE_PLACE_MAP.get(plc_type, "EACH_COLUMN"))
        plc.set("beneathText", "0")

    def _build_page_border_fill(self, sec_pr, pbf, pbf_type="BOTH"):
        """Build pageBorderFill element."""
        flags = pbf.get("flags", 0)
        pf = vm.extract_page_border_fill_flags(flags)

        elem = sub(sec_pr, "hp", "pageBorderFill")
        elem.set("type", pbf_type)
        elem.set("borderFillIDRef", str(pbf.get("borderfill_id", 1)))
        elem.set("textBorder", pf["textBorder"])
        elem.set("headerInside", pf["headerInside"])
        elem.set("footerInside", pf["footerInside"])
        elem.set("fillArea", pf["fillArea"])

        margin = pbf.get("margin", {})
        offset = sub(elem, "hp", "offset")
        offset.set("left", str(margin.get("left", 1417)))
        offset.set("right", str(margin.get("right", 1417)))
        offset.set("top", str(margin.get("top", 1417)))
        offset.set("bottom", str(margin.get("bottom", 1417)))
