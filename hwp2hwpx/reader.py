"""HWP file reader using pyhwp (hwp5) and olefile."""

import olefile
from hwp5.xmlmodel import Hwp5File


class HWPReader:
    """Reads an HWP file and provides structured access to its contents."""

    def __init__(self, filepath):
        self.filepath = filepath
        self.hwp5 = Hwp5File(filepath)
        self.ole = olefile.OleFileIO(filepath)
        self._docinfo_cache = None
        self._section_cache = {}

    def close(self):
        self.hwp5.close()
        self.ole.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    # --- DocInfo ---

    def get_docinfo_models(self):
        """Return list of all docinfo model dicts."""
        if self._docinfo_cache is None:
            self._docinfo_cache = list(self.hwp5.docinfo.models())
        return self._docinfo_cache

    def get_docinfo_by_tag(self, tagname):
        """Return list of models matching given tagname."""
        return [m for m in self.get_docinfo_models() if m.get("tagname") == tagname]

    def get_document_properties(self):
        models = self.get_docinfo_by_tag("HWPTAG_DOCUMENT_PROPERTIES")
        return models[0]["content"] if models else {}

    def get_id_mappings(self):
        models = self.get_docinfo_by_tag("HWPTAG_ID_MAPPINGS")
        return models[0]["content"] if models else {}

    def get_face_names(self):
        """Return all face names as list of content dicts."""
        return [m["content"] for m in self.get_docinfo_by_tag("HWPTAG_FACE_NAME")]

    def get_border_fills(self):
        return [m["content"] for m in self.get_docinfo_by_tag("HWPTAG_BORDER_FILL")]

    def get_char_shapes(self):
        return [m["content"] for m in self.get_docinfo_by_tag("HWPTAG_CHAR_SHAPE")]

    def get_tab_defs(self):
        return [m["content"] for m in self.get_docinfo_by_tag("HWPTAG_TAB_DEF")]

    def get_numberings(self):
        return [m["content"] for m in self.get_docinfo_by_tag("HWPTAG_NUMBERING")]

    def get_bullets(self):
        return [m["content"] for m in self.get_docinfo_by_tag("HWPTAG_BULLET")]

    def get_para_shapes(self):
        return [m["content"] for m in self.get_docinfo_by_tag("HWPTAG_PARA_SHAPE")]

    def get_styles(self):
        return [m["content"] for m in self.get_docinfo_by_tag("HWPTAG_STYLE")]

    def get_bin_data_list(self):
        return [m["content"] for m in self.get_docinfo_by_tag("HWPTAG_BIN_DATA")]

    def get_compatible_document(self):
        models = self.get_docinfo_by_tag("HWPTAG_COMPATIBLE_DOCUMENT")
        return models[0]["content"] if models else {}

    # --- BodyText ---

    def get_section_count(self):
        props = self.get_document_properties()
        return props.get("section_count", 1)

    def get_section_models(self, section_idx):
        """Return list of all models for a section."""
        if section_idx not in self._section_cache:
            self._section_cache[section_idx] = list(
                self.hwp5.bodytext.section(section_idx).models()
            )
        return self._section_cache[section_idx]

    # --- BinData (embedded files) ---

    def get_bindata_bytes(self, storage_id, ext):
        """Read embedded binary data from BinData stream."""
        stream_name = f"BinData/BIN{storage_id:04d}.{ext}"
        if self.ole.exists(stream_name):
            return self.ole.openstream(stream_name).read()
        # Try uppercase
        stream_name_upper = f"BinData/BIN{storage_id:04d}.{ext.upper()}"
        if self.ole.exists(stream_name_upper):
            return self.ole.openstream(stream_name_upper).read()
        return None

    # --- Summary Information ---

    def get_summary_info(self):
        """Read HwpSummaryInformation stream."""
        result = {}
        stream_name = "\x05HwpSummaryInformation"
        if not self.ole.exists(stream_name):
            stream_name = "HwpSummaryInformation"
        if self.ole.exists(stream_name):
            try:
                meta = self.ole.get_metadata()
                result["title"] = meta.title or ""
                result["subject"] = meta.subject or ""
                result["author"] = meta.author or ""
                result["keywords"] = meta.keywords or ""
                result["comments"] = meta.comments or ""
                result["last_saved_by"] = meta.last_saved_by or ""
                result["creating_application"] = meta.creating_application or ""
                result["create_time"] = meta.create_time
                result["last_saved_time"] = meta.last_saved_time
            except Exception:
                pass
        return result

    # --- FileHeader ---

    def get_file_header(self):
        """Read FileHeader for version info."""
        result = {"major": 5, "minor": 0, "micro": 0, "build": 0}
        if self.ole.exists("FileHeader"):
            data = self.ole.openstream("FileHeader").read()
            if len(data) >= 36:
                # Version DWORD at offset 32 is little-endian: [build_lo, build_hi, minor, major]
                result["major"] = data[35]
                result["minor"] = data[34]
                result["micro"] = data[33]
                result["build"] = data[32]
        return result

    # --- Preview ---

    def get_preview_text(self):
        """Read PrvText stream."""
        if self.ole.exists("PrvText"):
            data = self.ole.openstream("PrvText").read()
            try:
                return data.decode("utf-16-le")
            except Exception:
                return ""
        return ""

    def get_preview_image(self):
        """Read PrvImage stream (PNG bytes)."""
        if self.ole.exists("PrvImage"):
            return self.ole.openstream("PrvImage").read()
        return None
