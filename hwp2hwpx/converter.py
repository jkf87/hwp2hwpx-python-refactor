"""Main HWP to HWPX converter - orchestrates the conversion pipeline."""

import os
import zipfile
from datetime import datetime

from .reader import HWPReader
from .xml_builder import root_element, sub, to_xml_bytes, NS, COMMON_NSMAP
from .header_converter import build_header_xml
from .section_converter import build_section_xml
from . import value_maps as vm


def convert_file(input_path, output_path=None):
    """Convert an HWP file to HWPX format.

    Args:
        input_path: Path to input .hwp file
        output_path: Path to output .hwpx file (default: same name with .hwpx extension)

    Returns:
        Path to the output .hwpx file
    """
    if output_path is None:
        base = os.path.splitext(input_path)[0]
        output_path = base + ".hwpx"

    with HWPReader(input_path) as reader:
        hwpx_data = convert(reader)

    _write_hwpx_zip(hwpx_data, output_path)
    return output_path


def convert(reader):
    """Convert HWP data to HWPX structure.

    Args:
        reader: HWPReader instance

    Returns:
        dict with HWPX file contents (file paths -> bytes)
    """
    files = {}

    # 1. mimetype
    files["mimetype"] = b"application/hwp+zip"

    # 2. version.xml
    files["version.xml"] = _build_version_xml(reader)

    # 3. META-INF/container.xml
    files["META-INF/container.xml"] = _build_container_xml()

    # 4. META-INF/manifest.xml (empty for now)
    files["META-INF/manifest.xml"] = _build_manifest_xml()

    # 5. BinData (embedded images/objects) — resolve first so manifest and
    #    section XML reference only entries that actually have binary data.
    bin_data_list = reader.get_bin_data_list()
    # bindata_id_map: {1-based BIN_DATA list index → "imageN"} for entries
    # whose binary payload was successfully read.  Section XML uses this to
    # translate picture.bindata_id → a valid binaryItemIDRef.
    bindata_id_map = {}
    image_counter = 0
    for i, bd in enumerate(bin_data_list):
        bindata = bd.get("bindata", {})
        if isinstance(bindata, dict):
            storage_id = bindata.get("storage_id", i + 1)
            ext = bindata.get("ext", "png")
            data = reader.get_bindata_bytes(storage_id, ext)
            if data:
                image_counter += 1
                label = f"image{image_counter}"
                bindata_id_map[i + 1] = label  # 1-based index → label
                files[f"BinData/{label}.{ext}"] = data

    # 6. Contents/content.hpf (needs bindata_id_map for manifest)
    files["Contents/content.hpf"] = _build_content_hpf(reader, bindata_id_map)

    # 7. Contents/header.xml
    header = build_header_xml(reader)
    files["Contents/header.xml"] = to_xml_bytes(header)

    # 8. Contents/section[N].xml (needs bindata_id_map for binaryItemIDRef)
    section_count = reader.get_section_count()
    for i in range(section_count):
        section = build_section_xml(reader, i, bindata_id_map)
        files[f"Contents/section{i}.xml"] = to_xml_bytes(section)

    # 9. settings.xml
    files["settings.xml"] = _build_settings_xml()

    # 10. Preview
    prv_text = reader.get_preview_text()
    if prv_text:
        files["Preview/PrvText.txt"] = prv_text.encode("utf-8")
    prv_image = reader.get_preview_image()
    if prv_image:
        files["Preview/PrvImage.png"] = prv_image

    return files


def _build_version_xml(reader):
    """Build version.xml."""
    fh = reader.get_file_header()
    nsmap = {"hv": NS["hv"]}
    root = root_element("hv", "HCFVersion", nsmap=nsmap)
    root.set("tagetApplication", "WORDPROCESSOR")
    root.set("major", str(fh.get("major", 5)))
    root.set("minor", str(fh.get("minor", 0)))
    root.set("micro", str(fh.get("micro", 0)))
    root.set("buildNumber", str(fh.get("build", 0)))
    root.set("os", "10")
    root.set("xmlVersion", "1.5")
    root.set("application", "Hancom Office Hangul")
    root.set("appVersion", "12.30.0.6313 MAC64LEDarwin_24.6.0")
    return to_xml_bytes(root)


def _build_container_xml():
    """Build META-INF/container.xml."""
    nsmap = {
        "ocf": NS["ocf"],
        "hpf": NS["hpf"],
    }
    container = root_element("ocf", "container", nsmap=nsmap)
    rootfiles = sub(container, "ocf", "rootfiles")
    rf1 = sub(rootfiles, "ocf", "rootfile")
    rf1.set("full-path", "Contents/content.hpf")
    rf1.set("media-type", "application/hwpml-package+xml")
    rf2 = sub(rootfiles, "ocf", "rootfile")
    rf2.set("full-path", "Preview/PrvText.txt")
    rf2.set("media-type", "text/plain")
    return to_xml_bytes(container)


def _build_manifest_xml():
    """Build META-INF/manifest.xml (minimal, with ODF namespace)."""
    return b'<?xml version="1.0" encoding="UTF-8"?><odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'


def _build_content_hpf(reader, bindata_id_map):
    """Build Contents/content.hpf with metadata, manifest, and spine."""
    root = root_element("opf", "package")
    root.set("version", "")
    root.set("unique-identifier", "")
    root.set("id", "")

    # metadata
    metadata = sub(root, "opf", "metadata")

    summary = reader.get_summary_info()
    sub(metadata, "opf", "title", text=summary.get("title", ""))
    sub(metadata, "opf", "language", text="ko")

    _meta(metadata, "creator", summary.get("author", ""))
    _meta(metadata, "subject", summary.get("subject", ""))
    _meta(metadata, "description", summary.get("comments", ""))
    _meta(metadata, "lastsaveby", summary.get("last_saved_by", ""))

    create_time = summary.get("create_time")
    if create_time and isinstance(create_time, datetime):
        _meta(metadata, "CreatedDate", create_time.strftime("%Y-%m-%dT%H:%M:%SZ"))
    else:
        _meta(metadata, "CreatedDate", "")

    mod_time = summary.get("last_saved_time")
    if mod_time and isinstance(mod_time, datetime):
        _meta(metadata, "ModifiedDate", mod_time.strftime("%Y-%m-%dT%H:%M:%SZ"))
        _meta(metadata, "date", mod_time.strftime("%Y년 %m월 %d일"))
    else:
        _meta(metadata, "ModifiedDate", "")
        _meta(metadata, "date", "")

    _meta(metadata, "keyword", summary.get("keywords", ""))

    # manifest
    manifest = sub(root, "opf", "manifest")

    item_header = sub(manifest, "opf", "item")
    item_header.set("id", "header")
    item_header.set("href", "Contents/header.xml")
    item_header.set("media-type", "application/xml")

    section_count = reader.get_section_count()
    for i in range(section_count):
        item = sub(manifest, "opf", "item")
        item.set("id", f"section{i}")
        item.set("href", f"Contents/section{i}.xml")
        item.set("media-type", "application/xml")

    # BinData items — only entries that have actual binary data
    bin_data_list = reader.get_bin_data_list()
    for j, bd in enumerate(bin_data_list):
        label = bindata_id_map.get(j + 1)
        if label is None:
            continue  # no binary data for this entry
        bindata = bd.get("bindata", {})
        ext = bindata.get("ext", "png") if isinstance(bindata, dict) else "png"
        media_info = vm.IMAGE_TYPE_MAP.get(ext.lower(), ("application/octet-stream", ""))
        item = sub(manifest, "opf", "item")
        item.set("id", label)
        item.set("href", f"BinData/{label}.{ext}")
        item.set("media-type", media_info[0])
        item.set("isEmbeded", "1")

    item_settings = sub(manifest, "opf", "item")
    item_settings.set("id", "settings")
    item_settings.set("href", "settings.xml")
    item_settings.set("media-type", "application/xml")

    # spine
    spine = sub(root, "opf", "spine")
    sub(spine, "opf", "itemref", {"idref": "header", "linear": "yes"})
    for i in range(section_count):
        sub(spine, "opf", "itemref", {"idref": f"section{i}", "linear": "yes"})

    return to_xml_bytes(root)


def _meta(parent, name, value):
    """Create an opf:meta element."""
    m = sub(parent, "opf", "meta")
    m.set("name", name)
    m.set("content", "text")
    m.text = value
    return m


def _build_settings_xml():
    """Build settings.xml."""
    nsmap = {
        "ha": NS["ha"],
        "config": NS["config"],
    }
    root = root_element("ha", "HWPApplicationSetting", nsmap=nsmap)
    caret = sub(root, "ha", "CaretPosition")
    caret.set("listIDRef", "0")
    caret.set("paraIDRef", "0")
    caret.set("pos", "16")
    return to_xml_bytes(root)


def _write_hwpx_zip(files, output_path):
    """Write HWPX files dict to a ZIP file."""
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        # mimetype should be first and uncompressed
        if "mimetype" in files:
            zf.writestr("mimetype", files["mimetype"], compress_type=zipfile.ZIP_STORED)

        for path, data in sorted(files.items()):
            if path == "mimetype":
                continue
            if isinstance(data, str):
                data = data.encode("utf-8")
            # BinData and Preview images should be stored uncompressed
            # (matches Hancom reference HWPX behavior)
            if path.startswith("BinData/") or path.startswith("Preview/"):
                zf.writestr(path, data, compress_type=zipfile.ZIP_STORED)
            else:
                zf.writestr(path, data)
