"""Validation tests for BinData ↔ manifest ↔ section XML consistency.

Ensures generated HWPX never references non-existent BinData entries.
"""

import os
import re
import zipfile
import tempfile
from lxml import etree

import sys
sys.path.insert(0, os.path.dirname(__file__))

from hwp2hwpx.converter import convert_file


# Namespace map for parsing HWPX XML
NS = {
    "opf": "http://www.idpf.org/2007/opf/",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
}


def _extract_hwpx_info(hwpx_path):
    """Extract manifest image IDs, BinData files, and section binaryItemIDRefs."""
    manifest_ids = set()
    bindata_files = set()
    section_refs = set()

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        names = zf.namelist()

        # Collect actual BinData files in the ZIP
        for name in names:
            if name.startswith("BinData/"):
                # Extract the label part (e.g., "image1" from "BinData/image1.png")
                base = os.path.splitext(os.path.basename(name))[0]
                bindata_files.add(base)

        # Parse content.hpf manifest
        hpf_data = zf.read("Contents/content.hpf")
        hpf_tree = etree.fromstring(hpf_data)
        for item in hpf_tree.iter("{%s}item" % NS["opf"]):
            item_id = item.get("id", "")
            href = item.get("href", "")
            if "BinData/" in href:
                manifest_ids.add(item_id)

        # Parse all section XMLs for binaryItemIDRef
        for name in names:
            if re.match(r"Contents/section\d+\.xml", name):
                sec_data = zf.read(name)
                sec_tree = etree.fromstring(sec_data)
                for img in sec_tree.iter("{%s}img" % NS["hc"]):
                    ref = img.get("binaryItemIDRef", "")
                    if ref:
                        section_refs.add(ref)

    return manifest_ids, bindata_files, section_refs


def validate_hwpx(hwpx_path):
    """Validate BinData consistency in an HWPX file. Returns list of errors."""
    manifest_ids, bindata_files, section_refs = _extract_hwpx_info(hwpx_path)
    errors = []

    # Check 1: Every manifest entry must have a corresponding BinData file
    phantom_manifest = manifest_ids - bindata_files
    if phantom_manifest:
        errors.append(f"Manifest references missing BinData files: {sorted(phantom_manifest)}")

    # Check 2: Every section binaryItemIDRef must exist in manifest
    phantom_refs = section_refs - manifest_ids
    if phantom_refs:
        errors.append(f"Section XML references missing manifest entries: {sorted(phantom_refs)}")

    # Check 3: Every section binaryItemIDRef must have actual BinData file
    phantom_files = section_refs - bindata_files
    if phantom_files:
        errors.append(f"Section XML references missing BinData files: {sorted(phantom_files)}")

    return errors


def test_all_hwp_files():
    """Regression test: convert all test/*.hwp and validate BinData consistency."""
    test_dir = os.path.join(os.path.dirname(__file__), "test")
    if not os.path.isdir(test_dir):
        print("SKIP: test/ directory not found")
        return

    passed = 0
    failed = 0
    errors_found = []

    for entry in sorted(os.listdir(test_dir)):
        hwp_path = os.path.join(test_dir, entry, "from.hwp")
        if not os.path.exists(hwp_path):
            continue

        with tempfile.NamedTemporaryFile(suffix=".hwpx", delete=False) as tmp:
            tmp_path = tmp.name

        try:
            convert_file(hwp_path, tmp_path)
            errors = validate_hwpx(tmp_path)
            if errors:
                failed += 1
                errors_found.append((entry, errors))
                print(f"  FAIL  {entry}")
                for e in errors:
                    print(f"        {e}")
            else:
                passed += 1
                print(f"  OK    {entry}")
        except Exception as ex:
            failed += 1
            errors_found.append((entry, [str(ex)]))
            print(f"  ERROR {entry}: {ex}")
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    print(f"\nResults: {passed} passed, {failed} failed out of {passed + failed}")
    return failed == 0


def test_picture_case():
    """Specific test for the picture test case with known images."""
    hwp_path = os.path.join(os.path.dirname(__file__), "test", "picture", "from.hwp")
    if not os.path.exists(hwp_path):
        print("SKIP: test/picture/from.hwp not found")
        return True

    with tempfile.NamedTemporaryFile(suffix=".hwpx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        convert_file(hwp_path, tmp_path)
        manifest_ids, bindata_files, section_refs = _extract_hwpx_info(tmp_path)
        print(f"  Manifest IDs:    {sorted(manifest_ids)}")
        print(f"  BinData files:   {sorted(bindata_files)}")
        print(f"  Section refs:    {sorted(section_refs)}")

        errors = validate_hwpx(tmp_path)
        if errors:
            print("  FAIL: " + "; ".join(errors))
            return False
        else:
            print("  OK: All references valid")
            return True
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


if __name__ == "__main__":
    print("=== Picture case validation ===")
    test_picture_case()
    print()
    print("=== Full regression test ===")
    success = test_all_hwp_files()
    sys.exit(0 if success else 1)
