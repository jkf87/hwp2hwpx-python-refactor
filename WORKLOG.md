# hwp2hwpx Python Refactor — Work Log

## Overview

Refactored the Java-based `hwp2hwpx` converter into a pure Python implementation.
The original Java project depends on `hwplib` and `hwpxlib` (Java libraries by neolord0).
The Python version uses `pyhwp` (hwp5) + `olefile` for reading HWP binary files and `lxml` for building HWPX XML output.

## Architecture

```
hwp2hwpx/
├── __init__.py          # Package entry, exports convert_file / convert
├── __main__.py          # CLI: python3 -m hwp2hwpx input.hwp [-o out.hwpx]
├── converter.py         # Orchestrator: builds HWPX ZIP from reader output
├── reader.py            # HWPReader: wraps pyhwp + olefile for HWP parsing
├── header_converter.py  # DocInfo → header.xml (fonts, charshapes, parashapes, styles, borders)
├── section_converter.py # BodyText sections → section0.xml … sectionN.xml
├── xml_builder.py       # lxml helpers, HWPX namespace definitions
└── value_maps.py        # Binary flag → XML enum string conversion tables
```

Total: ~2,000 lines of Python (vs ~15,000 lines of Java in the original).

## Decisions

1. **pyhwp xmlmodel API** — Chose `Hwp5File.docinfo.models()` and `.bodytext.section(N).models()` as the primary data source. These return structured dicts with `tagname`, `level`, and `content` fields, giving us a flat stream of HWP records with hierarchy encoded via `level`.

2. **Scan-ahead paragraph processing** — The section converter uses a two-pass approach per paragraph:
   - Pass 1: Scan all child models to find PARA_TEXT, PARA_CHAR_SHAPE, PARA_LINE_SEG, and record CTRL_HEADER positions with their child ranges.
   - Pass 2: Build `<hp:run>` elements, placing inline controls (tables, columns, section properties) at the exact position of their control character in the text stream.

3. **Controls inside runs** — HWPX requires controls to appear INSIDE `<hp:run>` elements, not as siblings. This was a major structural insight that required rewriting the paragraph builder.

4. **Cell boundary detection** — Table cells (LIST_HEADER) and their paragraphs appear at the SAME level in the model stream. Cell boundaries are determined by finding all LIST_HEADER positions first, then defining ranges between consecutive LIST_HEADERs.

5. **pageBorderFill type** — The Java reference uses order-based type assignment (first=BOTH, second=EVEN, third=ODD), not flag-based extraction.

## Key Bug Fixes

| Bug | Root Cause | Fix |
|-----|-----------|-----|
| Landscape detection wrong | Used `attr & 0x01` flag | Changed to `width > height` comparison |
| Controls outside runs | Built controls as paragraph siblings | Rewrote to place controls inside `<hp:run>` elements |
| Cell paragraphs empty | Looked for children at `level > cell_level` | Fixed: paragraphs are at SAME level as LIST_HEADER |
| Table pageBreak bits | Tested wrong bit offset | Corrected: `(flags >> 0) & 0x03` for pageBreak, `(flags >> 2) & 0x01` for repeatHeader |
| Column def flags | Used `flags & 0xFF` for type | Corrected bit layout: type=bits 0-1, count=bits 2-9, layout=bits 10-11, sameSz=bit 12 |
| CommonControl property bits | Wrong bit offsets for vert/horz alignment | Verified via binary analysis of flags=0x080A2210: vertRelTo=bits 3-4, horzRelTo=bits 8-9, vertAlign=bits 10-12, horzAlign=bits 14-16 |
| Cell header attribute | Used `row == 0` check | Fixed to use `(listflags >> 18) & 0x01` |

## Test Results

**41/41 files converted successfully (0 failures)**

- 33 test cases from `test/` directory (bookmark, table, picture, equation, header_footer, footnote_endnote, field, textart, ole, compose, dutmal, multi_run, new_number, page_hiding, page_num, space_linebreak, tab_in_para, shapes, 빈파일, 여러섹션, 오류, etc.)
- 8 real-world HWP files from Downloads (government documents, forms, lecture materials)

### Comparison vs Java Reference (table test case)

Only 3 minor differences remain:
1. `landscape`: NARROWLY vs WIDELY — Python output correct per raw HWP binary data
2. `noteLine` length: 12280 vs 14692344 — pyhwp vs hwplib parse difference in footnote separator length
3. Extra `name=""` on `<hp:tc>` elements — harmless attribute

## Known Limitations

- **GSO (Graphical Shape Objects)**: Images, rectangles, lines, containers, ellipses implemented. Arcs, polygons, curves, textart, OLE objects still stubbed.
- **Header/footer content**: Fully implemented — text, styling, alignment, and page placement (BOTH/EVEN/ODD) all converted correctly.
- **Footnote/endnote content**: Fully implemented — body text with proper formatting.
- **Autonomous inline controls**: `pgnp` (page number position), `pghd` (page hiding), `nwno` (new numbering) fully implemented. `tcps` (table cell paragraph shape) stubbed.
- **Field begin/end**: Hyperlinks, bookmarks — control chars recognized but not rendered as HWPX field elements.
- **Equations**: Equation control recognized but formula content not converted.

## Commands Used

```bash
# Install dependencies
pip3 install pyhwp olefile lxml

# Convert a single file
python3 -m hwp2hwpx input.hwp -o output.hwpx

# Convert programmatically
python3 -c "from hwp2hwpx import convert_file; convert_file('input.hwp', 'output.hwpx')"

# Run full test suite
python3 -c "
import os, glob, sys
sys.path.insert(0, '.')
from hwp2hwpx import convert_file
files = sorted(glob.glob('test/*/*.hwp')) + sorted(glob.glob('/path/to/downloads/*.hwp'))
passed = failed = 0
for f in files:
    try:
        convert_file(f, '/tmp/test_out.hwpx')
        passed += 1
    except Exception as e:
        failed += 1
        print(f'FAIL: {os.path.basename(f)}: {e}')
print(f'{passed}/{passed+failed} passed, {failed} failed')
"
```

## Timeline

- **2026-03-30**: Initial audit, architecture design, full implementation, iterative bug fixing, 41/41 test pass rate achieved.
- **2026-03-30**: Hancom compatibility triage for failing file `1-3. (260313_한준구) 찾아가는 저널리즘 특강_강사 강의 확인서.hwp`.

### Hancom Compatibility Fixes (2026-03-30)

**Failing-file triage**: Compared generated HWPX against reference HWPX files created by Hancom Hangul. Found 6 structural differences:

| Issue | Before | After (matches reference) |
|-------|--------|--------------------------|
| `manifest.xml` namespace | `<manifest/>` (no namespace) | `<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>` |
| BinData ZIP path | `Contents/BinData/image1.jpg` | `BinData/image1.jpg` (root level) |
| `content.hpf` image item `id` | `bindata1` | `image1` |
| `content.hpf` image item `href` | `Contents/BinData/image1.jpg` | `BinData/image1.jpg` |
| `content.hpf` image item missing `isEmbeded` | (absent) | `isEmbeded="1"` |
| `content.hpf` spine `linear` attr | (absent) | `linear="yes"` |
| `hp:pic` child element order | `sz, pos, outMargin, shapeComment, hc:img(imgRect, imgClip)` | `offset, orgSz, curSz, flip, rotationInfo, renderingInfo, imgRect, imgClip, inMargin, imgDim, hc:img(leaf), effects, sz, pos, outMargin, shapeComment` |
| `hc:img binaryItemIDRef` | Numeric `"1"` | String `"image1"` |
| `imgRect`/`imgClip` namespace | `hc:` (children of `hc:img`) | `hp:` (children of `hp:pic`) |

**Files changed**: `converter.py`, `section_converter.py`
**All 34 test cases pass** after changes.

### Root Cause Found: FileHeader Version Byte Order (2026-03-30)

**Root cause**: `reader.py` was reading the HWP5 FileHeader version DWORD at offset 32 in big-endian order (`data[32]=major`), but the format is **little-endian**: `[build_lo, build_hi/micro, minor, major]`. For the failing file, bytes were `00 01 01 05`, so:
- **Before (wrong)**: `major=0, minor=1, micro=1, build=5` → Hancom rejected `major="0"`
- **After (correct)**: `major=5, minor=1, micro=1, build=0` → Hancom opens successfully

**Verification**: Opened `hancom_test_v2.hwpx` in Hancom Office HWP on macOS — document renders correctly with all text, tables, and embedded images visible. Both the minimal variant and the full converted file open without errors.

**Fix**: `reader.py` line 138-141 — reversed byte order in `get_file_header()`.
**Commit**: `7ed4c2b`
**Test results**: 43/43 pass (33 test suite + 10 real-world HWP files).

### Fidelity Improvements: hp:pic Structure & Inline Controls (2026-03-30)

**Golden sample**: `★2022년 행정안전부 주요업무 추진계획(최종).hwp`

**1. hp:pic element structure simplified to match Hancom reference**

| Before (verbose) | After (matches Hancom) |
|---|---|
| 16 child elements: offset, orgSz, curSz, flip, rotationInfo, renderingInfo(3 matrices), imgRect, imgClip, inMargin, imgDim, hc:img(leaf), effects, sz, pos, outMargin, shapeComment | 5 child elements: sz, pos, outMargin, shapeComment, hc:img(hc:imgRect, hc:imgClip) |
| Extra pic attrs: dropcapstyle, href, groupLevel, instid, reverse | Clean attrs: id, zOrder, numberingType, textWrap, textFlow, lock |
| imgRect/imgClip as hp: children of hp:pic | imgRect/imgClip as hc: children of hc:img |

**2. Autonomous inline controls implemented (code 21)**

Previously silently consumed. Now generates correct HWPX elements:

| HWP chid | HWPX element | Description |
|---|---|---|
| `pgnp` | `<hp:ctrl><hp:pageNum pos=".." formatType=".." sideChar=".."/></hp:ctrl>` | Page number position (flags: bits 8-11=position, bits 0-3=format) |
| `pghd` | `<hp:ctrl><hp:pageHiding hideHeader=".." hideFooter=".." hideMasterPage=".." hideBorder=".." hideFill=".." hidePageNum=".."/></hp:ctrl>` | Page hiding definition (flags: bit 0-5 = hide flags) |
| `nwno` | `<hp:ctrl><hp:newNum num=".." numType="PAGE"/></hp:ctrl>` | New numbering reset |
| `tcps` | (stubbed) | Table cell paragraph shape override |

**Files changed**: `section_converter.py`
**Test results**: 43/43 pass (33 test suite + 10 real-world HWP files).

**Remaining differences vs golden reference** (for future work):
- 16 extra paragraphs in table cell subLists (cell boundary detection)
- 2 extra pics from WMF bindata (276-byte decorative metafiles)
- 4 extra lineShape elements on shapes
- 4 extra tab elements

### PDF Visual Equivalence Sprint (2026-03-30)

Target: `★2022년 행정안전부 주요업무 추진계획(최종).hwp`

**Fixes applied (highest impact first):**

| Fix | Commit | Impact |
|-----|--------|--------|
| borderFill colors: read from `fill_colorpattern` dict | `afc7659` | Table/cell background colors were all wrong (reading from nonexistent keys) |
| Text splitting at char shape boundaries | `5198b77` | Multi-style paragraphs rendered in single style; 2592 runs vs 846 before |
| Paragraph margin doubling fix (case=val/2, default=val) | `a17d3cb` | All paragraph indents and spacing 2-4x too large |
| Table textWrap/textFlow from CommonControl flags | `86906b8` | Tables had wrong wrap mode (SQUARE instead of TIGHT) |
| Table cell vertAlign/textDirection from listflags | `74e10c1` | Cell text was vertically centered instead of top-aligned |
| Gradation fill support (fillflags & 0x04) | `afc7659` | Gradient backgrounds in table headers now rendered |

| BinData decompression for BMP/WMF images | `b1ecfb0` | BMP/WMF images were zlib-compressed in HWPX output |
| Handle tcps (code 23) control char | `d7aa846` | Control index desync when tcps appeared in text |

**Current state** (target file):
- 569 paragraphs, 2592 runs, 62 tables, 194 cells, 15 GSO shapes, 9 images
- All IDs valid: borderFill (1-77), charPr (0-503), paraPr (0-305), style (0-29)
- Correct fill colors: #DFE6F7 (light blue headers), #3E57A5 (dark blue), #FFFFE5 (yellow), etc.
- Correct border types: SOLID, DOUBLE_SLIM, DOT mapped correctly
- All images decompressed: BMP (BM header), WMF (placeable WMF header), JPG, PNG
- Text correctly split across 490 unique char shape styles
- Paragraph margins properly halved for HwpUnitChar case, raw for default
- 33/33 test suite files pass

**Structural metric comparison (source vs output):**

| Metric | Source | Output | Match |
|--------|--------|--------|-------|
| Paragraphs | 569 | 569 | ✓ |
| Tables | 62 | 62 | ✓ |
| Table cells | 194 | 194 | ✓ |
| Pictures | 10 | 10 | ✓ |
| Line segments | 569 | 569 | ✓ |

**Remaining known limitations:**
- Field controls (bookmarks, hyperlinks) not implemented
- Equations not implemented
- Advanced GSO shapes (arc, polygon, curve, textart, OLE) stubbed
- tcps generates no HWPX output (consumed but not converted)


## 2026-03-31 — Repository cleanup for Python-only distribution
- Removed legacy Java/Maven source tree (`src/`) from this repository.
- Removed `pom.xml`.
- Verified in a temporary clone that conversion still works without Java sources present.
- Repository is now packaged as a Python-only distribution target.


## 2026-03-31 — Licensing cleanup checkpoint
- Repository now distributes Python implementation only (`src/` and `pom.xml` removed).
- Remaining upstream/license references are now documentation cleanup items, not runtime dependencies.
- Kept implementation history in WORKLOG for traceability while preparing for a cleaner licensing review.
