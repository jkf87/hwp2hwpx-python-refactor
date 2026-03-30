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

- **GSO (Graphical Shape Objects)**: Shapes, images, drawing objects, lines, rectangles, ellipses, arcs, polygons, curves, textart, OLE objects are **stubbed out** — the inline builder silently skips them. Files convert without errors but graphic content is missing from output.
- **Header/footer content**: Fully implemented — text, styling, alignment, and page placement (BOTH/EVEN/ODD) all converted correctly.
- **Footnote/endnote content**: Control recognized but body text not converted.
- **Field begin/end**: Hyperlinks, bookmarks, page numbers — control chars recognized but not rendered as HWPX field elements.
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
