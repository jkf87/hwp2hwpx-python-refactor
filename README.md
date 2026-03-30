# hwp2hwpx-python-refactor

Pure Python HWP → HWPX converter with active Hancom-compatibility and fidelity work.

This project started from the original Java-based `hwp2hwpx` lineage, but this repository is now a **separate Python-first project** focused on:
- removing Java runtime/library dependency
- generating HWPX that actually opens in Hancom Hangul
- improving visual fidelity against original HWP rendering

> Upstream inspiration: [neolord0/hwp2hwpx](https://github.com/neolord0/hwp2hwpx)
> 
> This repository is **not just a mirror**. It is an independent Python refactor and compatibility/fidelity effort.

## Status

Current project state:
- Pure Python converter implemented
- Hancom-openability restored for real-world samples
- Ongoing work focused on **PDF-render fidelity** between:
  1. original `HWP → PDF`
  2. converted `HWPX → PDF`

Recent work includes:
- packaging compatibility fixes
- FileHeader version byte-order fix for Hancom openability
- BinData decompression handling
- `tcps` control-char handling improvements
- table/background fill fidelity improvements
- continued special-control and rendering-gap analysis

## Features

Implemented or substantially covered:
- paragraph text and character properties
- paragraph properties
- tables, merged cells, borders, fills
- section/page properties
- fonts, styles, tabs, numbering
- BinData/image packaging
- Hancom-oriented HWPX packaging structure
- header/footer support
- footnote/endnote body support
- substantial GSO/image support

Still being improved:
- page-number and special controls (`pgnp`, `pghd`, `nwno`, `tcps`)
- remaining GSO fidelity details
- field controls
- equations / other renderables
- PDF visual equivalence against complex real documents

## Installation

```bash
pip install -r requirements.txt
```

Main dependencies:
- `pyhwp`
- `olefile`
- `lxml`

## Usage

### CLI

```bash
python3 -m hwp2hwpx input.hwp
python3 -m hwp2hwpx input.hwp -o output.hwpx
python3 -m hwp2hwpx input.hwp -v
```

### Python API

```python
from hwp2hwpx import convert_file

output_path = convert_file("input.hwp", "output.hwpx")
```

Or convert into an in-memory file map:

```python
from hwp2hwpx import convert
from hwp2hwpx.reader import HWPReader

with HWPReader("input.hwp") as reader:
    files = convert(reader)  # dict[path, bytes]
```

## Project layout

```text
hwp2hwpx/
├── __init__.py
├── __main__.py
├── converter.py
├── reader.py
├── header_converter.py
├── section_converter.py
├── xml_builder.py
└── value_maps.py
```

## Conversion flow

1. `HWPReader` parses HWP record/model streams using `pyhwp` + `olefile`
2. `header_converter.py` builds `header.xml`
3. `section_converter.py` builds `section*.xml`
4. `converter.py` packages the final HWPX ZIP structure

## Fidelity policy

The practical success criteria for this repository are:
1. generated HWPX must open in Hancom Hangul
2. rendered output should match the original HWP as closely as possible
3. PDF output similarity is the final quality bar

That means this project optimizes for **real Hancom behavior and rendered results**, not just XML validity or Java parity.

## Development notes

Typical workflow used in this repo:
- convert real HWP samples
- open generated HWPX in Hancom
- export/compare PDFs
- isolate one highest-impact visual delta at a time
- patch converter
- update `WORKLOG.md`
- commit incremental improvements

## Testing

Run tests:

```bash
python3 -m pytest
```

Real-world validation is also done with sample `.hwp` files outside the unit test set.

## License

Apache-2.0

This project inherits ideas and file-format understanding from the original Java ecosystem, while the Python refactor and ongoing compatibility/fidelity work in this repository are maintained separately.
