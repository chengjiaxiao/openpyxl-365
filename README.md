# openpyxl-365

Modern Excel workbooks are complex OOXML packages. A single `.xlsx` or `.xlsm`
file may contain many interdependent XML documents, relationship files,
metadata records, drawings, and binary parts. `openpyxl` only understands part
of that package. When it loads and saves a workbook, unsupported parts or
attributes can be removed or rewritten even if you only changed one cell.

The result can be broken dynamic arrays, rewritten external-link caches,
missing data-validation dropdowns, lost threaded comments, damaged Rich Data,
or deleted macro-bound shapes.

`openpyxl-365` lets you keep using the familiar `openpyxl` API while preserving
these modern Excel 365 OOXML parts across an edit/save round trip.

## How it works

The project uses `openpyxl` for normal cell, formula, style, and worksheet
editing. Before saving, it records selected OOXML content that `openpyxl` does
not round-trip safely. After `openpyxl` writes the workbook, those parts,
relationships, attributes, and content-type declarations are restored.

This matters because modern formulas and workbook features are distributed
across formula text, cell metadata, relationships, content types, and extension
parts. Writing `=FILTER(...)` into a cell is only one piece of the job.

## Features

- Creates and round-trips Excel 365 dynamic-array formulas.
- Adds required `_xlfn.` / `_xlfn._xlws.` prefixes for supported modern
  functions.
- Converts formula-bar spill references such as `D1#` to the stored
  `_xlfn.ANCHORARRAY(D1)` representation.
- Preserves existing dynamic-array metadata, formula attributes, cached spill
  children, and rich-value cell metadata.
- Preserves external-link package parts byte-for-byte instead of letting
  `openpyxl` rewrite their cached data.
- Preserves VBA when editing `.xlsm` files.
- Preserves Rich Data package parts and their workbook relationships.
- Preserves Excel 365 x14 extended data validations.
- Preserves existing threaded comments, people metadata, and compatibility
  VML without creating duplicate legacy comments.
- Preserves unsupported DrawingML trees, including shapes and macro-bound
  buttons.
- Saves atomically and can request recalculation when Excel next opens the
  workbook.

## Installation

Until a PyPI release is published:

```bash
pip install git+https://github.com/chengjiaxiao/openpyxl-365.git
```

## Quick start

```python
from openpyxl365 import Excel365Workbook

doc = Excel365Workbook.open("model.xlsm")

# Use the normal openpyxl API for values, styles, and existing formulas.
doc.wb["Parameters"]["D11"] = 44_000_000_000

# Use set_dynamic_formula for a new or replaced spilling formula.
doc.set_dynamic_formula(
    "Summary",
    "H2",
    '=FILTER(A2:D100,D2:D100<>"")',
)

doc.save("model-updated.xlsm", recalculate_on_open=True)
```

Direct assignments of recognized dynamic functions are also upgraded during
save:

```python
doc.wb["Summary"]["J2"] = "=UNIQUE(A2:A100)"
```

Inspect the saved OOXML signals without opening Excel:

```python
from openpyxl365 import inspect_excel365_dynamic_arrays

report = inspect_excel365_dynamic_arrays("model-updated.xlsm")
assert report["valid"], report["errors"]
```

## Supported modern functions

The current normalizer recognizes common spill-capable functions including
`FILTER`, `UNIQUE`, `SORT`, `SORTBY`, `SEQUENCE`, `RANDARRAY`, `XLOOKUP`,
`VSTACK`, `HSTACK`, `TAKE`, `DROP`, `TOCOL`, `TOROW`, `MAP`, `REDUCE`, `SCAN`,
`BYROW`, `BYCOL`, `MAKEARRAY`, `LET`, and `LAMBDA`.

## Important limitations

- This project does **not** calculate formulas. Cached values remain stale
  until a compatible spreadsheet application recalculates the workbook.
- Preservation means existing unsupported parts are retained; it does not mean
  they can all be created or edited through Python.
- Existing threaded comments are preserved, but editing threaded conversations
  is not supported.
- DrawingML shapes and macro bindings are preserved. ActiveX and form controls
  use additional OOXML structures and are not yet guaranteed.
- Broad structural changes such as deleting sheets that own preserved package
  parts may require additional relationship cleanup.
- Digital signatures can be invalidated by any workbook modification.

## Verification

The automated tests cover dynamic arrays, spill references, shared-formula
followers, cached results, external links, DrawingML, x14 data validation,
Rich Data, and threaded comments. A PowerShell harness is included for optional
end-to-end validation in a running Windows Microsoft Excel session.

```bash
python -m unittest discover -s tests -p "test_*.py"
```

## Contributions

The project has a deliberately narrow goal: make selected modern Excel 365
features survive an `openpyxl` edit/save round trip. Reproducible fixtures,
OOXML edge cases, and compatibility reports from different Excel versions are
especially welcome. See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT. This project is independent and is not affiliated with the `openpyxl`
project or Microsoft.

中文说明见 [README.zh-CN.md](README.zh-CN.md)。
