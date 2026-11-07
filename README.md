# openpyxl-365

Edit Excel 365 workbooks with the familiar `openpyxl` API while preserving
selected OOXML features that ordinary load/save operations can lose.

[中文](README.zh-CN.md) · [Changelog](CHANGELOG.md)

## Install

Python 3.10+:

```bash
pip install git+https://github.com/chengjiaxiao/openpyxl-365.git
```

## Use

```python
from openpyxl365 import Excel365Workbook

doc = Excel365Workbook.open("model.xlsm")
doc.wb["Parameters"]["D11"] = 44_000_000_000
doc.set_dynamic_formula("Summary", "H2", '=FILTER(A2:D100,D2:D100<>"")')
doc.save("model-updated.xlsm")
```

Use `doc.wb` for ordinary cell, style and formula edits. Recognized dynamic
formulas such as `=UNIQUE(A2:A100)` are also handled on direct assignment.
Use `Excel365Workbook.create()` to start a new workbook.

## What it preserves

- Dynamic arrays, modern function prefixes, spill references and unchanged formula caches.
- VBA, external-link parts, Rich Data and existing Python-in-Excel payloads.
- x14 dropdowns, threaded comments and unsupported DrawingML shapes, including macro buttons.
- Rich-text whitespace when opened with `rich_text=True`.

Changed array anchors have their old spill caches cleared; explicit edits to
former spill cells are retained. HTTPS workbook links can be inspected, relocated
or created with `external_link_targets()`, `set_external_link_target()` and
`set_external_reference()`. These methods do not verify access or refresh data.

Saves stage in a local cache directory and atomically replace the destination.
Set `EXCEL365_STAGING_DIR` to override the staging directory.

## Limits

- **No formula calculation.** Recalculate and save in Microsoft Excel before
  relying on results. Saving removes the calculation chain; the default no
  longer writes forced full-recalculation flags. Pass
  `recalculate_on_open=True` to explicitly request full calculation on load.
- Preservation does not imply support for editing every advanced object.
  Threaded-conversation editing, ActiveX and form controls are not guaranteed.
- Sheet deletion and other structural changes require additional validation.
  Workbook modifications can invalidate digital signatures.

## Development

```bash
pip install -e .
python -m unittest discover -s tests -p "test_*.py"
```

See [CONTRIBUTING.md](CONTRIBUTING.md) and the optional Windows Excel validation
harness in `tests/validate_in_excel.ps1`.

MIT · Cheng Jiaxiao, with GPT (OpenAI) as collaborative author.
Independent of Microsoft and the openpyxl project.
