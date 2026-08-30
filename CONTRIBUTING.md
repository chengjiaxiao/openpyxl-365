# Contributing

Thank you for helping modern Excel workbooks survive Python automation.

## Good contributions

- A minimal workbook fixture that reproduces a round-trip failure.
- A test that identifies the exact OOXML part, relationship, or attribute lost.
- Compatibility results from a named Excel version and operating system.
- Focused preservation support that does not rewrite unrelated package data.

Do not submit confidential, customer-owned, licensed-data, or personally
identifying workbook content. Reduce fixtures to synthetic values first.

## Development

Create a virtual environment outside any synced project directory, install the
package in editable mode, and run:

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Please keep changes narrowly scoped and add a regression test for every OOXML
preservation fix.

## Reporting an issue

Include:

1. Python and `openpyxl` versions.
2. Excel version and operating system, if Excel validation was performed.
3. Whether the file is `.xlsx` or `.xlsm`.
4. The exact load/edit/save operation.
5. A synthetic reproduction file when possible.

Never upload a sensitive workbook to a public issue.
