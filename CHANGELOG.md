# Changelog

All notable changes to this project will be documented in this file.

## 0.2.0 - 2026-11-07

- Preserve Python-in-Excel payloads and rich-text whitespace.
- Restore shared-formula groups together; expand the group when a member changes.
- Resolve shared-string formula caches and empty formula spill children correctly.
- Clear stale array metadata and spill caches after replacing an anchor, while
  retaining explicit edits to former spill cells.
- Add HTTPS external-workbook reference creation, inspection and relocation.
- Stage saves outside synced folders, with atomic cross-volume publication.
- Report calculation properties as written in the source file and handle
  workbooks without a calculation-properties element.
- **Changed default:** `save()` now defaults to `recalculate_on_open=False`
  and omits calcMode/fullCalcOnLoad/forceFullCalc overrides. Pass `True` to
  explicitly write fullCalcOnLoad. Formula calculation still requires Excel.
- Simplify English and Chinese READMEs and extend synthetic regression coverage.

## 0.1.1 - 2026-08-30

- Add GPT (OpenAI) as a collaborative author in package metadata and project
  documentation.

## 0.1.0 - 2026-08-30

Initial open-source release.

- Add Excel 365 dynamic-array formula creation and round-trip preservation.
- Normalize future-function prefixes and spill references.
- Preserve formula metadata and selected cached results.
- Preserve VBA, external links, Rich Data, x14 data validations, threaded
  comments, and DrawingML package trees.
- Add atomic saves, OOXML inspection, regression tests, bilingual
  documentation, and an optional Microsoft Excel validation harness.
