# Style Fidelity

Open Workbook treats styling as a first-class invariant for template-driven workbooks.

## Lifecycle

1. Capture a style fingerprint with `excel.style.get_fingerprint`.
2. Compare a template/source sheet against a target with `excel.style.compare_fingerprint`.
3. Copy one style dimension, or repair all template styles, with a backup created before mutation.
4. Validate the target after copy so agents can decide whether to continue, repair, or roll back.

## Covered Dimensions

- Column widths and row heights are read and replayed explicitly.
- Fills, fonts, borders, alignment, number formats, conditional formatting, and data validation use Excel native format copy for fast high-fidelity replay.
- Direct style writes support fill color, font name/size/color/bold/italic, horizontal and vertical alignment, row height, and column width.
- Fingerprints include bounded cell-level samples for fills, fonts, alignment, and number formats.

## Performance

`excel.style.get_fingerprint` accepts `maxCellSamples`. Large ranges still capture range-level and row/column layout data, but skip per-cell style sampling once the limit is exceeded. This keeps daily office tasks responsive while still allowing detailed checks on headers, report blocks, and template regions.

Style copy operations are batched through Office.js and backed up before mutation. Agents should prefer copying a specific dimension or template region instead of reading and writing each cell individually.

## Current Limits

Excel does not have web-style padding or margins for cells; the closest controllable layout properties are row height, column width, wrapping, shrink-to-fit, indentation, alignment, and borders.

Workbook theme replay, freeze panes, print settings, page layout, and hidden row/column replay are exposed as explicit capability-status paths until we can capture and replay them deterministically across supported Excel hosts.
