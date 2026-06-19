# Range Metadata Reads

Range metadata reads expose workbook details that the backend may use before editing a workbook. Normal MCP clients should request these through `excel.agent.run` with `mode: "answer"` and a concrete target; the range-specific capabilities remain backend-owned.

## Implemented

- Hyperlink metadata from `Range.hyperlink`
- Merged areas from `Range.getMergedAreasOrNullObject`
- Data validation metadata from `Range.dataValidation`
- Conditional formatting metadata from `Range.conditionalFormats`
- Search matches from worksheet `findAllOrNullObject`
- Blank cells from special cells
- Formula errors from special cells

## Limitations

Comments and legacy notes return explicit unsupported warnings for now. Office.js exposes workbook and worksheet comment collections, but this runtime still needs reliable range-address mapping before enabling agent-facing tools.

Worksheet search currently ignores `searchDirection` because Office.js worksheet search criteria do not support it. The protocol keeps the field so a future range-level search implementation can honor it.
