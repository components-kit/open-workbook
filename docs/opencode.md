# OpenCode Example

OpenCode is one possible MCP client. The generic Open Workbook install flow is still:

```bash
npx -y @components-kit/open-workbook setup
```

OpenCode uses its own MCP config shape. Generate an OpenCode-shaped snippet from the local CLI with:

```bash
owb opencode config --id open-workbook --agent-name finance-agent
```

Example output:

```json
{
  "mcp": {
    "open-workbook": {
      "type": "local",
      "command": ["owb", "mcp", "--agent-name", "finance-agent"],
      "enabled": true
    }
  }
}
```

The generated OpenCode snippet uses the public `excel.agent.run` workflow surface. Agents should call only `excel.agent.run`; the backend handles workbook discovery, target resolution, reads, previews, applies, validation, and compact proof internally. The primitive operation catalog is backend/test capability, not a normal OpenCode tool surface. Agents should not fall back to Python, openpyxl, pandas, shell scripts, or offline `.xlsx` parsing for a connected live workbook unless the user explicitly asks for offline file analysis or approves a non-live fallback after MCP is unavailable.

For `npx`-based OpenCode config, keep the command equivalent to:

```json
{
  "mcp": {
    "open-workbook": {
      "type": "local",
      "command": ["npx", "-y", "@components-kit/open-workbook@latest", "mcp"],
      "enabled": true
    }
  }
}
```

Install the Open Workbook Skills skill into OpenCode:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-skills -a opencode -g -y
```

After Excel opens the add-in, useful default-surface calls are:

```text
excel.agent.run mode=status request="Check Open Workbook status"
excel.agent.run mode=prepare request="Prepare workbook context"
excel.agent.run mode=find request="Find the sheet or table I need"
excel.agent.run mode=answer detailLevel=workbook_summary request="Summarize the active workbook"
excel.agent.run mode=answer request="Compare January and February"
excel.agent.run mode=auto intent.action=write_values target.sheetName=Sales target.range=E2 values.values=[["Reviewed"]] request="Change Sales!E2 to Reviewed"
excel.agent.run mode=apply_update request="Apply the previewed update"
excel.agent.run mode=operation_status request="Check a pending operation" operationId="..."
```

For workbook overview prompts such as "what is this file?", "look into this workbook", or "summarize the workbook", use `detailLevel=workbook_summary` or `detailLevel=sheet_summary` and stop when the response says `nextAction=answer_now` or `maxRecommendedFollowupCalls=0`. Do not fetch `fullResultUri`, chunk-read sheets, list MCP resources, or call low-level resource reads unless the user explicitly asks for all raw rows or exact cell values.

For small exact value edits the user already requested, prefer `mode=auto` with `intent.action=write_values`, explicit `target`, and structured `values`. Safe scoped edits can return `taskOutcome=apply_complete` in one call; report the proof and stop. Use `preview_update` only when the user asks to review first, the target/value is ambiguous, or the edit is broad, formula/style/table/structural, or otherwise risky.

For dropdown option questions, call `intent.action=read_data_validation` once on the selected/current column or exact target. When the answer kind is `data_validation_summary`, answer from the inline validation metadata/options and stop; do not fetch `fullResultUri`, chunk-read sheets, list MCP resources, or read raw rows unless the user explicitly asks for raw audit metadata. If the user asks to read values from a source-list sheet such as `Dropdown Lists`, read the actual cell values with `read_values`/targeted range read; do not treat the sheet name itself as validation intent. To add a missing option, update the returned source-list cell/range with `mode=auto` when bounded; for inline comma-list validation, use one `preview_update intent.action=write_data_validation` with existing options plus the new option, then one `apply_update`.

For direct range writes, send the cell values in the structured `values` field, not only in the request text:

```text
excel.agent.run mode=preview_update intent.action=write_values target.sheetName=Booking target.range=A3:X7 values.values=[[46198,46198,"2X20'GP","2X20'GP","SC89","Loading at Rayong Factory",20,"2X20'GP","RAYONG DEPOT KM.5","STAFF2","038-123456","TER 3","STAFF2","038-654321","EVERGREEN V.123","27/6/26","EVERGREEN","SINGAPORE (SGSIN)","036GX11111","21/6/26","22/6/26","24/6/26","Before 12:00","Standard handling"]]
excel.agent.run mode=apply_update operationId="..." confirmationToken="..."
```

Pass `operationId` and `confirmationToken` as top-level tool arguments for `apply_update`; do not put the token in `values`.

For related edits across multiple ranges, send one grouped preview using `values.patches` and then apply that returned operation once. Do not issue one preview/apply pair per zone, column group, or row block unless the grouped apply returns a hard failure with actionable issue details.

For booking images or OCR-extracted client screenshots that need fields rotated into headers/values and styled like an existing sheet, send one `preview_update` with `intent.action: "replace_range_with_styled_table"`, structured `values`, and any style source ranges. Apply the returned operation once. Do not split the task into separate clear, write, autofit, and style calls.

Omitted mode or `mode=auto` remains compatible for casual prompts, but explicit modes are more predictable for agent UIs. The backend should either answer from compact proof in one call or return a precise `nextAction`; agents should not chain primitive compact tools.

`mode=status` reports workbook readiness with `connectionState`. `ready` means the add-in responded and an active workbook is available. `stale` means the backend saw an old or unresponsive taskpane session; reload or reopen the OpenWorkbook Local taskpane in Excel before retrying, and restart Excel only if the taskpane cannot reconnect.
