# Agent Instructions

Open Workbook ships an agent skill in the repository under `skills/open-workbook-excel/`.

Install it with skills.sh:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel
```

For a global OpenCode install:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel -a opencode -g -y
```

The skills.sh-compatible layout is:

```text
skills/
  open-workbook-excel/
    SKILL.md
    references/
      tool-selection.md
      workflows.md
      reliability.md
      performance.md
      multi-agent.md
```

The skill tells agents to inspect runtime status first, choose narrow Excel MCP tools, batch workbook writes, preserve templates and formulas, validate after changes, and report backup or transaction IDs.

## Fallback Instructions

For clients that do not support skills.sh, print or write a portable instruction bundle:

```bash
npx -y @component-kit/open-workbook instructions
npx -y @component-kit/open-workbook instructions --out open-workbook-excel.md
```
