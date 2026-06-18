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

The skill tells agents to use the public `excel.agent.run` workflow, preserve live Excel safety checks, pass structured intent fields when available, normalize multilingual requests into canonical routing fields, validate after changes, and report backup or transaction IDs. Primitive Excel capabilities are backend-owned and are not normal agent calls.

## Fallback Instructions

For clients that do not support skills.sh, print or write a portable instruction bundle:

```bash
npx -y @components-kit/open-workbook instructions
npx -y @components-kit/open-workbook instructions --out open-workbook-excel.md
```
