import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, "..");
const OUTPUT_FILE = join(ROOT, "llms-full.txt");

const HEADER = `# Open Workbook

> Local-first MCP runtime for fast, reversible, template-aware Excel automation through Office.js, a local backend, and any MCP-capable agent UI.

This file is generated for LLM and agent context. It combines the project README, core documentation, and workspace package READMEs into one reference.

## Quick Start

\`\`\`bash
npx -y @component-kit/open-workbook setup
\`\`\`

Then install the skill with \`npx skills add components-kit/open-workbook --skill open-workbook-excel\` and paste the printed MCP config into an agent UI.

## Production Rules

- Use stable MCP tools by default.
- Use \`excel.runtime.get_capabilities\` to inspect tool catalog and host capabilities.
- Use \`excel.plan.*\` or \`excel.batch.*\` for rollback-aware mutations.
- Do not bypass permissions, snapshots, backups, fingerprint checks, or transaction logging.
- Treat Office.js host limitations as capability-status metadata, not silent success.
- For multi-agent work, use tasks, locks, serialized transactions, conflict guidance, and rollback previews.
`;

const SOURCE_FILES = [
  "README.md",
  "skills/open-workbook-excel/SKILL.md",
  "skills/open-workbook-excel/references/tool-selection.md",
  "skills/open-workbook-excel/references/workflows.md",
  "skills/open-workbook-excel/references/reliability.md",
  "skills/open-workbook-excel/references/performance.md",
  "skills/open-workbook-excel/references/multi-agent.md",
  "docs/installation.md",
  "docs/mcp-clients.md",
  "docs/instructions.md",
  "docs/advanced-runtime.md",
  "docs/architecture.md",
  "docs/tool-surface.md",
  "docs/backup-lifecycle.md",
  "docs/template-system.md",
  "docs/style-fidelity.md",
  "docs/formula-intelligence.md",
  "docs/table-filter-sort.md",
  "docs/names-regions.md",
  "docs/pivot-chart.md",
  "docs/advanced-range-reads.md",
  "docs/validation-repair.md",
  "docs/permissions-cleaning.md",
  "docs/workbook-file-lifecycle.md",
  "docs/performance.md",
  "docs/multi-agent-runtime.md",
  "docs/production-readiness.md",
  "docs/service-wrapper.md",
  "docs/opencode.md",
  "docs/packaging.md",
  "docs/sideloading.md",
  "docs/roadmap.md",
  "packages/protocol/README.md",
  "packages/excel-core/README.md",
  "packages/office-js-engine/README.md",
  "packages/cli/README.md",
  "apps/backend/README.md",
  "apps/mcp-server/README.md",
  "apps/excel-addin/README.md"
];

const sections = SOURCE_FILES.map((relativePath) => {
  const absolutePath = join(ROOT, relativePath);
  let content = readFileSync(absolutePath, "utf-8").trim();
  content = content.replace(/^# /, "## ");
  return `Source: ${relativePath}\n\n${content}`;
});

const fullText = [HEADER, ...sections].join("\n\n---\n\n");

writeFileSync(OUTPUT_FILE, `${fullText}\n`);
console.warn(`Generated ${OUTPUT_FILE} (${SOURCE_FILES.length} sources, ${fullText.split("\n").length} lines)`);
