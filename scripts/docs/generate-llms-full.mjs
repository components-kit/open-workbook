import { readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { findRepoRoot } from "../lib/repo-root.mjs";

const ROOT = findRepoRoot(import.meta.url);
const OUTPUT_FILE = join(ROOT, "llms-full.txt");

const HEADER = `# Open Workbook

> Local Excel agent runtime that gives any MCP-capable agent one safe tool, \`excel.agent.run\`, for live desktop Excel workbooks.

This file is generated for LLM and agent context. It combines the project README, core documentation, and workspace package READMEs into one reference.

## Quick Start

\`\`\`bash
npx -y @components-kit/open-workbook setup
\`\`\`

Then install the skill with \`npx skills add components-kit/open-workbook --skill open-workbook-excel\` and paste the printed MCP launch command into an agent UI's local stdio MCP configuration.

## Production Rules

- Use \`excel.agent.run\` as the public MCP tool.
- Keep primitive Excel capabilities backend-owned unless explicitly working on backend tests.
- Normalize multilingual requests into canonical \`intent.action\`, \`intent.targetHints\`, \`target\`, and \`values\` fields while preserving the original user request.
- Do not bypass permissions, snapshots, backups, fingerprint checks, or transaction logging.
- Treat Office.js host limitations as capability-status metadata, not silent success.
- For multi-agent work, use tasks, locks, serialized transactions, conflict guidance, and rollback previews.
`;

const SOURCE_FILES = [
  "README.md",
  "CHANGELOG.md",
  "skills/open-workbook-excel/SKILL.md",
  "skills/open-workbook-excel/references/agent-run.md",
  "skills/open-workbook-excel/references/capability-map.md",
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
  "docs/release.md",
  "docs/sideloading.md",
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
