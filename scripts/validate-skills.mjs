#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const SKILLS_DIR = join(ROOT, "skills");
const REQUIRED_PRIMARY_TERMS = [
  "excel.runtime.get_status",
  "excel.runtime.get_capabilities",
  "excel.workbook.get_workbook_map",
  "excel.plan.*",
  "excel.batch.*",
  "snapshots",
  "backups",
  "fingerprints",
  "transaction",
  "CAPABILITY_UNAVAILABLE"
];

const errors = [];

function readText(path) {
  return readFileSync(path, "utf8");
}

function parseFrontmatter(path, contents) {
  const match = contents.match(/^---\n([\s\S]*?)\n---\n?/);
  if (!match) {
    errors.push(`${relative(ROOT, path)} is missing YAML frontmatter`);
    return {};
  }

  const metadata = {};
  for (const line of match[1].split("\n")) {
    const field = line.match(/^([A-Za-z0-9_-]+):\s*(.*)$/);
    if (!field) {
      errors.push(`${relative(ROOT, path)} has invalid frontmatter line: ${line}`);
      continue;
    }
    metadata[field[1]] = field[2].trim().replace(/^"|"$/g, "");
  }
  return metadata;
}

function listSkillFiles(dir) {
  if (!existsSync(dir)) {
    return [];
  }

  return readdirSync(dir)
    .map((entry) => join(dir, entry))
    .filter((entryPath) => statSync(entryPath).isDirectory())
    .map((skillDir) => join(skillDir, "SKILL.md"))
    .filter((skillFile) => existsSync(skillFile));
}

const skillFiles = listSkillFiles(SKILLS_DIR);
if (skillFiles.length === 0) {
  errors.push("No skills found under skills/*/SKILL.md");
}

for (const skillFile of skillFiles) {
  const contents = readText(skillFile);
  const metadata = parseFrontmatter(skillFile, contents);
  const skillDir = dirname(skillFile);
  const folderName = relative(SKILLS_DIR, skillDir);

  if (!metadata.name) {
    errors.push(`${relative(ROOT, skillFile)} frontmatter must include name`);
  } else if (!/^[a-z0-9][a-z0-9-]{0,62}[a-z0-9]$/.test(metadata.name)) {
    errors.push(`${relative(ROOT, skillFile)} name must be a slug-safe skill name`);
  } else if (metadata.name !== folderName) {
    errors.push(`${relative(ROOT, skillFile)} name must match folder name ${folderName}`);
  }

  if (!metadata.description || metadata.description.length < 80) {
    errors.push(`${relative(ROOT, skillFile)} description should be specific enough to trigger reliably`);
  }

  const markdownLinks = [...contents.matchAll(/\[[^\]]+\]\(([^)]+)\)/g)].map((match) => match[1]);
  const inlineReferences = [...contents.matchAll(/`(references\/[^`]+)`/g)].map((match) => match[1]);
  for (const link of [...markdownLinks, ...inlineReferences]) {
    if (!link.startsWith("references/")) {
      continue;
    }
    const target = join(skillDir, link);
    if (!existsSync(target)) {
      errors.push(`${relative(ROOT, skillFile)} references missing file ${link}`);
    }
  }

  if (metadata.name === "open-workbook-excel") {
    for (const term of REQUIRED_PRIMARY_TERMS) {
      if (!contents.includes(term)) {
        errors.push(`${relative(ROOT, skillFile)} must mention ${term}`);
      }
    }
  }
}

if (errors.length > 0) {
  console.error("Skill validation failed:");
  console.error(errors.map((error) => `- ${error}`).join("\n"));
  process.exit(1);
}

console.log(`Skill validation passed: ${skillFiles.length} skill(s) checked.`);
