export interface EngineCapability {
  name: string;
  supported: boolean;
  platforms: Array<"mac" | "windows" | "web">;
  notes?: string;
}

export interface RuntimeCapabilities {
  engine: {
    name: string;
    version: string;
    platform: "mac" | "windows" | "web" | "unknown";
  };
  capabilities: EngineCapability[];
}

export const RequiredV1Capabilities = [
  "workbook.context",
  "worksheet.list",
  "range.read.values",
  "range.read.formulas",
  "range.read.formats",
  "range.write.values",
  "range.write.formulas",
  "range.clear.values.keep_format",
  "worksheet.copy",
  "table.read",
  "filter.read",
  "backup.region_snapshot",
  "template.fingerprint"
] as const;

export type RequiredV1Capability = (typeof RequiredV1Capabilities)[number];
