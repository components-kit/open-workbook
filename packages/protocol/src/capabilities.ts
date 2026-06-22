export interface EngineCapability {
  name: string;
  supported: boolean;
  platforms: Array<"mac" | "windows" | "web">;
  notes?: string;
  requires?: Array<{
    set: string;
    version: string;
  }>;
}

export interface OfficeApiSetSupport {
  set: "ExcelApi" | "ExcelApiOnline" | string;
  version: string;
  supported: boolean;
}

export interface HostCapabilityStatus {
  name: string;
  supported: boolean;
  status: "supported" | "limited" | "unsupported" | "unknown";
  reason?: string;
  requires?: Array<{
    set: string;
    version: string;
  }>;
}

export interface RuntimeCapabilities {
  engine: {
    name: string;
    version: string;
    platform: "mac" | "windows" | "web" | "unknown";
    host?: string;
    officeVersion?: string;
    taskpaneBundleVersion?: string;
  };
  apiSets?: OfficeApiSetSupport[];
  capabilities: EngineCapability[];
  hostCapabilities?: HostCapabilityStatus[];
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
