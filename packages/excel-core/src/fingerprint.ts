import type { A1Range, RangeFingerprint, WorkbookFingerprint, WorkbookId } from "@components-kit/open-workbook-protocol";
import { cellCount } from "./range-address.js";

export function createRangeFingerprint(range: A1Range, payload: unknown, capturedAt = new Date().toISOString()): RangeFingerprint {
  return {
    range,
    hash: hashStable(payload),
    cellCount: cellCount(range.address),
    capturedAt
  };
}

export function createWorkbookFingerprint(
  workbookId: WorkbookId,
  workbookPayload: unknown,
  structurePayload: unknown,
  capturedAt = new Date().toISOString()
): WorkbookFingerprint {
  return {
    workbookId,
    workbookHash: hashStable(workbookPayload),
    structureHash: hashStable(structurePayload),
    capturedAt
  };
}

export function hashStable(value: unknown): string {
  const input = stableStringify(value);
  let hash = 0xcbf29ce484222325n;
  const prime = 0x100000001b3n;

  for (let i = 0; i < input.length; i += 1) {
    hash ^= BigInt(input.charCodeAt(i));
    hash = BigInt.asUintN(64, hash * prime);
  }

  return hash.toString(16).padStart(16, "0");
}

export function stableStringify(value: unknown): string {
  if (value === null || typeof value !== "object") {
    return JSON.stringify(value);
  }

  if (Array.isArray(value)) {
    return `[${value.map((item) => stableStringify(item)).join(",")}]`;
  }

  const record = value as Record<string, unknown>;
  return `{${Object.keys(record)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${stableStringify(record[key])}`)
    .join(",")}}`;
}
