import { OPEN_WORKBOOK_VERSION, type LockLeasePolicy, type RuntimeCapabilities } from "@components-kit/open-workbook-protocol";

export function runtimeVersion(): string {
  return process.env.OPEN_WORKBOOK_VERSION ?? OPEN_WORKBOOK_VERSION;
}

export function addinStaleTtlMs(): number {
  return positiveIntegerEnv("OPEN_WORKBOOK_ADDIN_STALE_TTL_MS", 15_000);
}

export function addinHealthTimeoutMs(): number {
  return positiveIntegerEnv("OPEN_WORKBOOK_ADDIN_HEALTH_TIMEOUT_MS", 2_500);
}

export function defaultLockLeasePolicy(): LockLeasePolicy {
  const maxTtlMs = positiveIntegerEnv("OPEN_WORKBOOK_LOCK_MAX_TTL_MS", 600_000);
  return {
    maxTtlMs,
    defaultTtlMs: Math.min(positiveIntegerEnv("OPEN_WORKBOOK_LOCK_DEFAULT_TTL_MS", 120_000), maxTtlMs),
    transactionTtlMs: Math.min(positiveIntegerEnv("OPEN_WORKBOOK_LOCK_TRANSACTION_TTL_MS", 120_000), maxTtlMs),
    allowManualLocks: process.env.OPEN_WORKBOOK_ALLOW_MANUAL_LOCKS !== "0"
  };
}

export function disconnectedRuntimeCapabilities(): RuntimeCapabilities {
  return {
    engine: {
      name: "disconnected",
      version: runtimeVersion(),
      platform: "unknown",
      host: "Excel"
    },
    apiSets: [],
    capabilities: [],
    hostCapabilities: [
      {
        name: "excel-addin",
        supported: false,
        status: "unknown",
        reason: "No Excel add-in session is connected."
      }
    ]
  };
}

export function positiveIntegerEnv(name: string, fallback: number): number {
  const value = Number(process.env[name] ?? fallback);
  return Number.isFinite(value) && value > 0 ? Math.round(value) : fallback;
}
