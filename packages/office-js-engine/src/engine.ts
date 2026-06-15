import type {
  BatchRequest,
  CompiledBatch,
  OperationResult,
  RuntimeCapabilities,
  WorkbookRef
} from "@components-kit/open-workbook-protocol";

export interface ExcelEngine {
  readonly name: string;
  readonly version: string;
  getCapabilities(): Promise<RuntimeCapabilities>;
  getActiveWorkbook(): Promise<WorkbookRef | undefined>;
  executeBatch(request: BatchRequest, compiled: CompiledBatch): Promise<OperationResult>;
}

export interface OfficeJsEngineOptions {
  chunkCellLimit: number;
  suspendCalculationCellThreshold: number;
  suspendScreenUpdatingCellThreshold: number;
}

export const DefaultOfficeJsEngineOptions: OfficeJsEngineOptions = {
  chunkCellLimit: 50_000,
  suspendCalculationCellThreshold: 10_000,
  suspendScreenUpdatingCellThreshold: 10_000
};

export class OfficeJsEngine implements ExcelEngine {
  readonly name = "office-js";
  readonly version = "0.1.6";

  constructor(private readonly options: OfficeJsEngineOptions = DefaultOfficeJsEngineOptions) {}

  async getCapabilities(): Promise<RuntimeCapabilities> {
    return {
      engine: {
        name: this.name,
        version: this.version,
        platform: "unknown"
      },
      capabilities: [
        { name: "workbook.context", supported: true, platforms: ["mac", "windows"] },
        { name: "worksheet.copy", supported: true, platforms: ["mac", "windows"] },
        { name: "range.write.values", supported: true, platforms: ["mac", "windows"] },
        { name: "range.write.formulas", supported: true, platforms: ["mac", "windows"] },
        { name: "template.fingerprint", supported: true, platforms: ["mac", "windows"] },
        {
          name: "workbook.restore.full_file",
          supported: false,
          platforms: ["mac", "windows"],
          notes: "Full file restore is coordinated by the backend and may require user or OS-level file replacement."
        }
      ]
    };
  }

  async getActiveWorkbook(): Promise<WorkbookRef | undefined> {
    return undefined;
  }

  async executeBatch(request: BatchRequest, compiled: CompiledBatch): Promise<OperationResult> {
    const started = Date.now();

    return {
      ok: false,
      rollbackAvailable: compiled.requiredBackups.length > 0,
      backups: [],
      warnings: [
        {
          code: "ENGINE_NOT_BOUND",
          message: "OfficeJsEngine is scaffolded but not yet bound to Excel.run in the add-in runtime."
        }
      ],
      telemetry: {
        durationMs: Date.now() - started,
        syncCount: 0,
        payloadBytes: JSON.stringify(request).length,
        cellsWritten: compiled.estimatedCellsTouched,
        rangeCount: compiled.targetFingerprints.length,
        chunkCount: Math.max(1, Math.ceil(compiled.estimatedCellsTouched / this.options.chunkCellLimit)),
        engineName: this.name,
        engineVersion: this.version,
        warningCount: 1
      }
    };
  }
}
