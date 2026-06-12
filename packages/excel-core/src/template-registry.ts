import type { TemplateId, TemplateRef, WorkbookId } from "@open-workbook/protocol";
import { makeId } from "@open-workbook/protocol";
import { hashStable } from "./fingerprint.js";

export interface TemplateFingerprint {
  structureHash: string;
  formulaHash: string;
  styleHash: string;
  filterHash: string;
  tableHash: string;
  printLayoutHash: string;
}

export interface TemplateRecord extends TemplateRef {
  workbookId?: WorkbookId;
  sourceSheetName: string;
  fingerprint: TemplateFingerprint;
  dataRegions: string[];
  createdAt: string;
  updatedAt: string;
}

export interface RegisterTemplateInput {
  name: string;
  scope: "workbook" | "local";
  sourceSheetName: string;
  workbookId?: WorkbookId;
  dataRegions?: string[];
  fingerprintPayload: {
    structure: unknown;
    formulas: unknown;
    styles: unknown;
    filters: unknown;
    tables: unknown;
    printLayout: unknown;
  };
}

export class TemplateRegistry {
  private readonly records = new Map<TemplateId, TemplateRecord>();

  register(input: RegisterTemplateInput): TemplateRecord {
    if (input.scope === "workbook" && !input.workbookId) {
      throw new Error("Workbook templates require workbookId");
    }

    const templateId = makeId<TemplateId>("template");
    const now = new Date().toISOString();
    const record: TemplateRecord = {
      templateId,
      name: input.name,
      scope: input.scope,
      version: 1,
      sourceSheetName: input.sourceSheetName,
      dataRegions: input.dataRegions ?? [],
      fingerprint: {
        structureHash: hashStable(input.fingerprintPayload.structure),
        formulaHash: hashStable(input.fingerprintPayload.formulas),
        styleHash: hashStable(input.fingerprintPayload.styles),
        filterHash: hashStable(input.fingerprintPayload.filters),
        tableHash: hashStable(input.fingerprintPayload.tables),
        printLayoutHash: hashStable(input.fingerprintPayload.printLayout)
      },
      createdAt: now,
      updatedAt: now
    };
    if (input.workbookId !== undefined) {
      record.workbookId = input.workbookId;
    }
    this.records.set(templateId, record);
    return record;
  }

  get(templateId: TemplateId): TemplateRecord | undefined {
    return this.records.get(templateId);
  }

  list(options: { workbookId?: WorkbookId } = {}): TemplateRecord[] {
    return [...this.records.values()].filter((record) => {
      if (record.scope === "local") {
        return true;
      }
      return !options.workbookId || record.workbookId === options.workbookId;
    });
  }
}
