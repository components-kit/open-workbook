import { makeId, type AgentOperationId, type AgentRunOutput, type ExcelOperation, type WorkbookId } from "@components-kit/open-workbook-protocol";

export interface PendingAgentOperation {
  operationId: AgentOperationId | string;
  confirmationToken: string;
  workbookContextId: string;
  workbookId: WorkbookId;
  operations: ExcelOperation[];
  changes: NonNullable<AgentRunOutput["changes"]>;
  createdAt: number;
  summary: string;
  sourceFingerprintHash?: string;
}

export class AgentOperationStore {
  private readonly pending = new Map<string, PendingAgentOperation>();

  create(input: Omit<PendingAgentOperation, "operationId" | "confirmationToken" | "createdAt">): PendingAgentOperation {
    const operation: PendingAgentOperation = {
      ...input,
      operationId: makeId<AgentOperationId>("agentop"),
      confirmationToken: makeId("confirm"),
      createdAt: Date.now()
    };
    this.pending.set(operation.operationId, operation);
    return operation;
  }

  get(operationId: string): PendingAgentOperation | undefined {
    return this.pending.get(operationId);
  }

  delete(operationId: string): void {
    this.pending.delete(operationId);
  }
}
