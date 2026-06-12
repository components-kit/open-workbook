import {
  type A1Range,
  type DiffSummary,
  type SnapshotId,
  type SnapshotRef,
  type WorkbookId,
  type WorkbookSnapshotResponse,
  makeId
} from "@open-workbook/protocol";

export interface WorkbookSnapshotRecord extends SnapshotRef {
  payload: WorkbookSnapshotResponse;
  invalidatedAt?: string;
}

export class SnapshotManager {
  private readonly snapshots = new Map<SnapshotId, WorkbookSnapshotRecord>();

  createSnapshot(input: {
    workbookId: WorkbookId;
    reason: string;
    affectedRanges: A1Range[];
    payload: WorkbookSnapshotResponse;
  }): WorkbookSnapshotRecord {
    const snapshotId = makeId<SnapshotId>("snapshot");
    const snapshot: WorkbookSnapshotRecord = {
      snapshotId,
      workbookId: input.workbookId,
      createdAt: new Date().toISOString(),
      reason: input.reason,
      affectedRanges: input.affectedRanges,
      payload: input.payload
    };
    this.snapshots.set(snapshotId, snapshot);
    return snapshot;
  }

  getSnapshot(snapshotId: SnapshotId): WorkbookSnapshotRecord | undefined {
    return this.snapshots.get(snapshotId);
  }

  listSnapshots(workbookId: WorkbookId): WorkbookSnapshotRecord[] {
    return [...this.snapshots.values()]
      .filter((snapshot) => snapshot.workbookId === workbookId)
      .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }

  invalidate(snapshotId: SnapshotId): WorkbookSnapshotRecord | undefined {
    const snapshot = this.snapshots.get(snapshotId);
    if (!snapshot) {
      return undefined;
    }
    snapshot.invalidatedAt = new Date().toISOString();
    return snapshot;
  }

  deleteSnapshot(snapshotId: SnapshotId): boolean {
    return this.snapshots.delete(snapshotId);
  }

  compare(leftId: SnapshotId, rightId: SnapshotId): DiffSummary | undefined {
    const left = this.snapshots.get(leftId);
    const right = this.snapshots.get(rightId);
    if (!left || !right) {
      return undefined;
    }

    const changedRanges: A1Range[] = [];
    let cellsChanged = 0;

    for (const rightRange of right.payload.rangeSnapshots) {
      const leftRange = left.payload.rangeSnapshots.find(
        (candidate) =>
          candidate.fingerprint.range.sheetName === rightRange.fingerprint.range.sheetName &&
          candidate.fingerprint.range.address === rightRange.fingerprint.range.address
      );
      if (!leftRange || leftRange.fingerprint.hash !== rightRange.fingerprint.hash) {
        changedRanges.push(rightRange.fingerprint.range);
        cellsChanged += rightRange.fingerprint.cellCount;
      }
    }

    return {
      title: `Snapshot diff ${leftId} -> ${rightId}`,
      changedRanges,
      cellsChanged,
      formulasChanged: 0,
      stylesChanged: 0,
      tablesChanged: 0,
      sheetsChanged: left.payload.workbookFingerprint.structureHash === right.payload.workbookFingerprint.structureHash ? 0 : 1,
      destructiveLevel: changedRanges.length > 0 ? "values" : "none"
    };
  }
}
