import type { ConnectionId, RuntimeCapabilities, WorkbookRef } from "@open-workbook/protocol";
import { makeId } from "@open-workbook/protocol";

export interface AddinSession {
  connectionId: ConnectionId;
  connectedAt: string;
  lastSeenAt: string;
  capabilities?: RuntimeCapabilities;
  activeWorkbook?: WorkbookRef;
}

export class SessionRegistry {
  private readonly sessions = new Map<ConnectionId, AddinSession>();
  private activeConnectionId: ConnectionId | undefined;

  createSession(): AddinSession {
    const connectionId = makeId<ConnectionId>("conn");
    const now = new Date().toISOString();
    const session: AddinSession = {
      connectionId,
      connectedAt: now,
      lastSeenAt: now
    };
    this.sessions.set(connectionId, session);
    this.activeConnectionId ??= connectionId;
    return session;
  }

  touch(connectionId: ConnectionId): void {
    const session = this.sessions.get(connectionId);
    if (session) {
      session.lastSeenAt = new Date().toISOString();
    }
  }

  update(connectionId: ConnectionId, patch: Partial<Omit<AddinSession, "connectionId" | "connectedAt">>): void {
    const session = this.sessions.get(connectionId);
    if (!session) {
      return;
    }
    Object.assign(session, patch, { lastSeenAt: new Date().toISOString() });
  }

  remove(connectionId: ConnectionId): void {
    this.sessions.delete(connectionId);
    if (this.activeConnectionId === connectionId) {
      this.activeConnectionId = this.list().sort((a, b) => b.lastSeenAt.localeCompare(a.lastSeenAt))[0]?.connectionId;
    }
  }

  list(): AddinSession[] {
    return [...this.sessions.values()];
  }

  getActive(): AddinSession | undefined {
    if (this.activeConnectionId) {
      return this.sessions.get(this.activeConnectionId);
    }
    return this.list().sort((a, b) => b.lastSeenAt.localeCompare(a.lastSeenAt))[0];
  }

  setActiveWorkbook(workbookIdOrName: string): AddinSession | undefined {
    const session = this.list().find(
      (candidate) =>
        candidate.activeWorkbook?.workbookId === workbookIdOrName || candidate.activeWorkbook?.name === workbookIdOrName
    );
    if (session) {
      this.activeConnectionId = session.connectionId;
    }
    return session;
  }
}
