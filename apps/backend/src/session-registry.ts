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

  createSession(): AddinSession {
    const connectionId = makeId<ConnectionId>("conn");
    const now = new Date().toISOString();
    const session: AddinSession = {
      connectionId,
      connectedAt: now,
      lastSeenAt: now
    };
    this.sessions.set(connectionId, session);
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
  }

  list(): AddinSession[] {
    return [...this.sessions.values()];
  }

  getActive(): AddinSession | undefined {
    return this.list().sort((a, b) => b.lastSeenAt.localeCompare(a.lastSeenAt))[0];
  }
}
