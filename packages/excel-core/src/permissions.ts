import type { DestructiveLevel } from "@component-kit/open-workbook-protocol";

export interface PermissionPolicy {
  allowWrites: boolean;
  allowDestructiveActions: boolean;
  allowWorkbookActions: boolean;
  requireConfirmationFor: DestructiveLevel[];
}

export const DefaultPermissionPolicy: PermissionPolicy = {
  allowWrites: true,
  allowDestructiveActions: false,
  allowWorkbookActions: false,
  requireConfirmationFor: ["values", "format", "structure", "workbook"]
};

export function requiresConfirmation(policy: PermissionPolicy, level: DestructiveLevel): boolean {
  return policy.requireConfirmationFor.includes(level);
}

export function assertMutationAllowed(policy: PermissionPolicy, level: DestructiveLevel): void {
  if (!policy.allowWrites && level !== "none") {
    throw new Error("Writes are disabled by permission policy");
  }
  if (!policy.allowDestructiveActions && (level === "structure" || level === "workbook")) {
    throw new Error("Destructive actions are disabled by permission policy");
  }
  if (!policy.allowWorkbookActions && level === "workbook") {
    throw new Error("Workbook-level actions are disabled by permission policy");
  }
}
