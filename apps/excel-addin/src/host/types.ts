export type HostDependency = "office-js" | "office-document-file" | "backend-only" | "unsupported";
export type HostRpcHandler = (params: unknown) => Promise<unknown> | unknown;

export interface HostMethodDefinition {
  method: string;
  implementationOwner: string;
  handler: HostRpcHandler;
  relatedBackendCapabilities: string[];
  operationKinds: string[];
  hostDependency: HostDependency;
  unitTestFile: string;
}
