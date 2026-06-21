import type { AgentRunOutput } from "@components-kit/open-workbook-protocol";

export function jsonResult(value: unknown) {
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(value, null, 2)
      }
    ]
  };
}

export function agentJsonResult(value: AgentRunOutput) {
  const jsonSafeValue = JSON.parse(JSON.stringify(value)) as AgentRunOutput;
  const resourceLinks = Array.isArray(jsonSafeValue.resourceLinks) ? jsonSafeValue.resourceLinks : [];
  return {
    content: [
      {
        type: "text" as const,
        text: compactAgentResultText(jsonSafeValue)
      }
    ],
    structuredContent: jsonSafeValue,
    resources: resourceLinks
      .filter((resource) => typeof resource?.uri === "string")
      .map((resource) => ({
        uri: resource.uri,
        name: resource.name ?? resource.uri,
        description: resource.description,
        mimeType: resource.mimeType ?? "application/json"
      }))
  };
}

function compactAgentResultText(value: AgentRunOutput): string {
  const lines = [
    `${value.status} ${value.mode}: ${value.summary}`,
    `nextAction: ${value.nextAction}`
  ];
  if (value.workbookContextId) {
    lines.push(`workbookContextId: ${value.workbookContextId}`);
  }
  if (value.operationId) {
    lines.push(`operationId: ${value.operationId}`);
  }
  if (value.confirmationToken) {
    lines.push("confirmationToken: present in structuredContent");
  }
  if (value.continuation?.resultUri) {
    lines.push(`resultUri: ${value.continuation.resultUri}`);
  }
  if (value.continuation?.fullResultUri) {
    lines.push(`fullResultUri: ${value.continuation.fullResultUri}`);
    lines.push("full detail: call excel.agent.run with fullResultUri in request/continuation; do not use webfetch for excel:// handles");
  }
  if (value.resourceLinks.length > 0) {
    lines.push(`resources: ${value.resourceLinks.map((resource) => resource.uri).join(", ")}`);
  }
  if (value.invalidatedContextIds?.length) {
    lines.push(`invalidatedContextIds: ${value.invalidatedContextIds.join(", ")}`);
  }
  if (value.invalidatedResourceUris?.length) {
    lines.push(`invalidatedResourceUris: ${value.invalidatedResourceUris.join(", ")}`);
  }
  if (value.warnings.length > 0) {
    lines.push(`warnings: ${value.warnings.slice(0, 3).join(" | ")}`);
  }
  const payloadBytes = value.telemetry?.payloadBytes;
  const estimatedTokens = value.telemetry?.estimatedTokens;
  if (typeof payloadBytes === "number" || typeof estimatedTokens === "number") {
    lines.push(`telemetry: ${payloadBytes ?? "?"} bytes, ~${estimatedTokens ?? "?"} tokens`);
  }
  return lines.join("\n");
}
