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
    ...jsonResult(jsonSafeValue),
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
