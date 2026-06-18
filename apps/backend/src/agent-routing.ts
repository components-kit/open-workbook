import type { AgentRunMode } from "@components-kit/open-workbook-protocol";
import { modeForIntentAction, type NormalizedAgentIntent } from "./agent-intent.js";

export interface IntentRoute {
  mode: AgentRunMode;
  matchedRule: string;
  confidence: number;
  reasons: string[];
}

interface IntentRouteRule {
  mode: AgentRunMode;
  matchedRule: string;
  confidence: number;
  pattern: RegExp;
  reason: string;
}

const ROUTE_RULES: IntentRouteRule[] = [
  {
    mode: "preview_update",
    matchedRule: "mutation.keyword",
    confidence: 0.82,
    pattern: /\b(add|update|set|change|replace|write|fill|append|edit|fix|repair|create|insert|delete|remove|clear|rename|format|style|duplicate|copy|template)\b/i,
    reason: "Request contains workbook mutation keywords."
  }
];

export function routeAgentRequest(request: string, requestedMode: AgentRunMode = "auto", intent?: NormalizedAgentIntent): IntentRoute {
  if (requestedMode !== "auto") {
    return {
      mode: requestedMode,
      matchedRule: "mode.explicit",
      confidence: 1,
      reasons: [`Caller explicitly requested ${requestedMode}.`]
    };
  }

  if (intent?.accepted && intent.action) {
    return {
      mode: modeForIntentAction(intent.action),
      matchedRule: "caller_intent.action",
      confidence: intent.confidence ?? 0.9,
      reasons: [intent.reason ?? `Caller provided structured intent action ${intent.action}.`]
    };
  }

  const matched = ROUTE_RULES.find((rule) => rule.pattern.test(request));
  if (matched) {
    return {
      mode: matched.mode,
      matchedRule: matched.matchedRule,
      confidence: matched.confidence,
      reasons: [matched.reason]
    };
  }

  return {
    mode: "answer",
    matchedRule: "default.answer",
    confidence: 0.65,
    reasons: ["No mutation route matched; defaulting to answer mode."]
  };
}
