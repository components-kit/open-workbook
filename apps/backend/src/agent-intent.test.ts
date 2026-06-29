import { describe, expect, it } from "vitest";
import { modeForIntentAction, normalizeAgentIntent } from "./agent-intent.js";

describe("agent intent normalization", () => {
  it("accepts common agent aliases for canonical actions", () => {
    const style = normalizeAgentIntent({
      request: "Apply styling from Employees to Booking",
      intent: { action: "apply_style_from_template" }
    });
    const filter = normalizeAgentIntent({
      request: "Add autofilter",
      intent: { action: "add_filter" }
    });

    expect(style.accepted).toBe(true);
    expect(style.action).toBe("copy_style_from_template");
    expect(filter.accepted).toBe(true);
    expect(filter.action).toBe("filter_range");
  });

  it("normalizes clear filter aliases to clear_table_filters", () => {
    const intent = normalizeAgentIntent({
      request: "Clear all filters so I can see full data",
      intent: { action: "clear_filters" }
    });

    expect(intent.accepted).toBe(true);
    expect(intent.action).toBe("clear_table_filters");
  });

  it("accepts improve_visual_readability as a high-level structured action", () => {
    const intent = normalizeAgentIntent({
      request: "Make this sheet easier to read",
      intent: { action: "improve_visual_readability" }
    });

    expect(intent.accepted).toBe(true);
    expect(intent.action).toBe("improve_visual_readability");
    expect(intent.rejectedReason).toBeUndefined();
  });

  it("routes style/design overview and permission actions through answer mode", () => {
    expect(modeForIntentAction("style_overview")).toBe("answer");
    expect(modeForIntentAction("workbook_design_overview")).toBe("answer");
    expect(modeForIntentAction("get_permissions")).toBe("answer");
    expect(modeForIntentAction("set_permissions")).toBe("answer");
    expect(modeForIntentAction("allow_destructive_actions")).toBe("answer");
  });
});
