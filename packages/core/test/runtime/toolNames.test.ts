import { describe, expect, it } from "vitest";
import { repairToolName } from "../../src/runtime/toolNames.ts";

const known = ["execute_office_js", "get_slide", "add_slide"];

describe("repairToolName", () => {
  it("maps a prefixed name to the registered tool it ends with", () => {
    expect(repairToolName("someone_execute_office_js", known)).toBe("execute_office_js");
    expect(repairToolName("functions.get_slide", known)).toBe("get_slide");
  });

  it("leaves known, unknown and ambiguous names alone", () => {
    expect(repairToolName("add_slide", known)).toBe("add_slide");
    expect(repairToolName("make_coffee", known)).toBe("make_coffee");
    expect(repairToolName("x_slide", ["get_slide", "x_slide_y"])).toBe("x_slide");
  });
});
