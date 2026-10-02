import { afterEach, expect, test, vi } from "vitest";
import { formatAge } from "../../src/ui/HistoryView.tsx";

afterEach(() => {
  vi.useRealTimers();
});

test("formatAge compares calendar dates, not elapsed hours", () => {
  vi.useFakeTimers();
  vi.setSystemTime(new Date(2026, 9, 2, 9, 0));
  expect(formatAge(new Date(2026, 9, 1, 22, 0).getTime())).toBe("yesterday");
  expect(formatAge(new Date(2026, 9, 2, 0, 30).getTime())).not.toMatch(/day/);
  expect(formatAge(new Date(2026, 8, 30, 23, 0).getTime())).toBe("2 days ago");
});
