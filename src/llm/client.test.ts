import { describe, expect, it, vi } from "vitest";
import { z } from "zod";

const create = vi.fn();
vi.mock("openai", async (orig) => {
  const real = await orig<typeof import("openai")>();
  class FakeOpenAI {
    chat = { completions: { create } };
  }
  Object.assign(FakeOpenAI, real.default); // keep the error classes for friendlyError
  return { default: FakeOpenAI };
});

const { structured } = await import("./client.js");
const schema = z.object({ summary: z.string(), priority: z.enum(["high", "low"]) });
const reply = (content: string, finish = "stop") => ({
  choices: [{ message: { content }, finish_reason: finish }],
  usage: { total_tokens: 10 },
});

describe("structured", () => {
  it("parses fenced JSON and sends the schema in json mode", async () => {
    create.mockReset().mockResolvedValueOnce(reply('```json\n{"summary":"s","priority":"high"}\n```'));
    const out = await structured(schema, "sys", "user");
    expect(out).toEqual({ data: { summary: "s", priority: "high" }, tokens: 10 });
    const req = create.mock.calls[0][0];
    expect(req.response_format).toEqual({ type: "json_object" });
    expect(req.max_completion_tokens).toBe(8000); // GPT-5 / o-series reject max_tokens
    expect(req.messages[0].content).toContain('"priority"');
  });
  it("retries once on output that fails the schema, then gives up", async () => {
    create
      .mockReset()
      .mockResolvedValueOnce(reply('{"summary":"s","priority":"urgent"}'))
      .mockResolvedValueOnce(reply("{"));
    await expect(structured(schema, "sys", "user")).rejects.toThrow("unexpected format");
    expect(create).toHaveBeenCalledTimes(2);
  });
  it("reports truncated output", async () => {
    create.mockReset().mockResolvedValueOnce(reply('{"summary":', "length"));
    await expect(structured(schema, "sys", "user")).rejects.toThrow("cut off");
  });
});
