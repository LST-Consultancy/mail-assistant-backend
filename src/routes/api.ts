import { Router } from "express";
import { config } from "../config.js";
import { friendlyError, streamText, structured, text } from "../llm/client.js";
import {
  AnalysisSchema,
  AnalyzeEmailRequest,
  BatchRequest,
  ComposeRequest,
  DigestItemSchema,
  OverviewRequest,
  ReplyRequest,
} from "../llm/schemas.js";
import { analyzeSystem, analyzeUser, digestSystem, digestUser, overviewSystem } from "../prompts/analyze.js";
import { composeSystem, composeUser, replySystem, replyUser } from "../prompts/reply.js";
import { sse } from "./sse.js";

export const api = Router();

api.post("/analyze-email", async (req, res) => {
  const { email, thread } = AnalyzeEmailRequest.parse(req.body);
  const { data, tokens } = await structured(AnalysisSchema, analyzeSystem, analyzeUser(email, thread));
  res.locals.tokens = tokens;
  res.json(data);
});

api.post("/reply", async (req, res) => {
  const b = ReplyRequest.parse(req.body);
  await sse(res, (write, signal) =>
    streamText(
      replySystem(b.tone, b.replyType, b.settings),
      replyUser(b.email, b.thread, b.instructions),
      write,
      signal,
    ),
  );
});

api.post("/compose", async (req, res) => {
  const b = ComposeRequest.parse(req.body);
  await sse(res, (write, signal) =>
    streamText(composeSystem(b.mode, b.tone, b.settings), composeUser(b), write, signal),
  );
});

/** Each thread is its own AI call, so one failure never sinks the batch. */
api.post("/digest/analyze-batch", async (req, res) => {
  const { threads, settings } = BatchRequest.parse(req.body);
  const system = digestSystem(settings);
  const results = await mapLimit(threads, config.llmConcurrency, async (t) => {
    try {
      const { data, tokens } = await structured(DigestItemSchema, system, digestUser(t), 4000);
      res.locals.tokens += tokens;
      return { id: t.id, ...data };
    } catch (err) {
      return { id: t.id, error: friendlyError(err).message };
    }
  });
  res.json(results);
});

api.post("/digest/overview", async (req, res) => {
  const { items } = OverviewRequest.parse(req.body);
  const list = items
    .map(
      (i) =>
        `- [${i.priority}/${i.category}${i.needsReply ? "/needs reply" : ""}] ${i.from} - "${i.subject}": ${i.summary}`,
    )
    .join("\n");
  const { text: markdown, tokens } = await text(overviewSystem, `<email>\nTriaged items:\n${list}\n</email>`);
  res.locals.tokens = tokens;
  res.json({ markdown });
});

export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const i = next++;
      out[i] = await fn(items[i]);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return out;
}
