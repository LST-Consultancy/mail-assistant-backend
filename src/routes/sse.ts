import type { Response } from "express";
import { friendlyError } from "../llm/client.js";

/** Relays a text stream as SSE: `data: {"text":…}` chunks, then `event: done` or `event: error`. */
export async function sse(
  res: Response,
  run: (write: (text: string) => void, signal: AbortSignal) => Promise<number>,
): Promise<void> {
  res.writeHead(200, {
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache",
    "X-Accel-Buffering": "no",
  });
  const ac = new AbortController();
  res.on("close", () => ac.abort()); // user closed the pane or hit Stop
  try {
    res.locals.tokens = await run((text) => res.write(`data: ${JSON.stringify({ text })}\n\n`), ac.signal);
    res.write("event: done\ndata: {}\n\n");
  } catch (err) {
    if (!ac.signal.aborted) {
      res.locals.errorKind = (err as Error)?.constructor?.name;
      res.write(`event: error\ndata: ${JSON.stringify({ error: friendlyError(err).message })}\n\n`);
    }
  }
  res.end();
}
