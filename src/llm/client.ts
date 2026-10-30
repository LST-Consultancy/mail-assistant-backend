import OpenAI from "openai";
import { z } from "zod";
import { config } from "../config.js";

// OpenAI API client. The SDK retries 408/409/429/5xx with exponential backoff.
const llm = new OpenAI({
  baseURL: config.llmBaseUrl,
  apiKey: config.llmApiKey || "none", // The SDK requires a value even when misconfigured
  maxRetries: 4,
  timeout: 120_000,
});

export class AiError extends Error {
  constructor(
    message: string,
    public status = 502,
  ) {
    super(message);
  }
}

type Usage = { total_tokens: number } | null | undefined;
// shortcut: if a provider omits usage, estimate ~4 chars/token so the daily budget still moves.
const tokensOf = (usage: Usage, chars: number) => usage?.total_tokens ?? Math.ceil(chars / 4);

function checkFinish(reason: string | null | undefined) {
  if (reason === "content_filter") throw new AiError("The AI declined to process this email.", 422);
  if (reason === "length") throw new AiError("The AI response was cut off. Try a shorter thread.");
}

const messages = (system: string, user: string): OpenAI.ChatCompletionMessageParam[] => [
  { role: "system", content: system },
  { role: "user", content: user },
];

function parseJson(text: string | null | undefined): unknown {
  try {
    // Some models wrap JSON in ```json fences despite json mode.
    return JSON.parse((text ?? "").replace(/^\s*```(?:json)?\s*|\s*```\s*$/g, ""));
  } catch {
    return undefined;
  }
}

/** JSON mode + the schema in the prompt, validated with zod; one retry on invalid output. */
export async function structured<S extends z.ZodType>(
  schema: S,
  system: string,
  user: string,
  maxTokens = 8000,
): Promise<{ data: z.infer<S>; tokens: number }> {
  const sys = `${system}\n\nRespond with only a JSON object that matches this JSON Schema:\n${JSON.stringify(z.toJSONSchema(schema))}`;
  let tokens = 0;
  for (let attempt = 0; ; attempt++) {
    const res = await llm.chat.completions.create({
      model: config.llmModel,
      max_completion_tokens: maxTokens,
      response_format: { type: "json_object" },
      messages: messages(sys, user),
    });
    const choice = res.choices[0];
    tokens += tokensOf(res.usage, sys.length + user.length + (choice?.message.content?.length ?? 0));
    checkFinish(choice?.finish_reason);
    const parsed = schema.safeParse(parseJson(choice?.message.content));
    if (parsed.success) return { data: parsed.data, tokens };
    if (attempt >= 1) throw new AiError("The AI returned an unexpected format.");
  }
}

export async function text(system: string, user: string, maxTokens = 8000): Promise<{ text: string; tokens: number }> {
  const res = await llm.chat.completions.create({
    model: config.llmModel,
    max_completion_tokens: maxTokens,
    messages: messages(system, user),
  });
  const choice = res.choices[0];
  checkFinish(choice?.finish_reason);
  const out = choice?.message.content ?? "";
  return { text: out, tokens: tokensOf(res.usage, system.length + user.length + out.length) };
}

/** Streams text deltas to onText; resolves with total tokens used. */
export async function streamText(
  system: string,
  user: string,
  onText: (delta: string) => void,
  signal: AbortSignal,
): Promise<number> {
  const stream = await llm.chat.completions.create(
    {
      model: config.llmModel,
      max_completion_tokens: 8000, // reasoning models spend part of this on hidden reasoning
      stream: true,
      stream_options: { include_usage: true },
      messages: messages(system, user),
    },
    { signal },
  );
  let chars = system.length + user.length;
  let usage: Usage;
  let finish: string | null | undefined;
  for await (const chunk of stream) {
    const delta = chunk.choices[0]?.delta?.content;
    if (delta) {
      chars += delta.length;
      onText(delta);
    }
    finish = chunk.choices[0]?.finish_reason ?? finish;
    usage = chunk.usage ?? usage;
  }
  checkFinish(finish);
  return tokensOf(usage, chars);
}

/** Maps SDK errors to a status + message safe to show in the task pane. */
export function friendlyError(err: unknown): { status: number; message: string } {
  if (err instanceof AiError) return { status: err.status, message: err.message };
  if (err instanceof OpenAI.AuthenticationError || err instanceof OpenAI.PermissionDeniedError)
    return { status: 503, message: "The AI service is not configured correctly. Contact your administrator." };
  if (err instanceof OpenAI.NotFoundError)
    return { status: 503, message: "The configured AI model was not found. Check LLM_MODEL." };
  if (err instanceof OpenAI.RateLimitError)
    return { status: 503, message: "The AI provider's rate limit was reached. Please retry in a minute." };
  if (err instanceof OpenAI.APIError && (err.status ?? 0) >= 500)
    return { status: 503, message: "The AI service is temporarily overloaded. Please retry." };
  if (err instanceof OpenAI.APIConnectionError) return { status: 503, message: "Couldn't reach the AI service." };
  return { status: 500, message: "Something went wrong while processing your request." };
}
