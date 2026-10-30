const env = process.env;

const list = (v: string | undefined) =>
  (v ?? "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

export const config = {
  port: Number(env.PORT ?? 8080),
  // OpenAI API.
  llmBaseUrl: env.LLM_BASE_URL || "https://api.openai.com/v1",
  llmApiKey: env.LLM_API_KEY ?? "",
  llmModel: env.LLM_MODEL || "gpt-4o-mini",
  /** Parallel AI calls per digest batch. */
  llmConcurrency: Number(env.LLM_CONCURRENCY ?? 5),
  entraClientId: env.ENTRA_CLIENT_ID ?? "",
  allowedTenantIds: list(env.ALLOWED_TENANT_IDS),
  addinOrigins: list(env.ADDIN_ORIGIN || "https://localhost:3000"),
  dailyTokenBudget: Number(env.DAILY_TOKEN_BUDGET_PER_USER ?? 500_000),
  requestsPerMinute: Number(env.REQUESTS_PER_MINUTE ?? 60),
  /** Accepts the literal token "dev" (used by the add-in's VITE_MOCK mode). Never in production. */
  devAuthBypass: env.DEV_AUTH_BYPASS === "1" && env.NODE_ENV !== "production",
  maxBodyChars: Number(env.MAX_BODY_CHARS ?? 6000),
};
