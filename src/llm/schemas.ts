import { z } from "zod";
import { config } from "../config.js";

// ---- Requests (trust boundary: everything from the add-in is validated here) ----

const str = (max: number) =>
  z
    .string()
    .max(max * 2)
    .transform((s) => s.slice(0, max));
const address = z.object({ name: str(200).default(""), address: str(320).default("") });

export const EmailSchema = z.object({
  id: str(500).default(""),
  subject: str(1000).default(""),
  from: address,
  to: z.array(address).max(100).default([]),
  cc: z.array(address).max(100).default([]),
  receivedDateTime: str(64).default(""),
  body: z
    .string()
    .max(200_000)
    .transform((s) => s.slice(0, config.maxBodyChars)),
  importance: str(16).optional(),
});
export type Email = z.infer<typeof EmailSchema>;

export const TONES = ["professional", "friendly", "short", "formal"] as const;

export const SettingsSchema = z
  .object({
    signatureName: str(200).optional(),
    /** "match" = reply in the sender's language, otherwise a language name. */
    language: str(50).default("match"),
  })
  .default({ language: "match" });
export type Settings = z.infer<typeof SettingsSchema>;

const thread = z.array(EmailSchema).max(20).optional();

export const AnalyzeEmailRequest = z.object({ email: EmailSchema, thread, settings: SettingsSchema });

export const ReplyRequest = z.object({
  email: EmailSchema,
  thread,
  tone: z.enum(TONES).default("professional"),
  replyType: z.enum(["reply", "replyAll"]).default("reply"),
  instructions: str(2000).optional(),
  settings: SettingsSchema,
});

export const BatchRequest = z.object({ threads: z.array(EmailSchema).min(1).max(10), settings: SettingsSchema });

export const OverviewRequest = z.object({
  items: z
    .array(
      z.object({
        subject: str(500),
        from: str(400),
        summary: str(1500),
        priority: str(16),
        category: str(32),
        needsReply: z.boolean().optional(),
      }),
    )
    .min(1)
    .max(500),
});

export const ComposeRequest = z.object({
  mode: z.enum(["draft", "improve"]),
  text: z
    .string()
    .max(100_000)
    .transform((s) => s.slice(0, config.maxBodyChars * 2))
    .optional(),
  instructions: str(2000).optional(),
  context: EmailSchema.optional(),
  tone: z.enum(TONES).default("professional"),
  settings: SettingsSchema,
});

// ---- Model outputs (validated by the SDK's zod parser) ----

export const CATEGORIES = [
  "action_required",
  "question",
  "fyi",
  "meeting",
  "newsletter_or_promo",
  "notification",
  "personal",
  "other",
] as const;
const priority = z.enum(["high", "medium", "low"]);

export const AnalysisSchema = z.object({
  summary: z.string(),
  keyPoints: z.array(z.string()),
  actionItems: z.array(z.object({ text: z.string(), deadline: z.string().nullable() })),
  category: z.enum(CATEGORIES),
  priority,
  needsReply: z.boolean(),
});

export const DigestItemSchema = z.object({
  summary: z.string(),
  category: z.enum(CATEGORIES),
  priority,
  needsReply: z.boolean(),
  replyDraft: z.string().describe("Empty string when needsReply is false"),
});
