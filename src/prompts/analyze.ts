import type { Email, Settings } from "../llm/schemas.js";
import { formatConversation, formatEmail, languageRule, SAFETY, signatureRule } from "./shared.js";

export const analyzeSystem = `${SAFETY}

Task: analyze the email for the user.
- summary: 2-4 sentences covering what the thread is about and where it stands now.
- keyPoints: the important facts, decisions and requests, one short line each (max 6).
- actionItems: what the USER must do, with any deadline stated in the email (null if none). Empty if nothing.
- category, priority (high = needs action soon or comes from someone waiting on the user; low = newsletters, notifications), needsReply (true only if a human expects an answer from the user).
Write the analysis in the same language as the email.`;

export function analyzeUser(email: Email, thread?: Email[]): string {
  return formatConversation(email, thread);
}

export const digestSystem = (s: Settings) => `${SAFETY}

Task: triage one email thread from the user's inbox.
- summary: 1-2 sentences.
- category and priority (high = someone is waiting on the user or there is a deadline; low = newsletters, promos, automated notifications).
- needsReply: true only if a human expects an answer from the user.
- replyDraft: if needsReply, a short ready-to-edit reply body (no subject line). ${languageRule(s)} ${signatureRule(s)} Otherwise "".`;

export function digestUser(email: Email): string {
  return formatEmail(email);
}

export const overviewSystem = `${SAFETY}

Task: write an executive overview of the user's inbox from the triaged items below (each is an AI summary of one thread). Markdown, under 250 words, with these sections:
## Key themes
## Waiting on you
(who is waiting and for what; most urgent first)
## Time-sensitive
(deadlines and dates mentioned)
Skip a section's bullets with "None." if empty. Do not list newsletters or notifications individually.`;
