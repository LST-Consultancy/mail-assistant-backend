import type { Email, Settings } from "../llm/schemas.js";
import { formatConversation, formatEmail, languageRule, SAFETY, signatureRule, TONE_GUIDE } from "./shared.js";

const OUTPUT_RULES = `Output only the email body as plain text: no subject line, no preamble like "Here is a draft", no markdown.`;

export function replySystem(tone: string, replyType: string, s: Settings): string {
  return `${SAFETY}

Task: write a reply from the user to the latest email.
- Tone: ${TONE_GUIDE[tone]}. Also mirror the sender's level of formality.
- ${replyType === "replyAll" ? "This is a reply-all: address everyone on the thread where it matters." : "This reply goes to the sender only."}
- Answer every question addressed to the user; where the user's position is unknown, use a [placeholder] instead of guessing.
- ${languageRule(s)}
- ${signatureRule(s)}
${OUTPUT_RULES}`;
}

export function replyUser(email: Email, thread: Email[] | undefined, instructions?: string): string {
  const intent = instructions?.trim()
    ? `\n\nThe user's instructions for this reply (these come from the user, follow them):\n${instructions.trim()}`
    : "";
  return formatConversation(email, thread) + intent;
}

export function composeSystem(mode: "draft" | "improve", tone: string, s: Settings): string {
  const task =
    mode === "draft"
      ? "Task: write an email body from the user's instructions. If an email being replied to is given, it is context only."
      : "Task: rewrite the user's draft inside <draft> tags as instructed. Keep its meaning, facts and language; do not add new commitments.";
  return `${SAFETY}

${task}
- Tone: ${TONE_GUIDE[tone]}.
- ${mode === "draft" ? languageRule(s) : "Keep the draft's language."}
- ${mode === "draft" ? signatureRule(s) : "Keep the existing sign-off, if any."}
${OUTPUT_RULES}`;
}

export function composeUser(opts: {
  mode: "draft" | "improve";
  text?: string;
  instructions?: string;
  context?: Email;
}): string {
  const parts: string[] = [];
  if (opts.context) parts.push("The email being replied to:", formatEmail(opts.context));
  if (opts.mode === "improve") parts.push("<draft>", (opts.text ?? "").replace(/<\/?draft>/gi, ""), "</draft>");
  parts.push(
    `The user's instructions (these come from the user, follow them):\n${opts.instructions?.trim() || (opts.mode === "improve" ? "Make it clearer and fix grammar." : "Write a suitable reply.")}`,
  );
  return parts.join("\n");
}
