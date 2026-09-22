import type { ChatMessage } from "../types/graph.js";
import { extractAttachmentSummaries } from "./attachments.js";
import { formatMessageContent } from "./html-to-markdown.js";

function messageSummary(message: ChatMessage, format: "raw" | "markdown") {
  if (!message?.id) throw new Error("Message not found.");
  return {
    id: message.id,
    replyToId: message.replyToId,
    content: formatMessageContent(
      message.body?.content,
      message.body?.contentType === "text" ? "raw" : format,
      message.mentions
    ),
    from: message.from?.user?.displayName ?? message.from?.application?.displayName,
    sender: message.from,
    createdDateTime: message.createdDateTime,
    lastModifiedDateTime: message.lastModifiedDateTime,
    deletedDateTime: message.deletedDateTime,
    subject: message.subject,
    importance: message.importance,
    webUrl: message.webUrl,
    mentions: message.mentions,
    attachments: extractAttachmentSummaries(message.attachments),
    reactions: message.reactions,
  };
}

/** Keep the list response envelope when fetching an individual message. */
export function singleMessageResult(message: ChatMessage, format: "raw" | "markdown") {
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(
          {
            totalReturned: 1,
            hasMore: false,
            messages: [messageSummary(message, format)],
          },
          null,
          2
        ),
      },
    ],
  };
}

interface MessageTarget {
  messageId: string;
  replyId?: string;
}

/** Fetch sequentially, retaining successful reads and identifying each failed target. */
export async function batchMessageResult(
  targets: MessageTarget[],
  fetchMessage: (target: MessageTarget) => Promise<ChatMessage>,
  format: "raw" | "markdown"
) {
  const messages: ReturnType<typeof messageSummary>[] = [];
  const errors: (MessageTarget & { error: string })[] = [];
  for (const target of targets) {
    try {
      messages.push(messageSummary(await fetchMessage(target), format));
    } catch (error) {
      errors.push({
        ...target,
        error: error instanceof Error ? error.message : "Unknown error occurred",
      });
    }
  }
  return {
    ...(messages.length === 0 && errors.length > 0 ? { isError: true } : {}),
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(
          {
            totalRequested: targets.length,
            totalReturned: messages.length,
            hasMore: false,
            messages,
            errors,
          },
          null,
          2
        ),
      },
    ],
  };
}
