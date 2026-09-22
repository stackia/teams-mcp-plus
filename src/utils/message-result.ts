import type { ChatMessage } from "../types/graph.js";
import { extractAttachmentSummaries } from "./attachments.js";
import { formatMessageContent } from "./html-to-markdown.js";

/** Keep the list response envelope when fetching an individual message. */
export function singleMessageResult(message: ChatMessage, format: "raw" | "markdown") {
  if (!message?.id) throw new Error("Message not found.");
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(
          {
            totalReturned: 1,
            hasMore: false,
            messages: [
              {
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
              },
            ],
          },
          null,
          2
        ),
      },
    ],
  };
}
