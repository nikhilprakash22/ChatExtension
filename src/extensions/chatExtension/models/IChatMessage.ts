export type MessageSender = 'user' | 'bot' | 'system';
export type MessageContentType = 'text' | 'adaptiveCard' | 'error';

// ── NEW: one quick-reply chip returned by Copilot Studio ─────────────────────
export interface ISuggestedAction {
  title:   string; // display label on the chip
  value:   string; // text sent when the chip is tapped (may differ from title)
}

export interface IChatMessage {
  id:          string;
  sender:      MessageSender;
  contentType: MessageContentType;
  text?:       string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  cardPayload?: any;
  timestamp:   Date;
  // ── NEW: optional suggested-action chips attached to this message ──────────
  // Populated by CopilotChatService when the activity contains suggestedActions.
  // Rendered as tappable chips below the bot bubble.
  // Cleared (set to []) after any chip is tapped or any new message is sent.
  suggestedActions?: ISuggestedAction[];
}