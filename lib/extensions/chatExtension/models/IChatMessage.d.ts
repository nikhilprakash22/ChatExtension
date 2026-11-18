export type MessageSender = 'user' | 'bot' | 'system';
export type MessageContentType = 'text' | 'adaptiveCard' | 'error';
export interface ISuggestedAction {
    title: string;
    value: string;
}
export interface IChatMessage {
    id: string;
    sender: MessageSender;
    contentType: MessageContentType;
    text?: string;
    cardPayload?: any;
    timestamp: Date;
    suggestedActions?: ISuggestedAction[];
}
//# sourceMappingURL=IChatMessage.d.ts.map