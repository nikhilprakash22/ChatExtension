import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client';
import { IChatMessage, ISuggestedAction } from '../models/IChatMessage';
import { IChatConfig } from '../models/IChatConfig';

/**
 * CopilotChatService
 * Wraps @microsoft/agents-copilotstudio-client v1.3.1.
 *
 *   Constructor:
 *     new CopilotStudioClient(settings: ConnectionSettings, token: string)
 *
 *   ConnectionSettings shape:
 *     {
 *       environmentId:    string   // Power Platform environment GUID
 *       agentIdentifier:  string   // Agent schema name (e.g. "cr123_myAgent")
 *       directConnectUrl: string   // Optional — overrides environmentId/agentIdentifier
 *     }
 *
 *   Methods (Promise-based):
 *     startConversationAsync(emitStartConversationEvent?: boolean): Promise<Activity[]>
 *     askQuestionAsync(question: string, conversationId?: string):  Promise<Activity[]>
 *
 * SECURITY: No token value is ever stored as a class property or logged.
 *
 * FIX: askQuestionAsync needs a started conversation (conversation.id).
 *      If the session start failed (e.g. sign-in popup cancelled), the next
 *      sendMessage now signs in and starts the conversation first.
 */
export class CopilotChatService {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _client: any | null = null;
  private _sessionActive: boolean = false;
  private _conversationStarted: boolean = false; // FIX: true only after startConversationAsync succeeded

  /**
   * Start a new conversation and return the bot's welcome messages.
   */
  public async startSession(
    config: IChatConfig,
    tokenGetter: () => Promise<string>
  ): Promise<IChatMessage[]> {
    this._conversationStarted = false; // FIX
    await this._initClient(config, tokenGetter);
    this._sessionActive = true;
    return this._startConversation();
  }

  /**
   * Send a user message and return the bot's reply messages.
   */
  public async sendMessage(
    text: string,
    config: IChatConfig,
    tokenGetter: () => Promise<string>
  ): Promise<IChatMessage[]> {
    // FIX: no client or no started conversation → sign in and start the conversation first
    if (!this._client || !this._sessionActive || !this._conversationStarted) {
      try {
        await this._initClient(config, tokenGetter);
      } catch (err) {
        const msg = (err instanceof Error) ? err.message : String(err);
        return [CopilotChatService._error('Could not sign in: ' + msg)];
      }
      this._sessionActive = true;

      const welcome: IChatMessage[] = await this._startConversation();
      if (!this._conversationStarted) {
        return welcome; // contains the error message; never ask without a conversation
      }

      const replies: IChatMessage[] = await this._askQuestion(text);
      return welcome.concat(replies);
    }

    return this._askQuestion(text);
  }

  /** Reset — tear down so the next call starts a fresh conversation. */
  public async endSession(): Promise<void> {
    this._client = null;
    this._sessionActive = false;
    this._conversationStarted = false; // FIX
  }

  // ─── Private ────────────────────────────────────────────────────────────────

  private async _initClient(
    config: IChatConfig,
    tokenGetter: () => Promise<string>
  ): Promise<void> {
    // SECURITY: token is a local variable only — passed to the client, never stored here
    const token: string = await tokenGetter();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const settings: any = {
      agentIdentifier: config.schemaName,
      environmentId:   config.environmentId,
    };

    if (config.directConnectUrl) {
      settings.directConnectUrl = config.directConnectUrl;
    }

    this._client = new CopilotStudioClient(settings, token);
  }

  private async _startConversation(): Promise<IChatMessage[]> {
    if (!this._client) { return []; }
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const activities: any[] = await this._client.startConversationAsync(true);
      this._conversationStarted = true; // FIX
      return CopilotChatService._activitiesToMessages(activities);
    } catch (err) {
      this._conversationStarted = false; // FIX
      const msg = (err instanceof Error) ? err.message : String(err);
      return [CopilotChatService._error('Could not start conversation: ' + msg)];
    }
  }

  private async _askQuestion(text: string): Promise<IChatMessage[]> {
    if (!this._client) { return []; }
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const activities: any[] = await this._client.askQuestionAsync(text);
      return CopilotChatService._activitiesToMessages(activities);
    } catch (err) {
      const msg = (err instanceof Error) ? err.message : String(err);
      return [CopilotChatService._error('Could not send message: ' + msg)];
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private static _activitiesToMessages(activities: any[]): IChatMessage[] {
    const results: IChatMessage[] = [];
    if (!activities || !Array.isArray(activities)) { return results; }

    // Collect suggested actions from ALL activities in the batch.
    const collectedActions: ISuggestedAction[] = [];

    for (let i = 0; i < activities.length; i++) {
      const activity = activities[i];
      if (!activity || activity.type !== 'message') { continue; }

      if (
        activity.suggestedActions &&
        Array.isArray(activity.suggestedActions.actions)
      ) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const raw: any[] = activity.suggestedActions.actions;
        for (let a = 0; a < raw.length; a++) {
          const act = raw[a];
          const title: string = (act.title || act.value || '').toString().trim();
          const value: string = (act.value || act.title || '').toString().trim();
          if (title) {
            collectedActions.push({ title: title, value: value });
          }
        }
      }

      if (activity.text) {
        results.push({
          id:          CopilotChatService._uid(),
          sender:      'bot',
          contentType: 'text',
          text:        activity.text,
          timestamp:   new Date(),
        });
      }

      if (activity.attachments && Array.isArray(activity.attachments)) {
        for (let j = 0; j < activity.attachments.length; j++) {
          const att = activity.attachments[j];
          if (att.contentType === 'application/vnd.microsoft.card.adaptive') {
            results.push({
              id:          CopilotChatService._uid(),
              sender:      'bot',
              contentType: 'adaptiveCard',
              cardPayload: att.content,
              timestamp:   new Date(),
            });
          }
        }
      }
    }

    // Attach collected suggested actions to the last bot message.
    if (collectedActions.length > 0 && results.length > 0) {
      for (let r = results.length - 1; r >= 0; r--) {
        if (results[r].sender === 'bot') {
          results[r].suggestedActions = collectedActions;
          break;
        }
      }
    }

    return results;
  }

  /** FIX: one place to build an error message bubble. */
  private static _error(text: string): IChatMessage {
    return {
      id:          CopilotChatService._uid(),
      sender:      'system',
      contentType: 'error',
      text:        text,
      timestamp:   new Date(),
    };
  }

  private static _uid(): string {
    return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
  }
}

export const copilotChatService = new CopilotChatService();
