import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client';
import { IChatMessage, ISuggestedAction } from '../models/IChatMessage';
import { IChatConfig } from '../models/IChatConfig';

/**
 * CopilotChatService
 * Wraps @microsoft/agents-copilotstudio-client v1.3.1.
 *
 * Verified API (from official Microsoft docs and npm page):
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
 *     NOTE: the old 'botId' field does NOT exist. Use agentIdentifier or schemaName.
 *
 *   Methods (Promise-based, not streaming — most compatible with ES5 + SPFx):
 *     startConversationAsync(emitStartConversationEvent?: boolean): Promise<Activity[]>
 *     askQuestionAsync(question: string, conversationId?: string):  Promise<Activity[]>
 *
 *   Do NOT use:
 *     startConversation()   — does not exist
 *     startConversationStreaming() — AsyncGenerator, needs ES2018 lib + downlevel iteration
 *     askQuestion()         — does not exist
 *     connect()             — does not exist
 *
 * SECURITY: No token value is ever stored as a class property.
 *
 * ── CHANGE LOG (suggested actions) ──────────────────────────────────────────
 * _activitiesToMessages now reads activity.suggestedActions.actions[] and
 * attaches the chips to the LAST text message in the activity batch.
 * No other logic was changed.
 */
export class CopilotChatService {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _client: any | null = null;
  private _sessionActive: boolean = false;

  /**
   * Start a new conversation and return the bot's welcome messages.
   */
  public async startSession(
    config: IChatConfig,
    tokenGetter: () => Promise<string>
  ): Promise<IChatMessage[]> {
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
    if (!this._client || !this._sessionActive) {
      await this._initClient(config, tokenGetter);
      this._sessionActive = true;
    }
    return this._askQuestion(text);
  }

  /** Reset — tear down so the next call starts a fresh conversation. */
  public async endSession(): Promise<void> {
    this._client = null;
    this._sessionActive = false;
  }

  // ─── Private ────────────────────────────────────────────────────────────────

  private async _initClient(
    config: IChatConfig,
    tokenGetter: () => Promise<string>
  ): Promise<void> {
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
      return CopilotChatService._activitiesToMessages(activities);
    } catch (err) {
      const msg = (err instanceof Error) ? err.message : String(err);
      return [{
        id:          CopilotChatService._uid(),
        sender:      'system',
        contentType: 'error',
        text:        'Could not start conversation: ' + msg,
        timestamp:   new Date(),
      }];
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
      return [{
        id:          CopilotChatService._uid(),
        sender:      'system',
        contentType: 'error',
        text:        'Could not send message: ' + msg,
        timestamp:   new Date(),
      }];
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private static _activitiesToMessages(activities: any[]): IChatMessage[] {
    const results: IChatMessage[] = [];
    if (!activities || !Array.isArray(activities)) { return results; }

    // ── NEW: collect suggested actions from ALL activities in the batch ───────
    // Copilot Studio places suggestedActions on the last message activity.
    // We gather them here and attach them to the last text message we produce.
    const collectedActions: ISuggestedAction[] = [];

    for (let i = 0; i < activities.length; i++) {
      const activity = activities[i];
      if (!activity || activity.type !== 'message') { continue; }

      // ── NEW: harvest suggestedActions from this activity ───────────────────
      // Shape: activity.suggestedActions.actions = [{ title, value, type }, ...]
      // We only read title + value; type is always "imBack" for quick replies.
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
            collectedActions.push({ title, value });
          }
        }
      }
      // ── END NEW ──────────────────────────────────────────────────────────────

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

    // ── NEW: attach collected suggested actions to the last bot message ───────
    // This mirrors how Copilot Studio web chat works — chips appear below
    // the final bot bubble in the batch, not repeated on each message.
    if (collectedActions.length > 0 && results.length > 0) {
      // Find the last bot message (text or adaptiveCard) to attach chips to
      for (let r = results.length - 1; r >= 0; r--) {
        if (results[r].sender === 'bot') {
          results[r].suggestedActions = collectedActions;
          break;
        }
      }
    }
    // ── END NEW ──────────────────────────────────────────────────────────────

    return results;
  }

  private static _uid(): string {
    return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
  }
}

export const copilotChatService = new CopilotChatService();