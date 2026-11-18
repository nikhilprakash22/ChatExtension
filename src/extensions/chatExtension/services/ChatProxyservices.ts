import { SPHttpClient } from '@microsoft/sp-http';
import { IChatConfig } from '../models/IChatConfig';
import { IChatMessage } from '../models/IChatMessage';
import { AuthService } from '../services/AuthService';
import { ConfigService } from './ConfigService';
import { RbacService } from './rbacService';
import { copilotChatService } from './CopilotChatService';

export interface IChatProxyInitResult {
  allowed:  boolean;
  config:   IChatConfig | null;
  welcome:  IChatMessage[];
  errorMsg: string;
}

export class ChatProxyService {
  private _spHttpClient:      SPHttpClient;
  private _hubSiteUrl:        string;
  private _listTitle:         string;
  private _siteId:            string;   // GUID — used for ConfigService list lookup
  private _siteUrl:           string;   // absolute URL — used for RbacService REST calls
  private _defaultTitleMatch: string;
  private _config:            IChatConfig | null = null;

  constructor(
    spHttpClient:      SPHttpClient,
    hubSiteUrl:        string,
    listTitle:         string,
    siteId:            string,
    siteUrl:           string,
    defaultTitleMatch: string
  ) {
    this._spHttpClient      = spHttpClient;
    this._hubSiteUrl        = hubSiteUrl;
    this._listTitle         = listTitle;
    this._siteId            = siteId;
    this._siteUrl           = siteUrl;
    this._defaultTitleMatch = defaultTitleMatch;
  }

  public async initialize(): Promise<IChatProxyInitResult> {
    // ── Step 1: Fetch config from SP list ─────────────────────────────────────
    let config = null;
    try {
      config = await ConfigService.getConfig(
        this._spHttpClient,
        this._hubSiteUrl,
        this._listTitle,
        this._siteId,
        this._defaultTitleMatch
      );
    } catch (err) {
      const msg = (err instanceof Error) ? err.message : String(err);
      return {
        allowed: false, config: null, welcome: [],
        // SECURITY: hubSiteUrl, listTitle, siteId excluded — internal infrastructure details.
        errorMsg: '[ERR-LIST-THROW] Could not read configuration list. Check hub site URL and list permissions. Detail: ' + msg,
      };
    }

    if (!config) {
      return {
        allowed: false, config: null, welcome: [],
        // SECURITY: hubSiteUrl, siteId, list internals excluded from UI error message.
        errorMsg: '[ERR-CONFIG] No configuration row found for this site. ' +
          'Check: Does a "' + this._defaultTitleMatch + '" row exist in the config list? ' +
          'Is the SiteId column set correctly?',
      };
    }

    if (!config.enabled) {
      return {
        allowed: false, config, welcome: [],
        errorMsg: '[ERR-DISABLED] Chatbot is disabled for this site (Enabled column = false).',
      };
    }

    // ── Step 2: RBAC check ────────────────────────────────────────────────────
    let allowed = false;
    try {
      allowed = await RbacService.isUserAllowed(
        this._spHttpClient,
        this._siteUrl,
        config
      );
    } catch (err) {
      const msg = (err instanceof Error) ? err.message : String(err);
      return {
        allowed: false, config, welcome: [],
        errorMsg: '[ERR-RBAC-THROW] RBAC check threw: ' + msg,
      };
    }

    if (!allowed) {
      return {
        allowed: false, config, welcome: [],
        // SECURITY: AllowedRoles value excluded — don't reveal group names to unauthorised users.
        errorMsg: '[ERR-RBAC] Access not granted for your account. Contact the site owner.',
      };
    }

    // ── Step 3: MSAL init ─────────────────────────────────────────────────────
    try {
      await AuthService.getInstance().initialize(config.appClientId, config.tenantId);
    } catch {
      return {
        allowed: false, config, welcome: [],
        // SECURITY: raw error detail and appClientId excluded from the UI message.
        errorMsg: '[ERR-MSAL] MSAL initialization failed. Check AppClientId and TenantId in the SP list row.',
      };
    }

    this._config = config;

    // ── Step 4: Start Copilot session ─────────────────────────────────────────
    let welcome: IChatMessage[] = [];
    try {
      welcome = await copilotChatService.startSession(
        config,
        () => AuthService.getInstance().acquireToken()
      );
    } catch (err) {
      const msg = (err instanceof Error) ? err.message : String(err);
      // Session start failure is non-fatal — surface as a bot error message
      // so the user can still type (the agent may respond on first send)
      welcome = [{
        id:          'init-err-' + Date.now(),
        sender:      'system',
        contentType: 'error',
        text:        '[ERR-SESSION] Could not start session: ' + msg,
        timestamp:   new Date(),
      }];
    }

    return { allowed: true, config, welcome, errorMsg: '' };
  }

  public async send(text: string): Promise<IChatMessage[]> {
    if (!this._config) {
      return [{
        id:          this._uid(),
        sender:      'system',
        contentType: 'error',
        text:        'Chat not initialized.',
        timestamp:   new Date(),
      }];
    }

    return copilotChatService.sendMessage(
      text,
      this._config,
      () => AuthService.getInstance().acquireToken()
    );
  }

  public async reset(): Promise<IChatMessage[]> {
    if (!this._config) { return []; }
    ConfigService.clearCache();
    await copilotChatService.endSession();
    return copilotChatService.startSession(
      this._config,
      () => AuthService.getInstance().acquireToken()
    );
  }

  private _uid(): string {
    return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
  }
}