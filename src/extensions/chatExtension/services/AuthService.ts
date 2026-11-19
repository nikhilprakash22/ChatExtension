/**
 * AuthService
 * Handles MSAL token acquisition for SPFx.
 *
 * CRITICAL SPFx RULE:
 *   acquireTokenRedirect MUST NOT be used in SPFx extensions.
 *   A full-page redirect navigates away from SharePoint, destroys the SPFx
 *   context, and returns to a blank page with no extension loaded.
 *
 *   Strategy:
 *     1. acquireTokenSilent  — from the in-memory cache (no UI, instant)
 *     2. ssoSilent           — reuses the user's existing Microsoft sign-in
 *                              in a hidden frame (no UI)
 *     3. acquireTokenPopup   — only when 1 and 2 fail (small popup window,
 *                              SharePoint page stays intact)
 *
 * Scope: https://api.powerplatform.com/CopilotStudio.Copilots.Invoke
 *
 * AAD App Registration requirements:
 *   - Platform: Single-page application (SPA)
 *   - Redirect URIs (both required):
 *       https://<tenant>.sharepoint.com/_forms/spfxsinglesignon.aspx   (production)
 *       https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx     (workbench)
 *   - API permission: Power Platform API → Delegated → CopilotStudio.Copilots.Invoke
 *
 * SECURITY:
 *   - Tokens are kept in memory only (memoryStorage) — never written to
 *     localStorage, sessionStorage or cookies.
 *   - Tokens are never logged or exposed to the DOM.
 *   - MSAL logging is switched off; no personal data is logged.
 */

const COPILOT_INVOKE_SCOPE = 'https://api.powerplatform.com/CopilotStudio.Copilots.Invoke';

export class AuthService {
  private static _instance: AuthService;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _msalInstance: any | undefined;
  private _clientId: string = '';
  private _tenantId: string = '';
  private _initialized: boolean = false;

  private constructor() { /* singleton */ }

  public static getInstance(): AuthService {
    if (!AuthService._instance) {
      AuthService._instance = new AuthService();
    }
    return AuthService._instance;
  }

  /**
   * Initialize MSAL once. Values come from the SP list row — never hardcoded.
   * Uses dynamic import so a missing @azure/msal-browser package surfaces a
   * clear [ERR-MSAL] error rather than crashing the extension at load time.
   */
  public async initialize(clientId: string, tenantId: string): Promise<void> {
    if (this._initialized && this._clientId === clientId && this._tenantId === tenantId) {
      return;
    }

    this._clientId = clientId;
    this._tenantId = tenantId;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const msal: any = await import('@azure/msal-browser');
    const PublicClientApplication = msal.PublicClientApplication;

    this._msalInstance = new PublicClientApplication({
      auth: {
        clientId,
        authority: 'https://login.microsoftonline.com/' + tenantId,
        // Pick the redirect URI that matches the AAD registration.
        // Both URIs must be registered as SPA redirect URIs in AAD.
        redirectUri: (function () {
          const origin = window.location.origin;
          const href   = window.location.href.toLowerCase();
          // Workbench — covers /_layouts/15/workbench.aspx and /sites/*/workbench.aspx
          if (href.indexOf('workbench.aspx') !== -1) {
            return origin + '/_layouts/15/workbench.aspx';
          }
          // Production SharePoint pages
          return origin + '/_forms/spfxsinglesignon.aspx';
        }()),
        // Popup flow only — no redirect to return from.
        navigateToLoginRequestUrl: false,
      },
      cache: {
        // Tokens in memory only — gone when the page closes or reloads.
        // Nothing is written to localStorage, sessionStorage or cookies.
        cacheLocation: 'memoryStorage',
        // Short-lived sign-in state for the popup (no tokens) is kept per tab.
        temporaryCacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false,
      },
      system: {
        loggerOptions: {
          // Completely silent — no tokens or auth state written to the console
          loggerCallback: function () { /* intentionally empty */ },
          piiLoggingEnabled: false,
        },
        allowNativeBroker: false,
      },
    });

    await this._msalInstance.initialize();
    this._initialized = true;
  }

  /**
   * Acquire a Copilot Studio invoke token.
   *
   * Flow:
   *   1. Silent from the in-memory cache.
   *   2. Silent SSO with the user's existing Microsoft sign-in (no popup).
   *   3. Popup, only if 1 and 2 fail.
   *
   * NEVER calls acquireTokenRedirect.
   *
   * SECURITY: The returned token is used in-memory only — never stored or logged.
   */
  public async acquireToken(): Promise<string> {
    if (!this._msalInstance) {
      throw new Error('[ERR-MSAL] AuthService not initialized. Call initialize() first.');
    }

    const accounts = this._msalInstance.getAllAccounts();

    const tokenRequest = {
      scopes:  [COPILOT_INVOKE_SCOPE],
      account: accounts && accounts.length > 0 ? accounts[0] : undefined,
    };

    // ── Step 1: Silent acquisition from the in-memory cache ──────────────────
    if (accounts && accounts.length > 0) {
      try {
        const silentResponse = await this._msalInstance.acquireTokenSilent({
          scopes:       [COPILOT_INVOKE_SCOPE],
          account:      accounts[0],
          forceRefresh: false,
        });
        // SECURITY: returned directly to caller — never stored on this class
        return silentResponse.accessToken;
      } catch {
        // Silent failed — fall through to SSO
      }
    }

    // ── Step 2: Silent SSO with the existing Microsoft sign-in ───────────────
    // The user is already signed in to SharePoint. ssoSilent reuses that
    // session in a hidden frame — no popup, no user action.
    try {
      const ssoResponse = await this._msalInstance.ssoSilent({
        scopes: [COPILOT_INVOKE_SCOPE],
      });
      return ssoResponse.accessToken;
    } catch {
      // SSO not possible (e.g. consent needed) — fall through to popup
    }

    // ── Step 3: Popup acquisition (only when steps 1 and 2 fail) ─────────────
    try {
      const popupResponse = await this._msalInstance.acquireTokenPopup(tokenRequest);
      return popupResponse.accessToken;
    } catch (popupErr) {
      const msg = (popupErr instanceof Error) ? popupErr.message : String(popupErr);
      // Common causes:
      //   - Popup blocked by browser → user must allow popups for this site
      //   - User closed the popup before completing sign-in
      //   - AAD consent not yet granted for CopilotStudio.Copilots.Invoke
      throw new Error(
        '[ERR-MSAL] Popup authentication failed: ' + msg +
        '. If popups are blocked, allow them for this SharePoint site and try again.'
      );
    }
  }
}
