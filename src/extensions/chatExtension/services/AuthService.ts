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
 *     1. acquireTokenSilent  — always tried first (no UI, instant)
 *     2. acquireTokenPopup   — fallback when silent fails (small popup window,
 *                              auth completes there, SharePoint page stays intact)
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
 * SECURITY: Tokens are never logged, stored in localStorage, or exposed to the DOM.
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
        // Dynamically pick the redirect URI that matches the AAD registration.
        // Both URIs must be registered as SPA redirect URIs in AAD.
        // acquireTokenPopup uses this URI internally for the popup window.
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
        // Do NOT navigate back to the original page after login —
        // we use popup so there is no redirect to return from.
        navigateToLoginRequestUrl: false,
      },
      cache: {
        // sessionStorage — scoped to the browser tab, never persisted across sessions
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false,
      },
      system: {
        loggerOptions: {
          // Completely silent — no tokens or auth state written to the console
          loggerCallback: function () { /* intentionally empty */ },
          piiLoggingEnabled: false,
        },
        // Allow the popup to open. Some SP tenant policies block popups —
        // if that happens the catch below surfaces a clear [ERR-MSAL] message.
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
   *   1. Try silent acquisition (no user interaction, no page navigation).
   *   2. If silent fails → open a small popup window for interactive consent.
   *      The popup closes itself after login; the SharePoint page is untouched.
   *
   * NEVER calls acquireTokenRedirect — that would navigate away from SharePoint
   * and leave a blank page on return.
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

    // ── Step 1: Silent acquisition ────────────────────────────────────────────
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
        // Silent failed — fall through to popup
      }
    }

    // ── Step 2: Popup acquisition ─────────────────────────────────────────────
    // Opens a small Microsoft login popup. The current SharePoint page stays
    // open and fully functional. The popup closes automatically after login.
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