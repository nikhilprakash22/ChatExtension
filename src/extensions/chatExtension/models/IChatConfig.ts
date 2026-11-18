export interface IChatConfig {
  siteId:               string;
  agentDisplayName:     string;
  environmentId:        string;
  schemaName:           string;
  appClientId:          string;
  tenantId:             string;
  directConnectUrl:     string;
  allowedRoles:         string;
  allowedUsersString:   string;
  showOnceText:         string;
  enabled:              boolean;
  suggestedPromptsJson: string;
  msFormsUrl:           string;
  // ── New optional fields ───────────────────────────────────────────────────
  // Present only when the agent row has these columns filled in.
  // When blank/missing all new features are simply skipped — existing
  // agents are completely unaffected.
  enableNotifications?:      boolean; // gates Updates tab + banner
  enableConnectorConsent?:          boolean; // gates Nearby tab
  dataSiteUrl?:              string;  // site hosting the two data lists
  liveUpdatesListTitle?:     string;  // e.g. "Summit_Live_Updates"
  recommendationsListTitle?: string;  // e.g. "NOLA_Recommendations"
  geocodeContext?:           string;  // venue label in Nearby tab
  agentIconURL?:            string;  // optional URL to agent icon (e.g. PNG)
}