import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IChatConfig } from '../models/IChatConfig';

interface IListItem {
  Title?:                    string;
  SiteId?:                   string;
  AgentDisplayName?:         string;
  EnvironmentId?:            string;
  SchemaName?:               string;
  AppClientId?:              string;
  TenantId?:                 string;
  DirectConnectUrl?:         string;
  AllowedRoles?:             string;
  AllowedUsersString?:       string;
  ShowOnceText?:             string;
  Enabled?:                  boolean;
  SuggestedPromptsJson?:     string;
  MSFormsUrl?:               string;
  EnableNotifications?:      boolean;
  EnableConnectorConsent?:   boolean;
  DataSiteUrl?:              string;
  LiveUpdatesListTitle?:     string;
  RecommendationsListTitle?: string;
  GeocodeContext?:           string;
  agentIconUrl?:             string; // ── NEW: custom bot avatar column
}

export class ConfigService {
  private static _cachedConfig: IChatConfig | null = null;
  private static _cacheKey: string = '';
  private static _cacheExpiry: number = 0;
  private static readonly CACHE_TTL_MS: number = 5 * 60 * 1000;

  public static async getConfig(
    spHttpClient: SPHttpClient,
    hubSiteUrl: string,
    listTitle: string,
    currentSiteId: string,
    defaultTitleMatch: string
  ): Promise<IChatConfig | null> {

    const cacheKey = hubSiteUrl + '|' + currentSiteId;
    const now = Date.now();
    if (
      ConfigService._cachedConfig &&
      ConfigService._cacheKey === cacheKey &&
      now < ConfigService._cacheExpiry
    ) {
      return ConfigService._cachedConfig;
    }

    const baseUrl = hubSiteUrl.replace(/\/$/, '');
    // No $select — SP returns all columns safely.
    // Prevents 400 Bad Request if any column name in $select
    // does not exactly match the SP internal field name.
    const url =
      baseUrl +
      "/_api/web/lists/getbytitle('" + listTitle + "')/items" +
      '?$top=100';

    const allItems = await ConfigService._fetchAllItems(spHttpClient, url);
    if (!allItems) { return null; }

    const normalisedId = currentSiteId.replace(/[{}]/g, '').toLowerCase();
    let matched: IListItem | null = null;

    for (let i = 0; i < allItems.length; i++) {
      const rowId = (allItems[i].SiteId || '').replace(/[{}]/g, '').toLowerCase();
      if (rowId && rowId === normalisedId) { matched = allItems[i]; break; }
    }

    if (!matched) {
      for (let i = 0; i < allItems.length; i++) {
        if ((allItems[i].Title || '') === defaultTitleMatch) {
          matched = allItems[i]; break;
        }
      }
    }

    if (!matched) { return null; }

    const config = ConfigService._mapItem(matched);
    ConfigService._cachedConfig = config;
    ConfigService._cacheKey     = cacheKey;
    ConfigService._cacheExpiry  = now + ConfigService.CACHE_TTL_MS;
    return config;
  }

  public static clearCache(): void {
    ConfigService._cachedConfig = null;
    ConfigService._cacheKey     = '';
    ConfigService._cacheExpiry  = 0;
  }

  private static async _fetchAllItems(
    spHttpClient: SPHttpClient,
    url: string
  ): Promise<IListItem[] | null> {
    let response: SPHttpClientResponse;
    try {
      response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    } catch { return null; }
    if (!response.ok) { return null; }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let data: { value: any[] };
    try { data = await response.json(); } catch { return null; }
    return data.value || null;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private static _str(val: any, fallback: string): string {
    if (val === null || val === undefined) { return fallback; }
    const s = String(val).trim();
    return s.length > 0 ? s : fallback;
  }

  private static _mapItem(item: IListItem): IChatConfig {
    const s = ConfigService._str.bind(ConfigService);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw: any = item;
    return {
      // ── Original fields — identical to original ────────────────────────────
      siteId:               s(item.SiteId,               ''),
      agentDisplayName:     s(item.AgentDisplayName,     'Assistant'),
      environmentId:        s(item.EnvironmentId,        ''),
      schemaName:           s(item.SchemaName,           ''),
      appClientId:          s(item.AppClientId,          ''),
      tenantId:             s(item.TenantId,             ''),
      directConnectUrl:     s(item.DirectConnectUrl,     ''),
      allowedRoles:         s(item.AllowedRoles,         ''),
      allowedUsersString:   s(item.AllowedUsersString,   ''),
      showOnceText:         s(item.ShowOnceText,         ''),
      enabled:              item.Enabled !== false,
      suggestedPromptsJson: s(item.SuggestedPromptsJson, '[]'),
      msFormsUrl:           s(item.MSFormsUrl,           ''),
      // ── New optional fields — exact SP internal column names ───────────────
      enableNotifications:      raw.EnableNotifications    === true,
      enableConnectorConsent:          raw.EnableConnectorConsent === true,
      dataSiteUrl:              s(raw.DataSiteUrl,              ''),
      liveUpdatesListTitle:     s(raw.LiveUpdatesListTitle,     ''),
      recommendationsListTitle: s(raw.RecommendationsListTitle, ''),
      geocodeContext:           s(raw.GeocodeContext,           ''),
      agentIconURL:             s(raw.agentIconUrl,             ''), // ── NEW
    };
  }
}