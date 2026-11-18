import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ILiveUpdate } from '../models/ILiveUpdate';

/**
 * LiveUpdatesService
 *
 * Reads Summit_Live_Updates from the site at DataSiteUrl.
 * No $select — avoids 400 if any column internal name differs.
 * Returns [] silently on any error so the widget degrades gracefully.
 */
export class LiveUpdatesService {

  public static async getUpdates(
    spHttpClient: SPHttpClient,
    dataSiteUrl:  string,
    listTitle:    string,
    userRole:     string
  ): Promise<ILiveUpdate[]> {

    if (!dataSiteUrl || !listTitle) { return []; }

    // No $select — let SP return all columns to avoid 400 errors
    const url: string =
      dataSiteUrl.replace(/\/$/, '') +
      "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')/items" +
      '?$top=200&$orderby=PriorityRank asc';

    let response: SPHttpClientResponse;
    try {
      response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    } catch { return []; }
    if (!response.ok) { return []; }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let data: { value: any[] };
    try { data = await response.json(); } catch { return []; }
    if (!data || !Array.isArray(data.value)) { return []; }

    const now: number = Date.now();
    const role: string = (userRole || 'All').trim().toLowerCase();
    const results: ILiveUpdate[] = [];

    for (let i: number = 0; i < data.value.length; i++) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const raw: any = data.value[i];

      // IsActive check
      if (raw.IsActive === false) { continue; }

      // Expiry check
      if (raw.ExpirationTime) {
        const exp: number = new Date(raw.ExpirationTime).getTime();
        if (!isNaN(exp) && exp < now) { continue; }
      }

      // AudienceRole filter
      const audience: string = (raw.AudienceRole || '').trim().toLowerCase();
      if (audience !== 'all' && role !== 'all' && audience !== role) { continue; }

      results.push({
        id:              String(raw.ID),
        title:           raw.Title           || '',
        updateType:      raw.UpdateType      || '',
        message:         raw.Message         || '',
        affectedSession: raw.AffectedSession || '',
        audienceRole:    raw.AudienceRole    || 'All',
        priority:        raw.Priority        || 'Low',
        priorityRank:    typeof raw.PriorityRank === 'number' ? raw.PriorityRank : 99,
        expirationTime:  raw.ExpirationTime ? new Date(raw.ExpirationTime) : null,
        isActive:        true,
      });
    }

    return results;
  }
}