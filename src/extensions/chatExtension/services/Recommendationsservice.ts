import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IRecommendation } from '../models/IRecommendation';

/**
 * RecommendationsService
 *
 * Reads NOLA_Recommendations from the site at DataSiteUrl.
 * No $select — avoids 400 if any column internal name differs.
 * MapLink is a SP Hyperlink column — returns {Url, Description} object.
 * Both plain string and Hyperlink object are handled safely.
 */
export class RecommendationsService {

  public static async getRecommendations(
    spHttpClient: SPHttpClient,
    dataSiteUrl:  string,
    listTitle:    string
  ): Promise<IRecommendation[]> {

    if (!dataSiteUrl || !listTitle) { return []; }

    // No $select — let SP return all columns to avoid 400 errors
    const url: string =
      dataSiteUrl.replace(/\/$/, '') +
      "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')/items" +
      '?$filter=IsActive eq 1&$top=200&$orderby=DistanceFromHotel asc';

    let response: SPHttpClientResponse;
    try {
      response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    } catch { return []; }
    if (!response.ok) { return []; }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let data: { value: any[] };
    try { data = await response.json(); } catch { return []; }
    if (!data || !Array.isArray(data.value)) { return []; }

    const results: IRecommendation[] = [];

    for (let i: number = 0; i < data.value.length; i++) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const raw: any = data.value[i];

      // MapLink can be:
      //   - SP Hyperlink column  → { Url: "https://...", Description: "label" }
      //   - Plain Text column    → "https://..."
      //   - Null / undefined     → ""
      const mapLink: string = RecommendationsService._extractUrl(raw.MapLink);

      results.push({
        id:                String(raw.ID),
        title:             raw.Title            || '',
        category:          raw.Category         || '',
        description:       raw.Description      || '',
        distanceFromHotel: typeof raw.DistanceFromHotel === 'number' ? raw.DistanceFromHotel : 0,
        walkingDistance:   raw.WalkingDistance  === true,
        eveningFriendly:   raw.EveningFriendly  === true,
        mapLink:           mapLink,
        website:           raw.Website          || '',
        isActive:          raw.IsActive         !== false,
      });
    }

    return results;
  }

  /**
   * Extracts a safe https:// URL from either:
   *   - A SP Hyperlink field object: { Url: "https://...", Description: "..." }
   *   - A plain string: "https://..."
   *   - Anything else: returns ''
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private static _extractUrl(raw: any): string {
    if (!raw) { return ''; }
    // SP Hyperlink column returns an object with a Url property
    if (typeof raw === 'object' && typeof raw.Url === 'string') {
      return RecommendationsService._safeUrl(raw.Url);
    }
    // Plain text column
    if (typeof raw === 'string') {
      return RecommendationsService._safeUrl(raw);
    }
    return '';
  }

  private static _safeUrl(raw: string): string {
    const t: string = raw.trim();
    return /^https?:\/\//i.test(t) ? t : '';
  }
}