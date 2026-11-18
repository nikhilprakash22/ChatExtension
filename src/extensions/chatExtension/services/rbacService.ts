import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IChatConfig } from '../models/IChatConfig';

/**
 * RbacService
 * Checks whether the current user may use the chatbot.
 *
 * Group matching strategy (all comparisons case-insensitive):
 *   1. Exact:            AllowedRoles="Visitors"  group="Visitors"
 *   2. Ends with role:   AllowedRoles="Visitors"  group="IT EU Visitors"
 *   3. Contains role:    AllowedRoles="Visitors"  group="Visitors (All)"
 *   4. Role in group:    AllowedRoles="IT EU Visitors"  group="Visitors"
 *
 * Fallback when /_api/web/currentuser/groups returns empty
 * (common when user is added via nested Entra/AD group):
 *   - Infers Owners   from IsSiteAdmin flag
 *   - Infers Members  from ManageLists effective permission bit
 *   - Infers Visitors from ViewPages effective permission bit
 */
export class RbacService {

  public static async isUserAllowed(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    config: IChatConfig
  ): Promise<boolean> {

    const rolesRaw: string = config.allowedRoles      != null ? String(config.allowedRoles)      : '';
    const usersRaw: string = config.allowedUsersString != null ? String(config.allowedUsersString) : '';

    const hasRoles = rolesRaw.trim().length > 0;
    const hasUsers = usersRaw.trim().length > 0;

    if (!hasRoles && !hasUsers) { return false; }

    try {
      const currentUser = await RbacService._getCurrentUser(spHttpClient, siteUrl);
      if (!currentUser) { return false; }

      // ── Check explicit user allow-list ──────────────────────────────────────
      if (hasUsers) {
        const allowedUsers = usersRaw
          .split(',')
          .map(function(u) { return u.trim().toLowerCase(); })
          .filter(Boolean);

        if (allowedUsers.indexOf(currentUser.loginName.toLowerCase()) !== -1) {
          return true;
        }
        if (currentUser.email &&
            allowedUsers.indexOf(currentUser.email.toLowerCase()) !== -1) {
          return true;
        }
      }

      // ── Check SP group membership ───────────────────────────────────────────
      if (hasRoles) {
        const allowedRoles = rolesRaw
          .split(',')
          .map(function(r) { return r.trim().toLowerCase(); })
          .filter(Boolean);

        const userGroups = await RbacService._getUserGroups(spHttpClient, siteUrl);

        if (RbacService._matchesRole(userGroups, allowedRoles)) {
          return true;
        }
      }

      return false;
    } catch {
      return false;
    }
  }

  // ── Role matching ───────────────────────────────────────────────────────────

  private static _matchesRole(userGroups: string[], allowedRoles: string[]): boolean {
    for (let gi = 0; gi < userGroups.length; gi++) {
      const groupLower = userGroups[gi].toLowerCase();
      for (let ri = 0; ri < allowedRoles.length; ri++) {
        const role = allowedRoles[ri];

        // 1. Exact match
        if (groupLower === role) { return true; }

        // 2. Group ends with " role" — handles "IT EU Visitors", "Test_Test Group Visitors"
        const suffix = ' ' + role;
        const hasSuffix =
          groupLower.length >= suffix.length &&
          groupLower.indexOf(suffix, groupLower.length - suffix.length) !== -1;
        if (hasSuffix) { return true; }

        // 3. Group contains role anywhere — handles "Visitors (All Users)"
        if (groupLower.indexOf(role) !== -1) { return true; }

        // 4. Role contains group — handles full group name in AllowedRoles config
        if (role.indexOf(groupLower) !== -1) { return true; }
      }
    }
    return false;
  }

  // ── Get user groups (with effective-permissions fallback) ───────────────────

  private static async _getUserGroups(
    spHttpClient: SPHttpClient,
    siteUrl: string
  ): Promise<string[]> {
    const base = siteUrl.replace(/\/$/, '');
    const results: string[] = [];

    // Primary: direct SP group memberships via currentuser/groups
    // NOTE: this only returns groups the user is a DIRECT member of.
    // If added via a nested Entra/AD security group it returns empty.
    try {
      const resp: SPHttpClientResponse = await spHttpClient.get(
        base + '/_api/web/currentuser/groups?$select=Title',
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const data: { value?: Array<{ Title?: string }> } = await resp.json();
        if (data.value) {
          for (let i = 0; i < data.value.length; i++) {
            const t = data.value[i].Title || '';
            if (t && results.indexOf(t) === -1) { results.push(t); }
          }
        }
      }
    } catch { /* ignore */ }

    // Fallback: if direct group call returned nothing, infer role from
    // effective permissions — handles nested Entra/AD group scenario
    if (results.length === 0) {
      try {
        // Check IsSiteAdmin → infer Owners
        const userResp: SPHttpClientResponse = await spHttpClient.get(
          base + '/_api/web/currentuser?$select=IsSiteAdmin',
          SPHttpClient.configurations.v1
        );
        if (userResp.ok) {
          const userData: { IsSiteAdmin?: boolean } = await userResp.json();
          if (userData.IsSiteAdmin) {
            results.push('Owners');
          }
        }
      } catch { /* ignore */ }

      try {
        // Check effective permissions to infer Members / Visitors
        const permResp: SPHttpClientResponse = await spHttpClient.get(
          base + '/_api/web/effectivebasepermissions',
          SPHttpClient.configurations.v1
        );
        if (permResp.ok) {
          const permData: { Low?: string; High?: string } = await permResp.json();
          const low = parseInt(permData.Low || '0', 10);

          // ManageLists bit (0x800) — Members and above have this
          const canEdit = (low & 0x800) !== 0;

          // ViewPages bit (0x20000) — Visitors and above have this
          const canView = (low & 0x20000) !== 0;

          if (canEdit) { results.push('Members'); }
          // Always push Visitors if user can view — covers Visitors role
          if (canView) { results.push('Visitors'); }
        }
      } catch { /* ignore */ }
    }

    return results;
  }

  // ── Current user ────────────────────────────────────────────────────────────

  private static async _getCurrentUser(
    spHttpClient: SPHttpClient,
    siteUrl: string
  ): Promise<{ loginName: string; email: string } | null> {
    const url = siteUrl.replace(/\/$/, '') + '/_api/web/currentuser?$select=LoginName,Email';
    let resp: SPHttpClientResponse;
    try {
      resp = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    } catch {
      return null;
    }
    if (!resp.ok) { return null; }
    const data: { LoginName?: string; Email?: string } = await resp.json();
    return {
      loginName: data.LoginName || '',
      email:     data.Email     || '',
    };
  }
}