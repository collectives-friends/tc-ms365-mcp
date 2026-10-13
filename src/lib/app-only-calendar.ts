/**
 * App-only (client credentials) access for calendar calls.
 *
 * The service normally acts as the signed-in agent account (delegated). Some
 * deployments want tenant-wide calendar access instead: an application
 * permission (`Calendars.ReadWrite.All`) with a client secret, and no /me
 * mailbox. When `MS365_MCP_CALENDAR_APP_ONLY=1` and a secret is configured,
 * calendar tools are executed with an app-only token and the caller passes
 * `calendarUser` (the organizer's UPN) so `/me/...` paths are rewritten to
 * `/users/{upn}/...`.
 */
import { ConfidentialClientApplication } from '@azure/msal-node';

type Env = Record<string, string | undefined>;

/** True when the deployment routes calendar calls through an app-only token. */
export function isAppOnlyCalendarEnabled(env: Env = process.env): boolean {
  return env.MS365_MCP_CALENDAR_APP_ONLY === '1' && Boolean(env.MS365_MCP_APP_CLIENT_SECRET);
}

const CALENDAR_SEGMENT =
  /(^|\/)(calendars?|calendarView|events|findMeetingTimes|getSchedule)(\/|$)/;

/** True for the Graph paths that carry calendar or event data. */
export function isCalendarToolPath(path: string): boolean {
  return CALENDAR_SEGMENT.test(path);
}

/** Rewrite a `/me/...` calendar path to the organizer's mailbox, or undefined. */
export function rewriteCalendarPath(path: string, userUpn: string): string | undefined {
  if (!path.startsWith('/me/')) return undefined;
  return path.replace(/^\/me\//, `/users/${encodeURIComponent(userUpn)}/`);
}

let cached: { token: string; expiresAt: number } | undefined;
let client: ConfidentialClientApplication | undefined;
let clientKey: string | undefined;

/** Acquire (and cache) an application token for Microsoft Graph. */
export async function getAppOnlyCalendarToken(env: Env = process.env): Promise<string> {
  const now = Date.now();
  if (cached && cached.expiresAt > now + 60_000) return cached.token;
  const clientId = env.MS365_MCP_CLIENT_ID;
  const tenantId = env.MS365_MCP_TENANT_ID;
  const secret = env.MS365_MCP_APP_CLIENT_SECRET;
  if (!clientId || !tenantId || !secret) {
    throw new Error(
      'App-only calendar is not configured (MS365_MCP_CLIENT_ID, MS365_MCP_TENANT_ID and MS365_MCP_APP_CLIENT_SECRET are required).'
    );
  }
  const key = `${tenantId}:${clientId}`;
  if (!client || clientKey !== key) {
    client = new ConfidentialClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        clientSecret: secret,
      },
    });
    clientKey = key;
  }
  const result = await client.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });
  if (!result?.accessToken) {
    throw new Error('The app-only calendar token request returned no access token.');
  }
  cached = {
    token: result.accessToken,
    expiresAt: result.expiresOn?.getTime() ?? now + 5 * 60_000,
  };
  return result.accessToken;
}
