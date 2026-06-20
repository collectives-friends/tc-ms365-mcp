// Apply default OData query params on GET requests when the model omitted them, so responses are
// trimmed without changing the tool surface (all tools stay registered, read and write).
//
// Two sources of defaults, in priority order:
//  1. A curated $select for a handful of high-traffic list endpoints whose resource fields are
//     stable and well known (below). id is always first so a follow-up single-item get works.
//  2. The endpoint's own llmTip ("Recommended default: $select=..."), so mail and any future
//     upstream recommendation is picked up with no per-endpoint maintenance.
//
// Safety: a wrong field name only makes Microsoft Graph return a 400. executeGraphTool detects
// that (isSelectError) and retries once with the default stripped (stripSelectFromPath), so a
// default can never break a tool. Worst case is the full, untrimmed response. An explicit model
// $select always wins. This module is self-contained to stay merge-friendly against upstream sync.

const RECOMMENDED_SELECT = /Recommended default:\s*\$select=([A-Za-z0-9_,]+)/i;

// Curated conservative defaults. Core, long-stable Graph properties only.
const EVENT_SELECT = 'id,subject,start,end,location,organizer,isAllDay,showAs';
const CONTACT_SELECT = 'id,displayName,emailAddresses,mobilePhone,companyName,jobTitle';
const DRIVE_SELECT = 'id,name,driveType,webUrl';

const DEFAULT_SELECT_BY_TOOL: Record<string, string> = {
  'list-calendar-events': EVENT_SELECT,
  'list-specific-calendar-events': EVENT_SELECT,
  'list-shared-calendar-events': EVENT_SELECT,
  'list-calendar-event-instances': EVENT_SELECT,
  'list-group-events': EVENT_SELECT,
  'list-calendars': 'id,name,color,isDefaultCalendar,canEdit',
  'list-folder-files': 'id,name,size,lastModifiedDateTime,webUrl,folder,file',
  'list-drives': DRIVE_SELECT,
  'list-sharepoint-site-drives': DRIVE_SELECT,
  'list-outlook-contacts': CONTACT_SELECT,
  'list-contact-folder-contacts': CONTACT_SELECT,
  'list-users': 'id,displayName,mail,userPrincipalName,jobTitle',
  'list-todo-tasks': 'id,title,status,importance,dueDateTime',
};

export function resolveDefaultSelect(
  toolName: string | undefined,
  llmTip: string | undefined
): string | undefined {
  if (toolName && DEFAULT_SELECT_BY_TOOL[toolName]) {
    return DEFAULT_SELECT_BY_TOOL[toolName];
  }
  if (llmTip) {
    const match = llmTip.match(RECOMMENDED_SELECT);
    if (match) return match[1];
  }
  return undefined;
}

/** Inject a default $select if the model omitted one. Returns true if one was injected. */
export function applyDefaultSelect(
  toolName: string | undefined,
  llmTip: string | undefined,
  queryParams: Record<string, string>
): boolean {
  if (queryParams['$select'] !== undefined) return false;
  const select = resolveDefaultSelect(toolName, llmTip);
  if (select) {
    queryParams['$select'] = select;
    return true;
  }
  return false;
}

/** Remove a $select clause from a built path's query string (for the fallback retry). */
export function stripSelectFromPath(path: string): string {
  return path
    .replace(/([?&])\$select=[^&]*&/i, '$1')
    .replace(/[?&]\$select=[^&]*$/i, '')
    .replace(/\?$/, '');
}

/** Whether a Graph error text looks like it was caused by an invalid $select. */
export function isSelectError(text: string): boolean {
  return /select|property/i.test(text);
}
