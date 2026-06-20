import { describe, expect, it } from 'vitest';
import {
  applyDefaultSelect,
  resolveDefaultSelect,
  stripSelectFromPath,
  isSelectError,
} from '../src/default-query-params.js';

const MAIL_TIP =
  'When searching emails... IMPORTANT: Always use $select to limit returned fields and reduce ' +
  'response size. Recommended default: ' +
  '$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments. ' +
  'Use bodyPreview instead of body for listings.';

describe('resolveDefaultSelect', () => {
  it('uses the curated default for a known high-traffic list tool', () => {
    expect(resolveDefaultSelect('list-calendar-events', undefined)).toContain(
      'id,subject,start,end'
    );
  });

  it('falls back to the llmTip recommendation when no curated entry exists', () => {
    expect(resolveDefaultSelect('list-mail-messages', MAIL_TIP)).toBe(
      'id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments'
    );
  });

  it('prefers the curated entry over the llmTip when both exist', () => {
    const curated = resolveDefaultSelect('list-calendar-events', MAIL_TIP);
    expect(curated).not.toContain('bodyPreview');
  });

  it('returns undefined when neither curated nor llmTip provides one', () => {
    expect(
      resolveDefaultSelect('list-some-obscure-thing', 'no recommendation here')
    ).toBeUndefined();
  });

  it('every curated default starts with id so follow-up single-item gets work', () => {
    // sample a couple of known tools
    for (const t of ['list-users', 'list-folder-files', 'list-outlook-contacts']) {
      const sel = resolveDefaultSelect(t, undefined);
      expect(sel).toBeDefined();
      expect(sel!.split(',')[0]).toBe('id');
    }
  });
});

describe('applyDefaultSelect', () => {
  it('injects a default and reports true when the model omitted $select', () => {
    const qp: Record<string, string> = {};
    const injected = applyDefaultSelect('list-users', undefined, qp);
    expect(injected).toBe(true);
    expect(qp['$select']).toContain('displayName');
  });

  it('never overrides an explicit $select and reports false', () => {
    const qp: Record<string, string> = { $select: 'id' };
    const injected = applyDefaultSelect('list-users', undefined, qp);
    expect(injected).toBe(false);
    expect(qp['$select']).toBe('id');
  });

  it('reports false when there is no default for the tool', () => {
    const qp: Record<string, string> = {};
    expect(applyDefaultSelect('list-obscure', undefined, qp)).toBe(false);
    expect(qp['$select']).toBeUndefined();
  });
});

describe('stripSelectFromPath (fallback when Graph rejects a default $select)', () => {
  it('removes $select when it is the only query param', () => {
    expect(stripSelectFromPath('/me/events?$select=id,subject')).toBe('/me/events');
  });

  it('removes $select in the middle, keeping the others', () => {
    expect(stripSelectFromPath('/me/events?$select=id,subject&$top=10')).toBe('/me/events?$top=10');
  });

  it('removes $select at the end, keeping the others', () => {
    expect(stripSelectFromPath('/me/events?$top=10&$select=id,subject')).toBe('/me/events?$top=10');
  });

  it('leaves a path without $select unchanged', () => {
    expect(stripSelectFromPath('/me/events?$top=10')).toBe('/me/events?$top=10');
  });
});

describe('isSelectError', () => {
  it('matches Graph errors about an invalid select or property', () => {
    expect(
      isSelectError('Microsoft Graph API error: 400 Bad Request - Parsing OData Select failed')
    ).toBe(true);
    expect(isSelectError("Could not find a property named 'foo' on type 'event'.")).toBe(true);
  });

  it('does not match unrelated errors', () => {
    expect(isSelectError('Microsoft Graph API error: 503 Service Unavailable')).toBe(false);
  });
});
