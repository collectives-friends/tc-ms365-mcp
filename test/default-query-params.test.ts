import { describe, expect, it } from 'vitest';
import { applyDefaultSelect } from '../src/default-query-params.js';

const MAIL_TIP =
  'When searching emails... IMPORTANT: Always use $select to limit returned fields and reduce ' +
  'response size. Recommended default: ' +
  '$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments. ' +
  'Use bodyPreview instead of body for listings.';

describe('applyDefaultSelect', () => {
  it('applies the upstream-recommended $select when the model omitted it', () => {
    const qp: Record<string, string> = {};
    applyDefaultSelect(MAIL_TIP, qp);
    expect(qp['$select']).toBe(
      'id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments'
    );
  });

  it('never overrides a $select the model supplied explicitly', () => {
    const qp: Record<string, string> = { $select: 'id,subject' };
    applyDefaultSelect(MAIL_TIP, qp);
    expect(qp['$select']).toBe('id,subject');
  });

  it('is a no-op when the llmTip has no "Recommended default" recommendation', () => {
    const qp: Record<string, string> = {};
    applyDefaultSelect('Use $select=displayName for an example, e.g. $select=mail.', qp);
    expect(qp['$select']).toBeUndefined();
  });

  it('is a no-op when there is no llmTip', () => {
    const qp: Record<string, string> = {};
    applyDefaultSelect(undefined, qp);
    expect(qp['$select']).toBeUndefined();
  });
});
