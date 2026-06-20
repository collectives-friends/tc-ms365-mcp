// Apply upstream-recommended default OData query params on GET requests when the model omitted
// them, so responses are trimmed without changing the tool surface (all tools stay registered).
//
// The default is read from the endpoint's own llmTip ("Recommended default: $select=...") rather
// than a hand-maintained map, so it needs no per-endpoint curation and extends automatically as
// upstream adds recommendations. Anchored on the explicit "Recommended default:" phrase so a
// $select that appears only in an example is never applied. An explicit model value always wins.

const RECOMMENDED_SELECT = /Recommended default:\s*\$select=([A-Za-z0-9_,]+)/i;

export function applyDefaultSelect(
  llmTip: string | undefined,
  queryParams: Record<string, string>
): void {
  if (queryParams['$select'] !== undefined || !llmTip) return;
  const match = llmTip.match(RECOMMENDED_SELECT);
  if (match) {
    queryParams['$select'] = match[1];
  }
}
