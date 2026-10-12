/**
 * Semantic search index over configured SharePoint drives.
 *
 * The service keeps one SQLite index (see search-index-store.ts) in sync with
 * the drives named in MS365_MCP_INDEX_DRIVES through the Graph delta feed:
 * new and changed files are downloaded, their text extracted with the same
 * parsers extract-document-text uses (Word, PDF, Excel, PowerPoint, plain
 * text), chunked, embedded with OpenAI and written to the store; deleted
 * items are removed. A tool call and the scheduled worker share one sync lock,
 * so a run is never doubled up.
 *
 * Everything here is derived data: the store can be deleted and rebuilt from
 * Graph at any time, and nothing on the Microsoft side is ever written. Log
 * lines carry counts and status codes only -- never file names or content.
 */

import { statSync } from 'fs';

import logger from '../logger.js';
import { tokenize } from './bm25.js';
import { detectDocumentKind, extractDocumentText, type DocumentKind } from './document-text.js';
import {
  DEFAULT_INDEX_DB_PATH,
  loadSearchIndexConfig,
  validateSearchIndexConfig,
  type Env,
  type SearchIndexConfig,
} from './search-index-config.js';
import { EmbedClient } from './search-index-embed.js';
import {
  openSearchIndexStore,
  type ItemRecord,
  type SearchChunkRow,
  type SearchIndexStore,
} from './search-index-store.js';
import { chunkText } from './text-chunk.js';

const EMBED_BATCH_SIZE = 64;
const MAX_CHUNKS_PER_DOC = 600;
const MAX_QUERY_CHARS = 2_000;
const RETRY_ERROR_ITEMS_PER_RUN = 10;
const INITIAL_SYNC_DELAY_MS = 10_000;
const COLD_RUN_MULTIPLIER = 10;
const COLD_RUN_MAX_MINUTES = 15;
const WARM_RUN_MAX_MINUTES = 5;
const SNAPSHOT_CHARS = 400;
const RRF_K = 60;

// ---------------------------------------------------------------------------
// Injected dependencies (structural types: the server passes its GraphClient
// and AuthManager, tests pass small fakes).
// ---------------------------------------------------------------------------

export interface GraphAccess {
  makeRequest(endpoint: string, options?: { accessToken?: string }): Promise<unknown>;
  downloadStream(
    endpoint: string,
    options?: { accessToken?: string }
  ): Promise<{
    body: NonNullable<Response['body']>;
    contentType: string;
    contentLength: number | null;
    contentDisposition: string | null;
  }>;
}

export interface TokenSource {
  isOAuthModeEnabled(): boolean;
  getTokenForAccount(identifier?: string): Promise<string>;
}

export interface SearchIndexDeps {
  graph: GraphAccess;
  auth?: TokenSource;
  /** Test seam: reuse an embed client instead of building one from the environment. */
  embed?: EmbedClient;
}

// ---------------------------------------------------------------------------
// Runtime
// ---------------------------------------------------------------------------

interface SearchIndexRuntime {
  config: SearchIndexConfig;
  store: SearchIndexStore;
  embed: EmbedClient;
  syncing: boolean;
}

let runtime: SearchIndexRuntime | null = null;
let blockReason: string | undefined;
let workerTimer: ReturnType<typeof setInterval> | null = null;
let initialTimer: ReturnType<typeof setTimeout> | null = null;

/**
 * Create the index store and start the scheduled sync worker. Safe to call
 * once per process: when the feature is disabled or misconfigured the reason
 * is remembered so the tools can explain themselves instead of failing
 * obscurely. Errors thrown here are the caller's to log.
 */
export async function initSearchIndex(
  deps: SearchIndexDeps,
  env: Env = process.env
): Promise<void> {
  if (runtime) return;
  const config = loadSearchIndexConfig(env);
  if (!config.enabled) {
    blockReason = 'The search index is disabled (MS365_MCP_INDEX_ENABLED is not set).';
    return;
  }
  const invalid = validateSearchIndexConfig(config);
  if (invalid) {
    blockReason = `The search index is not configured: ${invalid}`;
    logger.warn(`[SEARCH INDEX] not started: ${invalid}`);
    return;
  }

  const store = await openSearchIndexStore({
    dbPath: config.dbPath,
    dims: config.embedDims,
    embedModel: config.embedModel,
  });
  const embed =
    deps.embed ??
    new EmbedClient({
      apiKey: config.openaiApiKey ?? '',
      model: config.embedModel,
      dims: config.embedDims,
      baseUrl: config.openaiBaseUrl,
    });
  runtime = { config, store, embed, syncing: false };
  blockReason = undefined;
  logger.info(
    `[SEARCH INDEX] ready: ${store.vectorEnabled ? 'sqlite-vec' : 'fts5-only'} backend, ` +
      `${config.driveIds.length} drive(s), sync every ${config.syncMinutes} min`
  );

  if (config.syncMinutes > 0) {
    initialTimer = setTimeout(() => {
      void runScheduledSync(deps);
    }, INITIAL_SYNC_DELAY_MS);
    initialTimer.unref();
    workerTimer = setInterval(() => {
      void runScheduledSync(deps);
    }, config.syncMinutes * 60_000);
    workerTimer.unref();
  }
}

/** Stop timers and close the store (server shutdown and tests). */
export function stopSearchIndex(): void {
  if (workerTimer) clearInterval(workerTimer);
  if (initialTimer) clearTimeout(initialTimer);
  workerTimer = null;
  initialTimer = null;
  if (runtime) {
    runtime.store.close();
    runtime = null;
  }
  blockReason = undefined;
}

function requireRuntime(): SearchIndexRuntime {
  if (!runtime) {
    throw new Error(blockReason ?? 'The search index is not initialized.');
  }
  return runtime;
}

// ---------------------------------------------------------------------------
// Sync engine
// ---------------------------------------------------------------------------

export interface SurveyBucket {
  files: number;
  bytes: number;
  oversized: number;
}

export interface DriveSyncSummary {
  driveId: string;
  status: 'ok' | 'error' | 'deferred';
  seen: number;
  folders: number;
  indexed: number;
  unchanged: number;
  skipped: number;
  deleted: number;
  errors: number;
  moreRemaining: boolean;
  elapsedMs: number;
  error?: string;
  survey?: {
    byKind: Record<string, SurveyBucket>;
    unsupported: { files: number; bytes: number };
  };
}

interface SyncRunOptions {
  driveIds: string[];
  full: boolean;
  dryRun: boolean;
  maxMinutes: number;
  maxItems: number;
}

interface SyncRunResult {
  drives: DriveSyncSummary[];
  elapsedMs: number;
}

async function runScheduledSync(deps: SearchIndexDeps): Promise<void> {
  const active = runtime;
  if (!active || active.syncing) return;
  const cold = active.config.driveIds.some(
    (driveId) => !active.store.getDriveState(driveId)?.deltaLink
  );
  active.syncing = true;
  try {
    await runSync(deps, {
      driveIds: active.config.driveIds,
      full: false,
      dryRun: false,
      maxItems: cold
        ? active.config.maxItemsPerRun * COLD_RUN_MULTIPLIER
        : active.config.maxItemsPerRun,
      maxMinutes: cold ? COLD_RUN_MAX_MINUTES : WARM_RUN_MAX_MINUTES,
    });
  } catch (error) {
    logger.warn(`[SEARCH INDEX] scheduled sync failed: ${shortError(error)}`);
  } finally {
    active.syncing = false;
  }
}

async function runSync(deps: SearchIndexDeps, options: SyncRunOptions): Promise<SyncRunResult> {
  const active = requireRuntime();
  const started = Date.now();
  const deadline = started + options.maxMinutes * 60_000;
  const summaries: DriveSyncSummary[] = [];
  for (const driveId of options.driveIds) {
    if (Date.now() > deadline) {
      summaries.push(deferredSummary(driveId));
      continue;
    }
    summaries.push(await syncDrive(deps, active, driveId, { ...options, deadline }));
  }
  return { drives: summaries, elapsedMs: Date.now() - started };
}

function emptySummary(driveId: string): DriveSyncSummary {
  return {
    driveId,
    status: 'ok',
    seen: 0,
    folders: 0,
    indexed: 0,
    unchanged: 0,
    skipped: 0,
    deleted: 0,
    errors: 0,
    moreRemaining: false,
    elapsedMs: 0,
  };
}

function deferredSummary(driveId: string): DriveSyncSummary {
  return { ...emptySummary(driveId), status: 'deferred', moreRemaining: true };
}

// Only genuine work (new or changed entries) counts against the run's item
// budget. Re-walked unchanged entries are free, so a run that stops at the
// budget still makes real progress on the next run, even without a stored
// cursor; a feed that fits inside the budget completes and stores its token.
function processedWork(summary: DriveSyncSummary): number {
  return summary.folders + summary.indexed + summary.skipped + summary.deleted + summary.errors;
}

/**
 * The item budget bounds work per run (downloads, extraction, embeddings), not
 * items walked: entries that are already indexed, already skipped or no longer
 * present cost only a metadata comparison, so a run that stops at the budget
 * advances through them for free next time instead of re-spending its whole
 * budget on the same prefix. A run that walks the feed to the end within
 * budget persists the delta token.
 */
async function syncDrive(
  deps: SearchIndexDeps,
  active: SearchIndexRuntime,
  driveId: string,
  options: SyncRunOptions & { deadline: number }
): Promise<DriveSyncSummary> {
  const started = Date.now();
  const summary = emptySummary(driveId);
  const survey: NonNullable<DriveSyncSummary['survey']> = {
    byKind: {},
    unsupported: { files: 0, bytes: 0 },
  };
  try {
    const accessToken = await resolveAccessToken(deps);
    const state = active.store.getDriveState(driveId);
    if (options.full && !options.dryRun) {
      active.store.resetDrive(driveId);
    }
    if (!options.dryRun) {
      await ensureDriveRoot(deps, active, driveId, accessToken);
      await retryErroredItems(deps, active, driveId, accessToken, summary, options);
    }

    let url: string | null =
      state?.deltaLink && !options.full
        ? toGraphDeltaPath(state.deltaLink)
        : `/drives/${driveId}/root/delta`;
    let more = false;
    while (url) {
      if (processedWork(summary) >= options.maxItems || Date.now() > options.deadline) {
        more = true;
        break;
      }
      const page = asRecord(await deps.graph.makeRequest(url, { accessToken }));
      const items = Array.isArray(page?.value) ? page.value.filter(isRecord) : [];
      for (const item of items) {
        if (processedWork(summary) >= options.maxItems) {
          more = true;
          break;
        }
        summary.seen += 1;
        if (options.dryRun) {
          surveyItem(active, item, summary, survey);
          continue;
        }
        await applyDeltaItem(deps, active, driveId, item, accessToken, summary);
      }
      if (more) break;
      const nextLink = asString(page?.['@odata.nextLink']);
      const deltaLink = asString(page?.['@odata.deltaLink']);
      if (nextLink) {
        url = toGraphDeltaPath(nextLink);
        continue;
      }
      if (deltaLink && !options.dryRun) {
        active.store.saveDriveDelta(driveId, deltaLink);
      }
      url = null;
    }
    if (more) {
      summary.moreRemaining = true;
      if (!options.dryRun) active.store.recordDriveRun(driveId, 'partial');
    }
  } catch (error) {
    summary.status = 'error';
    summary.error = shortError(error);
    if (!options.dryRun) {
      try {
        active.store.recordDriveRun(driveId, 'error');
      } catch {
        // The original error is the one to report.
      }
    }
    logger.warn(`[SEARCH INDEX] drive sync failed: ${summary.error}`);
  }
  summary.elapsedMs = Date.now() - started;
  if (options.dryRun) summary.survey = survey;
  return summary;
}

async function resolveAccessToken(deps: SearchIndexDeps): Promise<string | undefined> {
  if (deps.auth && !deps.auth.isOAuthModeEnabled()) {
    return deps.auth.getTokenForAccount(undefined);
  }
  // OAuth/request-context mode: let the Graph client fall back to the request
  // token it already carries.
  return undefined;
}

/** Store the drive root so relative paths can be resolved for its children. */
async function ensureDriveRoot(
  deps: SearchIndexDeps,
  active: SearchIndexRuntime,
  driveId: string,
  accessToken: string | undefined
): Promise<void> {
  if (active.store.getDriveState(driveId)?.deltaLink) return;
  try {
    const root = asRecord(await deps.graph.makeRequest(`/drives/${driveId}/root`, { accessToken }));
    const itemId = asString(root?.id);
    if (itemId) {
      active.store.upsertItem({
        itemId,
        driveId,
        name: asString(root?.name),
        state: 'folder',
      });
    }
  } catch (error) {
    // Path building degrades to the parts the delta feed knows; not fatal.
    logger.warn(`[SEARCH INDEX] drive root lookup failed: ${shortError(error)}`);
  }
}

async function retryErroredItems(
  deps: SearchIndexDeps,
  active: SearchIndexRuntime,
  driveId: string,
  accessToken: string | undefined,
  summary: DriveSyncSummary,
  options: SyncRunOptions & { deadline: number }
): Promise<void> {
  const errored = active.store.itemsByState(driveId, 'error', RETRY_ERROR_ITEMS_PER_RUN);
  for (const item of errored) {
    if (processedWork(summary) >= options.maxItems || Date.now() > options.deadline) {
      summary.moreRemaining = true;
      return;
    }
    const kind = detectDocumentKind(item.mime, item.name);
    if (!kind) {
      active.store.upsertItem({ ...item, state: 'skipped', skipReason: 'unsupported' });
      continue;
    }
    summary.seen += 1;
    const outcome = await indexItemContent(deps, active, item, kind, accessToken);
    if (outcome === 'indexed') summary.indexed += 1;
    else if (outcome === 'skipped') summary.skipped += 1;
    else summary.errors += 1;
  }
}

async function applyDeltaItem(
  deps: SearchIndexDeps,
  active: SearchIndexRuntime,
  driveId: string,
  item: Record<string, unknown>,
  accessToken: string | undefined,
  summary: DriveSyncSummary
): Promise<void> {
  const itemId = asString(item.id);
  if (!itemId) return;
  if (item.deleted !== undefined) {
    if (active.store.getItem(itemId)) {
      active.store.deleteItemWithChunks(itemId);
      summary.deleted += 1;
    } else {
      // Deleted before we ever indexed it: nothing to remove, so no work.
      summary.unchanged += 1;
    }
    return;
  }

  const metadata = deltaItemMetadata(item);
  if (item.folder !== undefined) {
    const existing = active.store.getItem(itemId);
    if (existing && existing.state === 'folder' && isUnchanged(existing, metadata)) {
      summary.unchanged += 1;
      return;
    }
    active.store.upsertItem({
      itemId,
      driveId,
      parentId: metadata.parentId,
      name: metadata.name,
      state: 'folder',
    });
    summary.folders += 1;
    return;
  }
  if (item.file === undefined) {
    // Packages (OneNote etc.) and other non-file entries are not documents.
    if (isSkippedUnchanged(active.store.getItem(itemId), metadata, 'package')) {
      summary.unchanged += 1;
      return;
    }
    active.store.upsertItem({
      itemId,
      driveId,
      parentId: metadata.parentId,
      name: metadata.name,
      webUrl: metadata.webUrl,
      size: metadata.size,
      lastModified: metadata.lastModified,
      mime: metadata.mime,
      quickxor: metadata.quickxor,
      state: 'skipped',
      skipReason: 'package',
    });
    summary.skipped += 1;
    return;
  }

  if (!metadata.kind) {
    if (isSkippedUnchanged(active.store.getItem(itemId), metadata, 'unsupported')) {
      summary.unchanged += 1;
      return;
    }
    active.store.upsertItem({
      itemId,
      driveId,
      parentId: metadata.parentId,
      name: metadata.name,
      webUrl: metadata.webUrl,
      size: metadata.size,
      lastModified: metadata.lastModified,
      mime: metadata.mime,
      quickxor: metadata.quickxor,
      state: 'skipped',
      skipReason: 'unsupported',
    });
    summary.skipped += 1;
    return;
  }
  if (metadata.size !== undefined && metadata.size > active.config.maxFileBytes) {
    if (isSkippedUnchanged(active.store.getItem(itemId), metadata, 'too_large')) {
      summary.unchanged += 1;
      return;
    }
    active.store.upsertItem({
      itemId,
      driveId,
      parentId: metadata.parentId,
      name: metadata.name,
      webUrl: metadata.webUrl,
      size: metadata.size,
      lastModified: metadata.lastModified,
      mime: metadata.mime,
      quickxor: metadata.quickxor,
      state: 'skipped',
      skipReason: 'too_large',
    });
    summary.skipped += 1;
    return;
  }

  const existing = active.store.getItem(itemId);
  if (existing && existing.state === 'indexed' && isUnchanged(existing, metadata)) {
    summary.unchanged += 1;
    return;
  }

  const input: IndexableItem = {
    itemId,
    driveId,
    parentId: metadata.parentId,
    name: metadata.name,
    path: resolveItemPath(active.store, metadata.parentId, metadata.name),
    webUrl: metadata.webUrl,
    size: metadata.size,
    lastModified: metadata.lastModified,
    mime: metadata.mime,
    quickxor: metadata.quickxor,
  };
  const outcome = await indexItemContent(deps, active, input, metadata.kind, accessToken);
  if (outcome === 'indexed') summary.indexed += 1;
  else if (outcome === 'skipped') summary.skipped += 1;
  else summary.errors += 1;
}

interface DeltaItemMetadata {
  name?: string;
  parentId?: string;
  webUrl?: string;
  size?: number;
  lastModified?: string;
  mime?: string;
  quickxor?: string;
  kind?: DocumentKind;
}

function deltaItemMetadata(item: Record<string, unknown>): DeltaItemMetadata {
  const file = asRecord(item.file);
  const hashes = asRecord(file?.hashes);
  const parent = asRecord(item.parentReference);
  const name = asString(item.name);
  const mime = asString(file?.mimeType);
  return {
    name,
    parentId: asString(parent?.id),
    webUrl: asString(item.webUrl),
    size: asNumber(item.size),
    lastModified: asString(item.lastModifiedDateTime),
    mime,
    quickxor: asString(hashes?.quickXorHash),
    kind: detectDocumentKind(mime, name),
  };
}

interface IndexableItem {
  itemId: string;
  driveId: string;
  parentId?: string;
  name?: string;
  path?: string;
  webUrl?: string;
  size?: number;
  lastModified?: string;
  mime?: string;
  quickxor?: string;
}

function isUnchanged(existing: ItemRecord, fresh: DeltaItemMetadata): boolean {
  if (existing.name !== fresh.name || existing.parentId !== fresh.parentId) return false;
  if (fresh.quickxor && existing.quickxor) return fresh.quickxor === existing.quickxor;
  return existing.size === fresh.size && existing.lastModified === fresh.lastModified;
}

// A skipped entry (unsupported type, too large, package) needs no work again
// until its metadata changes or the skip reason no longer applies.
function isSkippedUnchanged(
  existing: ItemRecord | undefined,
  fresh: DeltaItemMetadata,
  reason: string
): boolean {
  return Boolean(
    existing &&
    existing.state === 'skipped' &&
    existing.skipReason === reason &&
    isUnchanged(existing, fresh)
  );
}

function resolveItemPath(
  store: SearchIndexStore,
  parentId: string | undefined,
  name: string | undefined
): string | undefined {
  const parts: string[] = [];
  let current = parentId;
  for (let depth = 0; current && depth < 64; depth++) {
    const parent = store.getItem(current);
    if (!parent) break;
    if (parent.name) parts.unshift(parent.name);
    current = parent.parentId;
  }
  if (name) parts.push(name);
  return parts.length > 0 ? `/${parts.join('/')}` : undefined;
}

type IndexOutcome = 'indexed' | 'skipped' | 'error';

async function indexItemContent(
  deps: SearchIndexDeps,
  active: SearchIndexRuntime,
  item: ItemRecord | IndexableItem,
  kind: DocumentKind,
  accessToken: string | undefined
): Promise<IndexOutcome> {
  const input: IndexableItem = {
    itemId: item.itemId,
    driveId: item.driveId,
    parentId: item.parentId,
    name: item.name,
    path: item.path,
    webUrl: item.webUrl,
    size: item.size,
    lastModified: item.lastModified,
    mime: item.mime,
    quickxor: item.quickxor,
  };
  try {
    const stream = await deps.graph.downloadStream(
      `/drives/${input.driveId}/items/${input.itemId}/content`,
      { accessToken }
    );
    try {
      if (stream.contentLength !== null && stream.contentLength > active.config.maxFileBytes) {
        await stream.body.cancel().catch(() => undefined);
        active.store.upsertItem({ ...input, state: 'skipped', skipReason: 'too_large' });
        return 'skipped';
      }
      const bytes = await readBytesWithCap(stream.body, active.config.maxFileBytes);
      const text = await extractDocumentText(bytes, kind);
      const capped = text.slice(0, active.config.maxCharsPerDoc);
      if (!capped.trim()) {
        active.store.upsertItem({ ...input, state: 'skipped', skipReason: 'empty' });
        return 'skipped';
      }
      const pieces = chunkText(capped, { maxChunks: MAX_CHUNKS_PER_DOC });
      if (pieces.length === 0) {
        active.store.upsertItem({ ...input, state: 'skipped', skipReason: 'empty' });
        return 'skipped';
      }
      const embeddings = await embedChunks(active.embed, pieces);
      active.store.replaceItemChunks(
        input,
        pieces.map((piece, index) => ({ text: piece, embedding: embeddings[index] }))
      );
      return 'indexed';
    } catch (error) {
      await stream.body.cancel().catch(() => undefined);
      throw error;
    }
  } catch (error) {
    const reason = shortError(error);
    try {
      active.store.upsertItem({ ...input, state: 'error', skipReason: reason });
    } catch {
      // The original error is the one to report.
    }
    logger.warn(`[SEARCH INDEX] item failed (${kind}): ${reason}`);
    return 'error';
  }
}

async function embedChunks(embed: EmbedClient, texts: string[]): Promise<Float32Array[]> {
  const vectors: Float32Array[] = [];
  for (let index = 0; index < texts.length; index += EMBED_BATCH_SIZE) {
    const batch = texts.slice(index, index + EMBED_BATCH_SIZE);
    const result = await embed.embedBatch(batch);
    vectors.push(...result.vectors);
  }
  return vectors;
}

/** Read a Graph byte stream into memory, refusing anything over `cap` bytes. */
async function readBytesWithCap(body: NonNullable<Response['body']>, cap: number): Promise<Buffer> {
  const reader = body.getReader();
  const chunks: Buffer[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (!value || value.byteLength === 0) continue;
    total += value.byteLength;
    if (total > cap) {
      await reader.cancel().catch(() => undefined);
      throw new Error(`Document is larger than the ${cap} byte limit; refusing to parse it.`);
    }
    chunks.push(Buffer.from(value));
  }
  return Buffer.concat(chunks);
}

function surveyItem(
  active: SearchIndexRuntime,
  item: Record<string, unknown>,
  summary: DriveSyncSummary,
  survey: NonNullable<DriveSyncSummary['survey']>
): void {
  if (item.deleted !== undefined) return;
  if (item.folder !== undefined) {
    summary.folders += 1;
    return;
  }
  if (item.file === undefined) {
    summary.skipped += 1;
    return;
  }
  const metadata = deltaItemMetadata(item);
  const size = metadata.size ?? 0;
  if (!metadata.kind) {
    survey.unsupported.files += 1;
    survey.unsupported.bytes += size;
    return;
  }
  const bucket = (survey.byKind[metadata.kind] ??= { files: 0, bytes: 0, oversized: 0 });
  bucket.files += 1;
  bucket.bytes += size;
  if (size > active.config.maxFileBytes) bucket.oversized += 1;
}

// ---------------------------------------------------------------------------
// Search
// ---------------------------------------------------------------------------

export interface SemanticSearchParams {
  query: string;
  driveId?: string;
  pathPrefix?: string;
  topK?: number;
  hybrid?: boolean;
  minSimilarity?: number;
}

export interface SemanticSearchHit {
  name?: string;
  webUrl?: string;
  path?: string;
  driveId: string;
  lastModified?: string;
  mime?: string;
  size?: number;
  score: number;
  similarity?: number;
  snippet: string;
}

export interface SemanticSearchResult {
  hits: SemanticSearchHit[];
  notes: string[];
}

export async function semanticSearch(params: SemanticSearchParams): Promise<SemanticSearchResult> {
  const active = requireRuntime();
  const topK = clamp(Math.floor(params.topK ?? 8), 1, 25);
  const hybrid = params.hybrid ?? true;
  if (active.store.searchableChunkCount() === 0) {
    return {
      hits: [],
      notes: ['The search index is empty; run refresh-search-index first.'],
    };
  }

  const notes: string[] = [];
  const filtered = Boolean(params.driveId || params.pathPrefix);
  const fetchSize = filtered ? 200 : Math.max(50, topK * 5);

  let vectorRanked: number[] = [];
  const similarityByRowid = new Map<number, number>();
  if (active.store.vectorEnabled) {
    try {
      const queryVector = await embedQuery(active.embed, params.query);
      const vectorRows = active.store.vectorSearch(queryVector, fetchSize);
      vectorRanked = vectorRows.map((row) => row.rowid);
      // sqlite-vec reports euclidean distance and the embedding model returns
      // unit-norm vectors, so cosine similarity is 1 - d²/2.
      for (const row of vectorRows) {
        const cosine = 1 - (row.distance * row.distance) / 2;
        similarityByRowid.set(row.rowid, clamp(cosine, -1, 1));
      }
    } catch (error) {
      notes.push(`Vector search unavailable (${shortError(error)}); used lexical matching only.`);
    }
  } else {
    notes.push('Vector search is not available on this host; used lexical matching only.');
  }

  const ftsQuery = buildFtsQuery(params.query);
  const lexicalRanked = ftsQuery
    ? active.store.lexicalSearch(ftsQuery, fetchSize).map((row) => row.rowid)
    : [];

  const fused = fuseRankings(vectorRanked, lexicalRanked, hybrid);
  const ranked = [...fused.entries()].sort((a, b) => b[1] - a[1]);
  const rows = active.store.chunksWithItems(ranked.slice(0, topK * 4).map(([rowid]) => rowid));
  const rowsById = new Map(rows.map((row) => [row.rowid, row]));

  const bestByItem = new Map<string, { row: SearchChunkRow; score: number; similarity?: number }>();
  for (const [rowid, score] of ranked) {
    const row = rowsById.get(rowid);
    if (!row || !matchesFilters(row, params)) continue;
    const current = bestByItem.get(row.itemId);
    if (!current || score > current.score) {
      bestByItem.set(row.itemId, { row, score, similarity: similarityByRowid.get(rowid) });
    }
  }

  let hits: SemanticSearchHit[] = [...bestByItem.values()]
    .sort((a, b) => b.score - a.score)
    .slice(0, topK)
    .map(({ row, score, similarity }) => ({
      name: row.name,
      webUrl: row.webUrl,
      path: row.path,
      driveId: row.driveId,
      lastModified: row.lastModified,
      mime: row.mime,
      size: row.size,
      score: Number(score.toFixed(6)),
      ...(similarity !== undefined ? { similarity: Number(similarity.toFixed(4)) } : {}),
      snippet: buildSnippet(row.text, params.query),
    }));

  if (params.minSimilarity !== undefined) {
    const minSimilarity = clamp(params.minSimilarity, 0, 1);
    if (hits.every((hit) => hit.similarity === undefined)) {
      notes.push('minSimilarity was ignored because no vector similarities were available.');
    } else {
      const before = hits.length;
      hits = hits.filter((hit) => hit.similarity !== undefined && hit.similarity >= minSimilarity);
      const dropped = before - hits.length;
      if (dropped > 0) notes.push(`${dropped} hit(s) below minSimilarity were dropped.`);
    }
  }

  if (hits.length === 0) notes.push('No matches.');
  return { hits, notes };
}

async function embedQuery(embed: EmbedClient, query: string): Promise<Float32Array> {
  const result = await embed.embedBatch([query.slice(0, MAX_QUERY_CHARS)]);
  const vector = result.vectors[0];
  if (!vector) throw new Error('The embedding response contained no vector.');
  return vector;
}

/**
 * Reciprocal-rank fusion over the vector and lexical rankings; both lists are
 * already best-first, so a mid-ranked match in either contributes without
 * needing calibrated scores from two very different scorers.
 */
export function fuseRankings(
  vectorRanked: number[],
  lexicalRanked: number[],
  hybrid: boolean
): Map<number, number> {
  const scores = new Map<number, number>();
  const add = (rowids: number[]): void => {
    rowids.forEach((rowid, index) => {
      scores.set(rowid, (scores.get(rowid) ?? 0) + 1 / (RRF_K + index + 1));
    });
  };
  if (hybrid) {
    add(vectorRanked);
    add(lexicalRanked);
  } else if (vectorRanked.length > 0) {
    add(vectorRanked);
  } else {
    add(lexicalRanked);
  }
  return scores;
}

/** FTS5 query: every query token quoted and OR'ed; undefined when none qualifies. */
export function buildFtsQuery(query: string): string | undefined {
  const tokens = tokenize(query)
    .filter((token) => token.length > 1)
    .slice(0, 12);
  if (tokens.length === 0) return undefined;
  return tokens.map((token) => `"${token.replace(/"/g, '""')}"`).join(' OR ');
}

export function buildSnippet(text: string, query: string, maxChars = SNAPSHOT_CHARS): string {
  const collapsed = text.replace(/\s+/g, ' ').trim();
  if (collapsed.length <= maxChars) return collapsed;
  const lower = collapsed.toLowerCase();
  let hit = -1;
  for (const token of tokenize(query)) {
    const index = lower.indexOf(token);
    if (index >= 0 && (hit === -1 || index < hit)) hit = index;
  }
  const start = hit <= 0 ? 0 : Math.max(0, hit - Math.floor(maxChars / 3));
  const end = Math.min(collapsed.length, start + maxChars);
  return `${start > 0 ? '…' : ''}${collapsed.slice(start, end)}${end < collapsed.length ? '…' : ''}`;
}

function matchesFilters(
  row: SearchChunkRow,
  filters: { driveId?: string; pathPrefix?: string }
): boolean {
  if (filters.driveId && row.driveId !== filters.driveId) return false;
  if (filters.pathPrefix) {
    const prefix = filters.pathPrefix.toLowerCase();
    if (!row.path || !row.path.toLowerCase().startsWith(prefix)) return false;
  }
  return true;
}

// ---------------------------------------------------------------------------
// Refresh tool entry point
// ---------------------------------------------------------------------------

export interface RefreshSearchIndexParams {
  driveId?: string;
  full?: boolean;
  dryRun?: boolean;
  maxMinutes?: number;
}

export interface RefreshSearchIndexResult {
  mode: 'delta' | 'full' | 'dry-run';
  drives: DriveSyncSummary[];
  elapsedMs: number;
}

export async function refreshSearchIndex(
  deps: SearchIndexDeps,
  params: RefreshSearchIndexParams
): Promise<RefreshSearchIndexResult> {
  const active = requireRuntime();
  if (active.syncing) {
    throw new Error('A search-index sync is already running; try again once it finishes.');
  }
  const driveIds = params.driveId ? [params.driveId] : active.config.driveIds;
  for (const driveId of driveIds) {
    if (!active.config.driveIds.includes(driveId)) {
      throw new Error(
        `driveId is not in MS365_MCP_INDEX_DRIVES; configured drives: ${active.config.driveIds.join(', ')}`
      );
    }
  }
  const full = params.full ?? false;
  const dryRun = params.dryRun ?? false;
  const maxMinutes = clamp(Math.round(params.maxMinutes ?? 5), 1, 30);
  const cold = driveIds.some((driveId) => !active.store.getDriveState(driveId)?.deltaLink);
  // The item budget bounds real work (downloads, extraction, embeddings). A
  // dry-run only walks metadata, so it runs unbounded and the time budget
  // alone stops it; the survey then covers the whole drive in one call.
  const maxItems = dryRun
    ? Number.POSITIVE_INFINITY
    : cold
      ? active.config.maxItemsPerRun * COLD_RUN_MULTIPLIER
      : active.config.maxItemsPerRun;

  active.syncing = true;
  try {
    const result = await runSync(deps, { driveIds, full, dryRun, maxMinutes, maxItems });
    return {
      mode: dryRun ? 'dry-run' : full ? 'full' : 'delta',
      drives: result.drives,
      elapsedMs: result.elapsedMs,
    };
  } finally {
    active.syncing = false;
  }
}

// ---------------------------------------------------------------------------
// Status tool entry point
// ---------------------------------------------------------------------------

export interface SearchIndexStatus {
  enabled: boolean;
  ready: boolean;
  vectorBackend: 'sqlite-vec' | 'fts5-only' | 'none';
  model: string;
  dims: number;
  syncMinutes: number;
  drives: string[];
  syncing: boolean;
  documents: number;
  chunks: number;
  byDrive: Array<{ driveId: string; documents: number; errors: number; chunks: number }>;
  driveStates: Array<{
    driveId: string;
    lastRunAt?: string;
    lastStatus?: string;
    hasDeltaToken: boolean;
  }>;
  dbBytes?: number;
  reason?: string;
}

export function searchIndexStatus(env: Env = process.env): SearchIndexStatus {
  const config = loadSearchIndexConfig(env);
  const base: SearchIndexStatus = {
    enabled: config.enabled,
    ready: false,
    vectorBackend: 'none',
    model: config.embedModel,
    dims: config.embedDims,
    syncMinutes: config.syncMinutes,
    drives: config.driveIds,
    syncing: false,
    documents: 0,
    chunks: 0,
    byDrive: [],
    driveStates: [],
  };
  const active = runtime;
  if (!active) {
    return { ...base, reason: blockReason ?? 'The search index is not initialized.' };
  }
  const counts = active.store.counts();
  return {
    ...base,
    ready: true,
    vectorBackend: active.store.vectorEnabled ? 'sqlite-vec' : 'fts5-only',
    syncing: active.syncing,
    documents: counts.documents,
    chunks: counts.chunks,
    byDrive: counts.byDrive,
    driveStates: active.store.listDriveStates().map((state) => ({
      driveId: state.driveId,
      lastRunAt: state.lastRunAt,
      lastStatus: state.lastStatus,
      hasDeltaToken: Boolean(state.deltaLink),
    })),
    dbBytes: fileSize(active.store.dbPath),
  };
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function shortError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  return message.length > 200 ? `${message.slice(0, 200)}…` : message;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return isRecord(value) ? value : undefined;
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function asNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

/** Graph delta links come absolute; the client takes a version-relative path. */
export function toGraphDeltaPath(url: string): string {
  let path = url;
  try {
    const parsed = new URL(url);
    path = `${parsed.pathname}${parsed.search}`;
  } catch {
    // Already a relative path.
  }
  return path.replace(/^\/(v1\.0|beta)(?=\/)/, '');
}

function fileSize(path: string): number | undefined {
  try {
    return statSync(path).size;
  } catch {
    return undefined;
  }
}

/** Test seam: forget the runtime and drop the default db path marker. */
export function resetSearchIndexForTests(): void {
  stopSearchIndex();
}

export { DEFAULT_INDEX_DB_PATH };
