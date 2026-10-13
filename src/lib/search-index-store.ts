/**
 * SQLite store for the semantic search index.
 *
 * One file on the service volume holds everything: the per-drive Graph delta
 * tokens, drive item metadata, chunk text, an FTS5 lexical index and (when the
 * sqlite-vec extension loads) a vec0 KNN index over the chunk embeddings. The
 * store is derived data only -- it can be deleted and rebuilt from Graph at
 * any time.
 *
 * node:sqlite (Node 22.13+/24) is imported dynamically so a host without it
 * can still load this module; `openSearchIndexStore` then throws with a clear
 * message at the point where an index is actually requested. The sqlite-vec
 * loader is attempted last and only degrades vector search (FTS5 keeps
 * working) when the extension cannot load.
 */

import { mkdirSync } from 'fs';
import { dirname } from 'path';

export const INDEX_SCHEMA_VERSION = 1;

// Minimal structural types over the dynamically imported modules: only the
// surface the store uses is covered, and no node:sqlite types leak into the
// rest of the build.
interface SqliteStatement {
  run(...params: unknown[]): { changes?: number | bigint; lastInsertRowid?: number | bigint };
  get(...params: unknown[]): Record<string, unknown> | undefined;
  all(...params: unknown[]): Array<Record<string, unknown>>;
}

interface SqliteDatabase {
  exec(sql: string): void;
  prepare(sql: string): SqliteStatement;
  close(): void;
  loadExtension?(path: string): void;
}

interface SqliteModule {
  DatabaseSync: new (path: string, options?: { allowExtension?: boolean }) => SqliteDatabase;
}

type SqliteVecModule = {
  load?: (db: unknown) => void;
  default?: { load?: (db: unknown) => void };
};

/** True when this Node build ships node:sqlite (22.13+/23.4+/24). */
export async function isNodeSqliteAvailable(): Promise<boolean> {
  try {
    await import('node:sqlite');
    return true;
  } catch {
    return false;
  }
}

async function loadSqliteModule(): Promise<SqliteModule> {
  try {
    const imported: unknown = await import('node:sqlite');
    const candidate = (imported as { default?: unknown }).default ?? imported;
    if (typeof (candidate as { DatabaseSync?: unknown }).DatabaseSync !== 'function') {
      throw new Error('node:sqlite did not expose DatabaseSync');
    }
    return candidate as SqliteModule;
  } catch {
    throw new Error(
      'The search index needs node:sqlite, which this Node build does not provide (requires Node 22.13+ or 24).'
    );
  }
}

async function tryLoadVectorExtension(db: SqliteDatabase): Promise<boolean> {
  if (typeof db.loadExtension !== 'function') return false;
  let loader: ((db: unknown) => void) | undefined;
  try {
    const imported = (await import('@photostructure/sqlite-vec')) as unknown as SqliteVecModule;
    loader = imported.default?.load ?? imported.load;
  } catch {
    return false;
  }
  if (typeof loader !== 'function') return false;
  try {
    loader(db);
    return true;
  } catch {
    return false;
  }
}

/** Parse a persisted run record; malformed rows fall back to "no record". */
function parseRunRecord(value: unknown): DriveRunRecord | undefined {
  if (value === null || value === undefined) return undefined;
  try {
    const parsed: unknown = JSON.parse(String(value));
    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
      return parsed as DriveRunRecord;
    }
  } catch {
    // Ignore malformed bookkeeping; the rest of the status call still works.
  }
  return undefined;
}

const KNOWN_SKIP_REASONS = new Set(['unsupported', 'package', 'too_large', 'empty']);

/** Map a stored skip/error reason to a short, safe code for the status output. */
function normalizeReason(reason: string | undefined): string {
  if (!reason) return 'unknown';
  if (KNOWN_SKIP_REASONS.has(reason)) return reason;
  if (/too large|expand past|size limit/i.test(reason)) return 'too_large';
  if (/not read|could not be read|corrupt|password|encrypted|unreadable|no text/i.test(reason)) {
    return 'unreadable';
  }
  if (/timeout|abort/i.test(reason)) return 'timeout';
  if (/graph|http \d|fetch|network|ENOTFOUND|ECONN/i.test(reason)) return 'graph_error';
  return 'error';
}

/** Counts persisted for one drive after a sync run, shown by `search-index-status`. */
export interface DriveRunRecord {
  at: string;
  status: string;
  seen: number;
  folders: number;
  indexed: number;
  unchanged: number;
  skipped: number;
  deleted: number;
  errors: number;
  moreRemaining: boolean;
  elapsedMs: number;
}

export interface DriveState {
  driveId: string;
  deltaLink?: string;
  lastRunAt?: string;
  lastStatus?: string;
  lastRunSummary?: DriveRunRecord;
}

export interface ItemRecord {
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
  state: string;
  skipReason?: string;
  chunkCount: number;
  indexedAt?: string;
}

export interface ChunkInsert {
  text: string;
  embedding: Float32Array;
}

export interface SearchChunkRow {
  rowid: number;
  itemId: string;
  chunkIx: number;
  text: string;
  name?: string;
  webUrl?: string;
  path?: string;
  driveId: string;
  lastModified?: string;
  mime?: string;
  size?: number;
}

export interface SearchIndexStoreOptions {
  dbPath: string;
  dims: number;
  embedModel: string;
}

function toText(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function toNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

function asItemRecord(row: Record<string, unknown>): ItemRecord {
  return {
    itemId: String(row.item_id),
    driveId: String(row.drive_id),
    parentId: toText(row.parent_id),
    name: toText(row.name),
    path: toText(row.path),
    webUrl: toText(row.web_url),
    size: toNumber(row.size),
    lastModified: toText(row.last_modified),
    mime: toText(row.mime),
    quickxor: toText(row.quickxor),
    state: toText(row.state) ?? 'pending',
    skipReason: toText(row.skip_reason),
    chunkCount: toNumber(row.chunk_count) ?? 0,
    indexedAt: toText(row.indexed_at),
  };
}

export class SearchIndexStore {
  readonly dbPath: string;
  readonly dims: number;
  readonly embedModel: string;
  readonly vectorEnabled: boolean;

  private readonly db: SqliteDatabase;

  private constructor(
    db: SqliteDatabase,
    options: SearchIndexStoreOptions,
    vectorEnabled: boolean
  ) {
    this.db = db;
    this.dbPath = options.dbPath;
    this.dims = options.dims;
    this.embedModel = options.embedModel;
    this.vectorEnabled = vectorEnabled;
  }

  static async open(options: SearchIndexStoreOptions): Promise<SearchIndexStore> {
    const sqlite = await loadSqliteModule();
    if (dirname(options.dbPath) && dirname(options.dbPath) !== '.') {
      mkdirSync(dirname(options.dbPath), { recursive: true });
    }
    // allowExtension must be set at construction; older Node builds ignore the
    // unknown option and simply cannot load the extension (vector search then
    // degrades to FTS5).
    let db: SqliteDatabase;
    try {
      db = new sqlite.DatabaseSync(options.dbPath, { allowExtension: true });
    } catch {
      db = new sqlite.DatabaseSync(options.dbPath);
    }
    const vectorEnabled = await tryLoadVectorExtension(db);
    const store = new SearchIndexStore(db, options, vectorEnabled);
    store.ensureSchema();
    return store;
  }

  close(): void {
    this.db.close();
  }

  // -- schema ---------------------------------------------------------------

  private ensureSchema(): void {
    this.db.exec('PRAGMA journal_mode = WAL');
    this.db.exec('PRAGMA busy_timeout = 5000');
    this.db.exec('PRAGMA synchronous = NORMAL');
    this.db.exec('CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)');

    const storedVersion = this.getMeta('schema_version');
    const storedDims = this.getMeta('embed_dims');
    const storedModel = this.getMeta('embed_model');
    const mismatch =
      (storedVersion !== undefined && Number(storedVersion) !== INDEX_SCHEMA_VERSION) ||
      (storedDims !== undefined && Number(storedDims) !== this.dims) ||
      (storedModel !== undefined && storedModel !== this.embedModel);
    if (mismatch) {
      // Derived data only: drop and let the next sync rebuild from Graph.
      for (const table of ['chunk_vectors', 'chunk_fts', 'chunks', 'items', 'drives']) {
        this.db.exec(`DROP TABLE IF EXISTS ${table}`);
      }
    }

    this.setMeta('schema_version', String(INDEX_SCHEMA_VERSION));
    this.setMeta('embed_dims', String(this.dims));
    this.setMeta('embed_model', this.embedModel);

    this.db.exec(
      'CREATE TABLE IF NOT EXISTS drives (' +
        'drive_id TEXT PRIMARY KEY, delta_link TEXT, last_run_at TEXT, last_status TEXT, ' +
        'last_run_summary TEXT)'
    );
    // Additive migration for databases created before last_run_summary existed
    // (the column is bookkeeping only, so the index itself stays intact).
    const driveColumns = this.db
      .prepare('PRAGMA table_info(drives)')
      .all()
      .map((row) => String(row.name));
    if (!driveColumns.includes('last_run_summary')) {
      this.db.exec('ALTER TABLE drives ADD COLUMN last_run_summary TEXT');
    }
    this.db.exec(
      'CREATE TABLE IF NOT EXISTS items (' +
        'item_id TEXT PRIMARY KEY, drive_id TEXT NOT NULL, parent_id TEXT, name TEXT, path TEXT, ' +
        'web_url TEXT, size INTEGER, last_modified TEXT, mime TEXT, quickxor TEXT, ' +
        "state TEXT NOT NULL DEFAULT 'pending', skip_reason TEXT, " +
        'chunk_count INTEGER NOT NULL DEFAULT 0, indexed_at TEXT)'
    );
    this.db.exec('CREATE INDEX IF NOT EXISTS items_drive_state ON items (drive_id, state)');
    this.db.exec(
      'CREATE TABLE IF NOT EXISTS chunks (' +
        'rowid INTEGER PRIMARY KEY AUTOINCREMENT, item_id TEXT NOT NULL, chunk_ix INTEGER NOT NULL, ' +
        'text TEXT NOT NULL, char_len INTEGER NOT NULL)'
    );
    this.db.exec('CREATE INDEX IF NOT EXISTS chunks_item ON chunks (item_id)');
    this.db.exec(
      "CREATE VIRTUAL TABLE IF NOT EXISTS chunk_fts USING fts5(text, tokenize='unicode61 remove_diacritics 2')"
    );
    if (this.vectorEnabled) {
      this.db.exec(
        `CREATE VIRTUAL TABLE IF NOT EXISTS chunk_vectors USING vec0(embedding float[${this.dims}])`
      );
    }
  }

  private transaction<T>(run: () => T): T {
    this.db.exec('BEGIN IMMEDIATE');
    try {
      const result = run();
      this.db.exec('COMMIT');
      return result;
    } catch (error) {
      try {
        this.db.exec('ROLLBACK');
      } catch {
        // The original error is the one that matters.
      }
      throw error;
    }
  }

  // -- meta -----------------------------------------------------------------

  getMeta(key: string): string | undefined {
    const row = this.db.prepare('SELECT value FROM meta WHERE key = ?').get(key);
    return row ? String(row.value) : undefined;
  }

  setMeta(key: string, value: string): void {
    this.db
      .prepare(
        'INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value'
      )
      .run(key, value);
  }

  // -- drives ---------------------------------------------------------------

  getDriveState(driveId: string): DriveState | undefined {
    const row = this.db
      .prepare(
        'SELECT drive_id, delta_link, last_run_at, last_status, last_run_summary FROM drives WHERE drive_id = ?'
      )
      .get(driveId);
    if (!row) return undefined;
    return {
      driveId: String(row.drive_id),
      deltaLink: toText(row.delta_link),
      lastRunAt: toText(row.last_run_at),
      lastStatus: toText(row.last_status),
      lastRunSummary: parseRunRecord(row.last_run_summary),
    };
  }

  listDriveStates(): DriveState[] {
    return this.db
      .prepare(
        'SELECT drive_id, delta_link, last_run_at, last_status, last_run_summary FROM drives ORDER BY drive_id'
      )
      .all()
      .map((row) => ({
        driveId: String(row.drive_id),
        deltaLink: toText(row.delta_link),
        lastRunAt: toText(row.last_run_at),
        lastStatus: toText(row.last_status),
        lastRunSummary: parseRunRecord(row.last_run_summary),
      }));
  }

  saveDriveDelta(driveId: string, deltaLink: string, status = 'ok'): void {
    this.db
      .prepare(
        'INSERT INTO drives (drive_id, delta_link, last_run_at, last_status) VALUES (?, ?, ?, ?) ' +
          'ON CONFLICT(drive_id) DO UPDATE SET delta_link = excluded.delta_link, ' +
          'last_run_at = excluded.last_run_at, last_status = excluded.last_status'
      )
      .run(driveId, deltaLink, new Date().toISOString(), status);
  }

  recordDriveRun(driveId: string, status: string): void {
    this.db
      .prepare(
        'INSERT INTO drives (drive_id, delta_link, last_run_at, last_status) VALUES (?, NULL, ?, ?) ' +
          'ON CONFLICT(drive_id) DO UPDATE SET last_run_at = excluded.last_run_at, ' +
          'last_status = excluded.last_status'
      )
      .run(driveId, new Date().toISOString(), status);
  }

  /**
   * Persist the counts of the drive's most recent sync run so a status call can
   * show them even when the sync call itself timed out on the client.
   */
  saveDriveRunSummary(driveId: string, record: DriveRunRecord): void {
    this.db
      .prepare(
        'INSERT INTO drives (drive_id, last_run_summary) VALUES (?, ?) ' +
          'ON CONFLICT(drive_id) DO UPDATE SET last_run_summary = excluded.last_run_summary'
      )
      .run(driveId, JSON.stringify(record));
  }

  /** Drop every derived row for one drive (full re-index path). */
  resetDrive(driveId: string): void {
    const itemIds = this.db
      .prepare('SELECT item_id FROM items WHERE drive_id = ?')
      .all(driveId)
      .map((row) => String(row.item_id));
    this.transaction(() => {
      for (const itemId of itemIds) this.deleteChunksForItem(itemId);
      this.db.prepare('DELETE FROM items WHERE drive_id = ?').run(driveId);
      this.db.prepare('DELETE FROM drives WHERE drive_id = ?').run(driveId);
    });
  }

  // -- items ----------------------------------------------------------------

  getItem(itemId: string): ItemRecord | undefined {
    const row = this.db.prepare('SELECT * FROM items WHERE item_id = ?').get(itemId);
    return row ? asItemRecord(row) : undefined;
  }

  itemsByState(driveId: string, state: string, limit: number): ItemRecord[] {
    return this.db
      .prepare('SELECT * FROM items WHERE drive_id = ? AND state = ? ORDER BY name LIMIT ?')
      .all(driveId, state, limit)
      .map((row) => asItemRecord(row));
  }

  /** Insert or update one item's Graph metadata without touching its chunks. */
  upsertItem(record: {
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
    state?: string;
    skipReason?: string;
  }): void {
    this.db
      .prepare(
        'INSERT INTO items (item_id, drive_id, parent_id, name, path, web_url, size, last_modified, ' +
          'mime, quickxor, state, skip_reason) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) ' +
          'ON CONFLICT(item_id) DO UPDATE SET drive_id = excluded.drive_id, ' +
          'parent_id = excluded.parent_id, name = excluded.name, path = excluded.path, ' +
          'web_url = excluded.web_url, size = excluded.size, last_modified = excluded.last_modified, ' +
          'mime = excluded.mime, quickxor = excluded.quickxor, state = excluded.state, ' +
          'skip_reason = excluded.skip_reason'
      )
      .run(
        record.itemId,
        record.driveId,
        record.parentId ?? null,
        record.name ?? null,
        record.path ?? null,
        record.webUrl ?? null,
        record.size ?? null,
        record.lastModified ?? null,
        record.mime ?? null,
        record.quickxor ?? null,
        record.state ?? 'pending',
        record.skipReason ?? null
      );
  }

  deleteItemWithChunks(itemId: string): void {
    this.transaction(() => {
      this.deleteChunksForItem(itemId);
      this.db.prepare('DELETE FROM items WHERE item_id = ?').run(itemId);
    });
  }

  private deleteChunksForItem(itemId: string): void {
    const rowids = this.db
      .prepare('SELECT rowid FROM chunks WHERE item_id = ?')
      .all(itemId)
      .map((row) => Number(row.rowid));
    for (const rowid of rowids) {
      if (this.vectorEnabled) {
        this.db.prepare('DELETE FROM chunk_vectors WHERE rowid = ?').run(BigInt(rowid));
      }
      this.db.prepare('DELETE FROM chunk_fts WHERE rowid = ?').run(BigInt(rowid));
    }
    this.db.prepare('DELETE FROM chunks WHERE item_id = ?').run(itemId);
  }

  /**
   * Replace one item's chunks atomically: old chunks (and their FTS/vector
   * rows) are removed, the new chunks are written, and the item row flips to
   * state "indexed" in the same transaction, so a crash can never leave an
   * item marked indexed without its chunks.
   */
  replaceItemChunks(
    record: {
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
    },
    chunks: ChunkInsert[]
  ): void {
    const insertChunk = this.db.prepare(
      'INSERT INTO chunks (item_id, chunk_ix, text, char_len) VALUES (?, ?, ?, ?)'
    );
    const insertFts = this.db.prepare('INSERT INTO chunk_fts (rowid, text) VALUES (?, ?)');
    const insertVector = this.vectorEnabled
      ? this.db.prepare('INSERT INTO chunk_vectors (rowid, embedding) VALUES (?, ?)')
      : null;
    const indexedAt = new Date().toISOString();

    this.transaction(() => {
      this.deleteChunksForItem(record.itemId);
      for (let index = 0; index < chunks.length; index++) {
        const chunk = chunks[index];
        const result = insertChunk.run(record.itemId, index, chunk.text, chunk.text.length);
        const rowid = Number(result.lastInsertRowid ?? 0);
        if (!rowid) throw new Error('chunk insert did not return a rowid');
        insertFts.run(BigInt(rowid), chunk.text);
        if (insertVector) insertVector.run(BigInt(rowid), chunk.embedding);
      }
      this.db
        .prepare(
          'INSERT INTO items (item_id, drive_id, parent_id, name, path, web_url, size, last_modified, ' +
            'mime, quickxor, state, skip_reason, chunk_count, indexed_at) ' +
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'indexed', NULL, ?, ?) " +
            'ON CONFLICT(item_id) DO UPDATE SET drive_id = excluded.drive_id, ' +
            'parent_id = excluded.parent_id, name = excluded.name, path = excluded.path, ' +
            'web_url = excluded.web_url, size = excluded.size, last_modified = excluded.last_modified, ' +
            'mime = excluded.mime, quickxor = excluded.quickxor, state = excluded.state, ' +
            'skip_reason = excluded.skip_reason, chunk_count = excluded.chunk_count, ' +
            'indexed_at = excluded.indexed_at'
        )
        .run(
          record.itemId,
          record.driveId,
          record.parentId ?? null,
          record.name ?? null,
          record.path ?? null,
          record.webUrl ?? null,
          record.size ?? null,
          record.lastModified ?? null,
          record.mime ?? null,
          record.quickxor ?? null,
          chunks.length,
          indexedAt
        );
    });
  }

  // -- search ---------------------------------------------------------------

  searchableChunkCount(): number {
    const row = this.db.prepare('SELECT COUNT(*) AS count FROM chunks').get();
    return toNumber(row?.count) ?? 0;
  }

  vectorSearch(query: Float32Array, k: number): Array<{ rowid: number; distance: number }> {
    if (!this.vectorEnabled) return [];
    const capped = Math.max(1, Math.min(Math.floor(k), 1000));
    return this.db
      .prepare(
        `SELECT rowid, distance FROM chunk_vectors WHERE embedding MATCH ? AND k = ${capped} ` +
          'ORDER BY distance'
      )
      .all(query)
      .map((row) => ({ rowid: Number(row.rowid), distance: toNumber(row.distance) ?? 1 }));
  }

  lexicalSearch(ftsQuery: string, limit: number): Array<{ rowid: number; score: number }> {
    if (!ftsQuery) return [];
    const capped = Math.max(1, Math.min(Math.floor(limit), 1000));
    return this.db
      .prepare(
        'SELECT rowid, bm25(chunk_fts) AS score FROM chunk_fts WHERE chunk_fts MATCH ? ' +
          `ORDER BY score LIMIT ${capped}`
      )
      .all(ftsQuery)
      .map((row) => ({ rowid: Number(row.rowid), score: toNumber(row.score) ?? 0 }));
  }

  chunksWithItems(rowids: number[]): SearchChunkRow[] {
    if (rowids.length === 0) return [];
    const placeholders = rowids.map(() => '?').join(', ');
    const rows = this.db
      .prepare(
        'SELECT c.rowid AS rowid, c.item_id AS item_id, c.chunk_ix AS chunk_ix, c.text AS text, ' +
          'i.name AS name, i.web_url AS web_url, i.path AS path, i.drive_id AS drive_id, ' +
          'i.last_modified AS last_modified, i.mime AS mime, i.size AS size ' +
          'FROM chunks c JOIN items i ON i.item_id = c.item_id ' +
          `WHERE c.rowid IN (${placeholders})`
      )
      .all(...rowids.map((rowid) => BigInt(rowid)));
    return rows.map((row) => ({
      rowid: Number(row.rowid),
      itemId: String(row.item_id),
      chunkIx: Number(row.chunk_ix),
      text: String(row.text),
      name: toText(row.name),
      webUrl: toText(row.web_url),
      path: toText(row.path),
      driveId: String(row.drive_id),
      lastModified: toText(row.last_modified),
      mime: toText(row.mime),
      size: toNumber(row.size),
    }));
  }

  // -- status ---------------------------------------------------------------

  counts(): {
    documents: number;
    chunks: number;
    byDrive: Array<{
      driveId: string;
      documents: number;
      errors: number;
      chunks: number;
      reasons: Record<string, number>;
    }>;
  } {
    const documents =
      toNumber(
        this.db.prepare("SELECT COUNT(*) AS count FROM items WHERE state = 'indexed'").get()?.count
      ) ?? 0;
    const chunks = this.searchableChunkCount();
    const reasonsByDrive = new Map<string, Record<string, number>>();
    const reasonRows = this.db
      .prepare(
        'SELECT drive_id, state, skip_reason, COUNT(*) AS count FROM items ' +
          "WHERE state IN ('skipped', 'error') GROUP BY drive_id, state, skip_reason"
      )
      .all();
    for (const row of reasonRows) {
      const bucket = reasonsByDrive.get(String(row.drive_id)) ?? {};
      const key = `${String(row.state)}/${normalizeReason(toText(row.skip_reason))}`;
      bucket[key] = (bucket[key] ?? 0) + (toNumber(row.count) ?? 0);
      reasonsByDrive.set(String(row.drive_id), bucket);
    }
    const byDrive = this.db
      .prepare(
        'SELECT drive_id, ' +
          "SUM(CASE WHEN state = 'indexed' THEN 1 ELSE 0 END) AS documents, " +
          "SUM(CASE WHEN state = 'error' THEN 1 ELSE 0 END) AS errors, " +
          'COALESCE(SUM(chunk_count), 0) AS chunks ' +
          'FROM items GROUP BY drive_id ORDER BY drive_id'
      )
      .all()
      .map((row) => ({
        driveId: String(row.drive_id),
        documents: toNumber(row.documents) ?? 0,
        errors: toNumber(row.errors) ?? 0,
        chunks: toNumber(row.chunks) ?? 0,
        reasons: reasonsByDrive.get(String(row.drive_id)) ?? {},
      }));
    return { documents, chunks, byDrive };
  }
}

export async function openSearchIndexStore(
  options: SearchIndexStoreOptions
): Promise<SearchIndexStore> {
  return SearchIndexStore.open(options);
}
