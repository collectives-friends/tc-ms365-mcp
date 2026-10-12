/**
 * Environment configuration for the semantic search index.
 *
 * Every value is read once per call from the environment (names only in the
 * docs; the values are set on the service). The index is off unless
 * MS365_MCP_INDEX_ENABLED is truthy, which keeps the tool surface and the
 * worker inert in every environment that has not opted in.
 */

import { MAX_DOCUMENT_BYTES } from './document-text.js';
import { MAX_UNCOMPRESSED_PART_BYTES } from './ooxml-text.js';
import {
  DEFAULT_EMBED_DIMS,
  DEFAULT_EMBED_MODEL,
  DEFAULT_OPENAI_BASE_URL,
} from './search-index-embed.js';

export const DEFAULT_INDEX_DB_PATH = '/data/search-index.sqlite';
export const DEFAULT_MAX_ITEMS_PER_RUN = 50;
export const DEFAULT_MAX_CHARS_PER_DOC = 500_000;
export const DEFAULT_SYNC_MINUTES = 15;

/** Environment map shape used across the index modules (process.env fits). */
export type Env = Record<string, string | undefined>;

export interface SearchIndexConfig {
  enabled: boolean;
  dbPath: string;
  driveIds: string[];
  syncMinutes: number;
  maxItemsPerRun: number;
  maxCharsPerDoc: number;
  maxFileBytes: number;
  maxPartBytes: number;
  embedModel: string;
  embedDims: number;
  openaiApiKey?: string;
  openaiBaseUrl: string;
}

const TRUTHY = new Set(['1', 'true', 'yes', 'on']);

function intFromEnv(value: string | undefined, fallback: number, min: number, max: number): number {
  const parsed = value === undefined ? Number.NaN : Number.parseInt(value, 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(Math.max(parsed, min), max);
}

export function isSearchIndexEnabled(env: Env = process.env): boolean {
  return TRUTHY.has((env.MS365_MCP_INDEX_ENABLED ?? '').trim().toLowerCase());
}

export function loadSearchIndexConfig(env: Env = process.env): SearchIndexConfig {
  return {
    enabled: isSearchIndexEnabled(env),
    dbPath: env.MS365_MCP_INDEX_DB?.trim() || DEFAULT_INDEX_DB_PATH,
    driveIds: (env.MS365_MCP_INDEX_DRIVES ?? '')
      .split(',')
      .map((id) => id.trim())
      .filter(Boolean),
    syncMinutes: intFromEnv(env.MS365_MCP_INDEX_SYNC_MINUTES, DEFAULT_SYNC_MINUTES, 0, 24 * 60),
    maxItemsPerRun: intFromEnv(
      env.MS365_MCP_INDEX_MAX_ITEMS_PER_RUN,
      DEFAULT_MAX_ITEMS_PER_RUN,
      1,
      10_000
    ),
    maxCharsPerDoc: intFromEnv(
      env.MS365_MCP_INDEX_MAX_CHARS_PER_DOC,
      DEFAULT_MAX_CHARS_PER_DOC,
      1_000,
      5_000_000
    ),
    maxFileBytes: intFromEnv(
      env.MS365_MCP_INDEX_MAX_FILE_BYTES,
      MAX_DOCUMENT_BYTES,
      1024,
      200 * 1024 * 1024
    ),
    maxPartBytes: intFromEnv(
      env.MS365_MCP_INDEX_MAX_PART_BYTES,
      MAX_UNCOMPRESSED_PART_BYTES,
      1024 * 1024,
      256 * 1024 * 1024
    ),
    embedModel: env.OPENAI_EMBED_MODEL?.trim() || DEFAULT_EMBED_MODEL,
    embedDims: intFromEnv(env.OPENAI_EMBED_DIMS, DEFAULT_EMBED_DIMS, 64, 3072),
    openaiApiKey: env.OPENAI_API_KEY?.trim() || undefined,
    openaiBaseUrl: env.OPENAI_BASE_URL?.trim() || DEFAULT_OPENAI_BASE_URL,
  };
}

/** Validate an enabled configuration; returns the reason it cannot run, if any. */
export function validateSearchIndexConfig(config: SearchIndexConfig): string | undefined {
  if (!config.enabled) return undefined;
  if (config.driveIds.length === 0) {
    return 'MS365_MCP_INDEX_DRIVES must list at least one drive id when the index is enabled.';
  }
  if (!config.openaiApiKey) {
    return 'OPENAI_API_KEY must be set when the index is enabled.';
  }
  return undefined;
}
