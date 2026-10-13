/**
 * Char-bounded chunking for the semantic search index.
 *
 * Extraction output is a flat string (pages, sheets and slides already
 * flattened by document-text.ts); this module turns it into overlapping
 * retrieval units. Chunk boundaries prefer paragraph, line and sentence
 * breaks inside a look-back window so a chunk usually ends at a natural
 * boundary instead of mid-word; a hard cut is the fallback.
 *
 * The chunker is deterministic: the same text and options always produce the
 * same chunks, which keeps re-indexing idempotent.
 */

export const DEFAULT_CHUNK_CHARS = 1500;
export const DEFAULT_CHUNK_OVERLAP_CHARS = 200;

const MIN_CHUNK_CHARS = 100;
const MAX_CHUNK_CHARS = 20_000;
const BOUNDARY_LOOKBACK_CHARS = 400;

export interface ChunkTextOptions {
  chunkChars?: number;
  overlapChars?: number;
  /** Hard cap on the number of chunks returned for one document. */
  maxChunks?: number;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

export function chunkText(text: string, options: ChunkTextOptions = {}): string[] {
  const chunkChars = clamp(
    Math.floor(options.chunkChars ?? DEFAULT_CHUNK_CHARS),
    MIN_CHUNK_CHARS,
    MAX_CHUNK_CHARS
  );
  const overlapChars = clamp(
    Math.floor(options.overlapChars ?? DEFAULT_CHUNK_OVERLAP_CHARS),
    0,
    Math.floor(chunkChars / 2)
  );
  const maxChunks = options.maxChunks ?? Number.POSITIVE_INFINITY;

  const normalized = text.replace(/\r\n?/g, '\n');
  const chunks: string[] = [];
  let start = 0;

  while (start < normalized.length && chunks.length < maxChunks) {
    let end = Math.min(start + chunkChars, normalized.length);
    if (end < normalized.length) {
      // Look for a natural boundary in the tail of the window; the earliest
      // match wins so the chunk stays as close to chunkChars as possible.
      const floor = Math.max(start + Math.floor(chunkChars * 0.6), end - BOUNDARY_LOOKBACK_CHARS);
      let boundary = -1;
      for (const token of ['\n\n', '\n', '. ']) {
        const candidate = normalized.lastIndexOf(token, end);
        if (candidate > boundary) boundary = candidate;
      }
      if (boundary >= floor) {
        end = boundary + 1;
      }
    }
    const piece = normalized.slice(start, end).trim();
    if (piece) chunks.push(piece);
    if (end >= normalized.length) break;
    start = Math.max(end - overlapChars, start + 1);
  }

  return chunks;
}
