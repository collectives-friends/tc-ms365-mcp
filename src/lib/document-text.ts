/**
 * Text extraction for the document formats extract-document-text serves.
 *
 * docx/docm go through mammoth (raw text, tables flattened), pdf through
 * unpdf's serverless PDF.js build. Both parsers are imported lazily so the
 * PDF.js payload is only paid for when a PDF is actually read, not on server
 * startup.
 *
 * Resource limits are the caller's job: this module never reads from the
 * network, and extract-document-text caps the bytes (and characters) before
 * parsing. Parsing itself is CPU-bound and runs on the server's event loop, so
 * the byte cap is also the latency guard.
 */

export const DEFAULT_MAX_DOCUMENT_CHARS = 50_000;
export const HARD_MAX_DOCUMENT_CHARS = 200_000;
export const MAX_DOCUMENT_BYTES = 30 * 1024 * 1024;

export type DocumentKind = 'docx' | 'pdf' | 'text';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
// Macro-enabled Word documents share the same OOXML package layout, and mammoth reads them.
const DOCM_MIME = 'application/vnd.ms-word.document.macroenabled.12';
const TEXT_MIME_TYPES = new Set([
  'text/plain',
  'text/markdown',
  'text/x-markdown',
  'text/csv',
  'application/json',
]);
const TEXT_EXTENSIONS = new Set(['txt', 'text', 'md', 'markdown', 'csv', 'json', 'log']);

export const SUPPORTED_DOCUMENT_DESCRIPTION =
  'Word (.docx, .docm), PDF (.pdf), and plain-text (.txt, .md, .csv, .json, .log) documents';

function fileExtension(fileName: string | undefined): string | undefined {
  if (!fileName) return undefined;
  const match = /\.([A-Za-z0-9]+)$/.exec(fileName);
  return match?.[1]?.toLowerCase();
}

/** Lowercase a MIME type and drop parameters ("application/pdf; charset=binary"). */
export function normalizeMimeType(mimeType: string | undefined): string {
  return (mimeType ?? '').split(';')[0].trim().toLowerCase();
}

/**
 * Map a file name and MIME type to a parser kind, or undefined when neither
 * identifies a supported document. Extension and MIME are checked independently:
 * SharePoint reports generic content types for some uploads, and vice versa.
 */
export function detectDocumentKind(
  mimeType: string | undefined,
  fileName: string | undefined
): DocumentKind | undefined {
  const mime = normalizeMimeType(mimeType);
  const ext = fileExtension(fileName);
  if (mime === DOCX_MIME || mime === DOCM_MIME || ext === 'docx' || ext === 'docm') return 'docx';
  if (mime === 'application/pdf' || ext === 'pdf') return 'pdf';
  if (TEXT_MIME_TYPES.has(mime) || (ext !== undefined && TEXT_EXTENSIONS.has(ext))) return 'text';
  return undefined;
}

type MammothModule = {
  extractRawText(input: { buffer: Buffer }): Promise<{ value: string }>;
};

/** Resolve mammoth from either its CJS export shape or an ESM namespace. */
async function loadMammoth(): Promise<MammothModule> {
  const imported: unknown = await import('mammoth');
  const candidate = (imported as { default?: unknown }).default ?? imported;
  const extract = (candidate as { extractRawText?: unknown }).extractRawText;
  if (typeof extract !== 'function') {
    throw new Error('mammoth did not expose extractRawText');
  }
  return candidate as MammothModule;
}

/**
 * unpdf's published type surface imports the optional native @napi-rs/canvas
 * package (rendering), which this server never installs or uses. Loading the
 * module through a non-literal specifier keeps the type graph out of the build
 * - the runtime import itself is unchanged - and the local type below covers
 * the two text-extraction entry points we call.
 */
const UNPDF_SPECIFIER = 'unpdf';

type UnpdfModule = {
  getDocumentProxy(data: Uint8Array): Promise<unknown>;
  extractText(
    data: unknown,
    options?: { mergePages?: boolean }
  ): Promise<{ totalPages: number; text: string | string[] }>;
};

async function loadUnpdf(): Promise<UnpdfModule> {
  const imported: unknown = await import(UNPDF_SPECIFIER);
  const candidate = (imported as { default?: unknown }).default ?? imported;
  const extractText = (candidate as { extractText?: unknown }).extractText;
  const getDocumentProxy = (candidate as { getDocumentProxy?: unknown }).getDocumentProxy;
  if (typeof extractText !== 'function' || typeof getDocumentProxy !== 'function') {
    throw new Error('unpdf did not expose extractText/getDocumentProxy');
  }
  return candidate as UnpdfModule;
}

/**
 * Extract the full plain text of `bytes`. Throws with a user-presentable
 * message for corrupt, encrypted or otherwise unreadable documents.
 */
export async function extractDocumentText(bytes: Buffer, kind: DocumentKind): Promise<string> {
  if (bytes.length === 0) {
    throw new Error('The document is empty (0 bytes).');
  }
  if (kind === 'text') {
    return bytes.toString('utf8');
  }
  if (kind === 'docx') {
    const mammoth = await loadMammoth();
    const result = await mammoth.extractRawText({ buffer: bytes });
    return result.value;
  }
  const { extractText, getDocumentProxy } = await loadUnpdf();
  // Pass a view instead of copying the Buffer's bytes into a new Uint8Array.
  const data = new Uint8Array(bytes.buffer as ArrayBuffer, bytes.byteOffset, bytes.byteLength);
  const pdf = await getDocumentProxy(data);
  const result = await extractText(pdf, { mergePages: true });
  return typeof result.text === 'string' ? result.text : result.text.join('\n\n');
}

/**
 * Slice extraction output for one response, keeping payloads bounded.
 * `truncated` + `nextOffset` let the caller continue where this call stopped.
 */
export function pageDocumentText(
  text: string,
  offset: number,
  maxChars: number
): { text: string; truncated: boolean; nextOffset?: number } {
  const start = Math.min(offset, text.length);
  const end = Math.min(start + maxChars, text.length);
  const truncated = end < text.length;
  return {
    text: text.slice(start, end),
    truncated,
    ...(truncated ? { nextOffset: end } : {}),
  };
}
