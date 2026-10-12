/**
 * Text extraction for the document formats extract-document-text serves.
 *
 * docx/docm go through mammoth (raw text, tables flattened), pdf through
 * unpdf's serverless PDF.js build, and xlsx/xlsm + pptx/pptm through the
 * OOXML reader in ooxml-text.ts. All parsers are imported lazily so the
 * PDF.js payload is only paid for when a PDF is actually read, not on server
 * startup.
 *
 * Resource limits are the caller's job: this module never reads from the
 * network, and extract-document-text caps the bytes before parsing. Extraction
 * itself is additionally bounded by EXTRACTION_MAX_CHARS so no single document
 * can expand into unbounded text, and the PDF walk releases each page as it
 * goes instead of holding the whole document open. Parsing is CPU-bound and
 * runs on the server's event loop, so the byte cap is also the latency guard.
 */

export const DEFAULT_MAX_DOCUMENT_CHARS = 50_000;
export const HARD_MAX_DOCUMENT_CHARS = 200_000;
export const MAX_DOCUMENT_BYTES = 30 * 1024 * 1024;
/**
 * Upper bound on the text one extraction pass materializes. The OOXML reader
 * applies the same bound internally (OOXML_MAX_TEXT_CHARS); without it a PDF
 * would stream every page of the document into one string.
 */
export const EXTRACTION_MAX_CHARS = 2_000_000;

export type DocumentKind = 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'text';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
// Macro-enabled Word documents share the same OOXML package layout, and mammoth reads them.
const DOCM_MIME = 'application/vnd.ms-word.document.macroenabled.12';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const XLSM_MIME = 'application/vnd.ms-excel.sheet.macroenabled.12';
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const PPTM_MIME = 'application/vnd.ms-powerpoint.presentation.macroenabled.12';
const TEXT_MIME_TYPES = new Set([
  'text/plain',
  'text/markdown',
  'text/x-markdown',
  'text/csv',
  'application/json',
]);
const TEXT_EXTENSIONS = new Set(['txt', 'text', 'md', 'markdown', 'csv', 'json', 'log']);

export const SUPPORTED_DOCUMENT_DESCRIPTION =
  'Word (.docx, .docm), Excel (.xlsx, .xlsm), PowerPoint (.pptx, .pptm), PDF (.pdf), and plain-text (.txt, .md, .csv, .json, .log) documents';

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
  if (mime === XLSX_MIME || mime === XLSM_MIME || ext === 'xlsx' || ext === 'xlsm') return 'xlsx';
  if (mime === PPTX_MIME || mime === PPTM_MIME || ext === 'pptx' || ext === 'pptm') return 'pptx';
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
 * - the runtime import itself is unchanged - and the local types below cover
 * the page-by-page surface the PDF walk uses.
 */
const UNPDF_SPECIFIER = 'unpdf';

interface PdfTextItem {
  str?: string;
  hasEOL?: boolean;
}

interface PdfPageHandle {
  getTextContent(): Promise<{ items: PdfTextItem[] }>;
  cleanup?(): unknown;
}

interface PdfDocumentHandle {
  numPages: number;
  getPage(pageNumber: number): Promise<PdfPageHandle>;
  destroy?(): unknown;
}

type UnpdfModule = {
  getDocumentProxy(data: Uint8Array): Promise<PdfDocumentHandle>;
};

async function loadUnpdf(): Promise<UnpdfModule> {
  const imported: unknown = await import(UNPDF_SPECIFIER);
  const candidate = (imported as { default?: unknown }).default ?? imported;
  const getDocumentProxy = (candidate as { getDocumentProxy?: unknown }).getDocumentProxy;
  if (typeof getDocumentProxy !== 'function') {
    throw new Error('unpdf did not expose getDocumentProxy');
  }
  return candidate as UnpdfModule;
}

/**
 * Extract the full plain text of `bytes`, at most `maxChars` characters.
 * Throws with a user-presentable message for corrupt, encrypted or otherwise
 * unreadable documents.
 */
export async function extractDocumentText(
  bytes: Buffer,
  kind: DocumentKind,
  options: { maxChars?: number } = {}
): Promise<string> {
  const maxChars = options.maxChars ?? EXTRACTION_MAX_CHARS;
  if (bytes.length === 0) {
    throw new Error('The document is empty (0 bytes).');
  }
  if (kind === 'text') {
    return bytes.toString('utf8').slice(0, maxChars);
  }
  if (kind === 'docx') {
    const mammoth = await loadMammoth();
    const result = await mammoth.extractRawText({ buffer: bytes });
    return result.value.slice(0, maxChars);
  }
  if (kind === 'xlsx' || kind === 'pptx') {
    const ooxml = await import('./ooxml-text.js');
    return kind === 'xlsx' ? ooxml.extractXlsxText(bytes) : ooxml.extractPptxText(bytes);
  }
  // Pass a view instead of copying the Buffer's bytes into a new Uint8Array.
  const data = new Uint8Array(bytes.buffer as ArrayBuffer, bytes.byteOffset, bytes.byteLength);
  return extractPdfText(data, maxChars);
}

/**
 * unpdf's `extractText` starts every page at once (`Promise.all` over the page
 * range) and never releases a document the caller handed it, so a large PDF
 * both spikes and leaks memory. This walk reads one page at a time, cleans the
 * page up before moving on, stops once `maxChars` are collected, and destroys
 * the document in every case. Text assembly mirrors unpdf's: items joined per
 * page with `hasEOL` line breaks, pages joined with newlines and whitespace
 * collapsed without flattening lines.
 */
async function extractPdfText(data: Uint8Array, maxChars: number): Promise<string> {
  const { getDocumentProxy } = await loadUnpdf();
  const pdf = await getDocumentProxy(data);
  try {
    const pages: string[] = [];
    let collected = 0;
    for (let pageNumber = 1; pageNumber <= pdf.numPages && collected < maxChars; pageNumber++) {
      const page = await pdf.getPage(pageNumber);
      try {
        const content = await page.getTextContent();
        const text = content.items
          .filter((item) => item.str != null)
          .map((item) => `${item.str ?? ''}${item.hasEOL ? '\n' : ''}`)
          .join('');
        pages.push(text);
        collected += text.length + 1;
      } finally {
        page.cleanup?.();
      }
    }
    return normalizeMergedPdfText(pages).slice(0, maxChars);
  } finally {
    try {
      await pdf.destroy?.();
    } catch {
      // Cleanup is best-effort; a destroy failure must not mask the result.
    }
  }
}

function normalizeMergedPdfText(pages: string[]): string {
  return pages
    .join('\n')
    .replace(/[^\S\n]+/g, ' ')
    .replace(/ ?\n ?/g, '\n')
    .replace(/\n{3,}/g, '\n\n');
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
