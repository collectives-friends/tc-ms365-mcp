import { mkdtempSync, rmSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import JSZip from 'jszip';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { detectDocumentKind, extractDocumentText } from '../src/lib/document-text.js';
import { MAX_UNCOMPRESSED_PART_BYTES } from '../src/lib/ooxml-text.js';
import {
  loadSearchIndexConfig,
  validateSearchIndexConfig,
} from '../src/lib/search-index-config.js';
import { EmbedClient } from '../src/lib/search-index-embed.js';
import {
  isNodeSqliteAvailable,
  openSearchIndexStore,
  type SearchIndexStore,
} from '../src/lib/search-index-store.js';
import {
  buildFtsQuery,
  buildSnippet,
  fuseRankings,
  initSearchIndex,
  refreshSearchIndex,
  resetSearchIndexForTests,
  searchIndexStatus,
  semanticSearch,
  toGraphDeltaPath,
  type GraphAccess,
  type SearchIndexDeps,
} from '../src/lib/search-index.js';
import { chunkText } from '../src/lib/text-chunk.js';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const REL_TYPE_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
// The config clamps OPENAI_EMBED_DIMS to 64..3072, so the fake embeddings use
// the smallest allowed width instead of a toy width the store would reject.
const TEST_DIMS = 64;

const sqliteAvailable = await isNodeSqliteAvailable();

// ---------------------------------------------------------------------------
// OOXML fixtures
// ---------------------------------------------------------------------------

async function buildXlsxFixture(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Budget" sheetId="1" r:id="rId1"/>' +
      '<sheet name="Tomt" sheetId="2" r:id="rId2"/></sheets></workbook>'
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_TYPE_BASE}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL_TYPE_BASE}/worksheet" Target="worksheets/sheet2.xml"/>` +
      '</Relationships>'
  );
  zip.file(
    'xl/sharedStrings.xml',
    '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<si><t>Kvartals</t><t>regnskab</t></si><si><t>Omsætning</t></si></sst>'
  );
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData>' +
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1234.5</v></c></row>' +
      '<row r="2"><c r="A2" t="s"><v>1</v></c>' +
      '<c r="C2" t="inlineStr"><is><t>Noter her</t></is></c></row>' +
      '<row r="3"/>' +
      '</sheetData></worksheet>'
  );
  zip.file(
    'xl/worksheets/sheet2.xml',
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1"/></sheetData></worksheet>'
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function buildPptxFixture(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    'ppt/presentation.xml',
    '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId1"/></p:sldIdLst>' +
      '</p:presentation>'
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_TYPE_BASE}/slide" Target="slides/slide1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL_TYPE_BASE}/slide" Target="slides/slide2.xml"/>` +
      '</Relationships>'
  );
  zip.file(
    'ppt/slides/slide1.xml',
    '<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>' +
      '<p:sp><p:txBody><a:p><a:r><a:t>Strategi 2027</a:t></a:r></a:p>' +
      '<a:p><a:r><a:t>Vækst i Norden</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld></p:sld>'
  );
  zip.file(
    'ppt/slides/slide2.xml',
    '<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>' +
      '<p:sp><p:txBody><a:p><a:r><a:t>Bilag: tal</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld></p:sld>'
  );
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_TYPE_BASE}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>` +
      '</Relationships>'
  );
  zip.file(
    'ppt/notesSlides/notesSlide1.xml',
    '<?xml version="1.0"?><p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Husk talegenkendelse</a:t></a:r></a:p>' +
      '</p:txBody></p:sp></p:spTree></p:cSld></p:notes>'
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}

// ---------------------------------------------------------------------------
// Test doubles for the sync engine
// ---------------------------------------------------------------------------

interface FakeFile {
  id: string;
  name: string;
  mime: string;
  bytes: Buffer;
  quickxor?: string;
}

function streamOf(bytes: Buffer): ReadableStream<Uint8Array> {
  return new Response(bytes).body as ReadableStream<Uint8Array>;
}

function makeFakeGraph(files: Map<string, FakeFile>, deltaItems: () => unknown[]): GraphAccess {
  return {
    makeRequest: vi.fn(async (url: string) => {
      if (url === '/drives/drv-accounting/root') {
        return { id: 'root-1', name: 'Documents' };
      }
      if (url.includes('/root/delta')) {
        return {
          value: deltaItems(),
          '@odata.deltaLink':
            'https://graph.microsoft.com/v1.0/drives/drv-accounting/root/delta?token=token-1',
        };
      }
      throw new Error(`unexpected fake URL: ${url}`);
    }) as unknown as GraphAccess['makeRequest'],
    downloadStream: vi.fn(async (endpoint: string) => {
      const itemId = endpoint.split('/items/')[1]?.split('/')[0] ?? '';
      const file = files.get(itemId);
      if (!file) throw new Error(`no fake bytes for item ${itemId}`);
      return {
        body: streamOf(file.bytes),
        contentType: file.mime,
        contentLength: file.bytes.length,
        contentDisposition: null,
      };
    }) as unknown as GraphAccess['downloadStream'],
  };
}

function oneHot(index: number): number[] {
  const vector = new Array<number>(TEST_DIMS).fill(0);
  vector[index] = 1;
  return vector;
}

function embeddingFor(text: string): number[] {
  const lower = text.toLowerCase();
  if (lower.includes('alpha')) return oneHot(0);
  if (lower.includes('gamma')) return diagonalVector();
  return oneHot(1);
}

/** Unit vector at 45° to oneHot(0): cosine 0.7071, euclidean distance 0.7654. */
function diagonalVector(): number[] {
  const vector = new Array<number>(TEST_DIMS).fill(0);
  vector[0] = Math.SQRT1_2;
  vector[1] = Math.SQRT1_2;
  return vector;
}

// Structural stand-in for RequestInit: eslint's node globals do not define the
// DOM type, and the fake fetch only reads the JSON body.
type FetchInit = { body?: unknown };

function makeFakeEmbed(): EmbedClient {
  const fakeFetch = vi.fn(async (_url: unknown, init?: FetchInit) => {
    const body = JSON.parse(String(init?.body)) as { input: string[] };
    const data = body.input.map((text, index) => ({
      index,
      embedding: embeddingFor(text),
    }));
    return new Response(JSON.stringify({ data, usage: { prompt_tokens: body.input.length * 2 } }), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    });
  });
  return new EmbedClient({
    apiKey: 'test-key',
    model: 'test-embed',
    dims: TEST_DIMS,
    baseUrl: 'http://embeddings.test',
    fetchImpl: fakeFetch as unknown as typeof fetch,
  });
}

function deltaFile(id: string, name: string, mime: string, extra: Record<string, unknown> = {}) {
  return {
    id,
    name,
    size: 32,
    webUrl: `https://contoso.sharepoint.com/sites/team/Documents/${name}`,
    lastModifiedDateTime: '2026-10-11T10:00:00Z',
    parentReference: { id: 'folder-general', driveId: 'drv-accounting' },
    file: { mimeType: mime, hashes: { quickXorHash: `hash-${id}` } },
    ...extra,
  };
}

// ---------------------------------------------------------------------------

describe('document kinds and OOXML extraction', () => {
  it('detects spreadsheet and presentation kinds by MIME type and extension', () => {
    expect(detectDocumentKind(XLSX_MIME, 'budget.xlsx')).toBe('xlsx');
    expect(detectDocumentKind('application/vnd.ms-excel.sheet.macroEnabled.12', null)).toBe('xlsx');
    expect(detectDocumentKind(PPTX_MIME, 'deck.pptx')).toBe('pptx');
    expect(detectDocumentKind(null, 'notes.pptm')).toBe('pptx');
    expect(detectDocumentKind('application/vnd.ms-excel', 'legacy.xls')).toBeUndefined();
  });

  it('extracts xlsx sheet text, shared and inline strings, skipping empty rows and sheets', async () => {
    const text = await extractDocumentText(await buildXlsxFixture(), 'xlsx');
    expect(text).toContain('[Sheet: Budget]');
    expect(text).toContain('Kvartalsregnskab');
    expect(text).toContain('1234.5');
    expect(text).toContain('Omsætning');
    expect(text).toContain('Noter her');
    expect(text).not.toContain('[Sheet: Tomt]');
  });

  it('extracts pptx slide text in presentation order with speaker notes', async () => {
    const text = await extractDocumentText(await buildPptxFixture(), 'pptx');
    expect(text).toContain('[Slide 1]');
    expect(text).toContain('Bilag: tal');
    expect(text).toContain('[Slide 2]');
    expect(text).toContain('Strategi 2027');
    expect(text).toContain('Vækst i Norden');
    expect(text).toContain('[Notes]');
    expect(text).toContain('Husk talegenkendelse');
    // rId2 points at slide2.xml and is listed first: presentation order wins.
    expect(text.indexOf('Bilag: tal')).toBeLessThan(text.indexOf('Strategi 2027'));
  });

  it('reports password-protected or corrupt OOXML packages clearly', async () => {
    await expect(extractDocumentText(Buffer.from('not a zip at all'), 'xlsx')).rejects.toThrow(
      /spreadsheet could not be read/i
    );
    await expect(extractDocumentText(Buffer.from('not a zip at all'), 'pptx')).rejects.toThrow(
      /presentation could not be read/i
    );
  });

  it('rejects a package part that would expand past the safe extraction bound', async () => {
    const zip = new JSZip();
    zip.file(
      'ppt/slides/slide1.xml',
      `<?xml version="1.0"?><p:sld><a:t>${'x'.repeat(MAX_UNCOMPRESSED_PART_BYTES + 1)}</a:t></p:sld>`
    );
    const bytes = await zip.generateAsync({ type: 'nodebuffer' });
    await expect(extractDocumentText(bytes, 'pptx')).rejects.toThrow(/too large/i);
  });
});

describe('chunkText', () => {
  it('returns one chunk for short text', () => {
    expect(chunkText('kort tekst')).toEqual(['kort tekst']);
  });

  it('splits long text into bounded, overlapping chunks', () => {
    const text = Array.from({ length: 40 }, (_, i) => `Linje ${i} med noget indhold her.`).join(
      '\n'
    );
    const chunks = chunkText(text, { chunkChars: 300, overlapChars: 50 });
    expect(chunks.length).toBeGreaterThan(1);
    for (const chunk of chunks) {
      expect(chunk.length).toBeLessThanOrEqual(300);
      expect(chunk.trim()).toBe(chunk);
    }
    // Consecutive chunks overlap: the tail of one chunk reappears in the next.
    const tail = chunks[0].slice(-40);
    const overlapZone = chunks[1].slice(0, 120);
    expect(overlapZone).toContain(tail.slice(0, 20));
  });

  it('prefers a paragraph boundary inside the look-back window', () => {
    const text = `${'a'.repeat(700)}\n\n${'b'.repeat(700)}`;
    const chunks = chunkText(text, { chunkChars: 900, overlapChars: 100 });
    expect(chunks[0].endsWith('a')).toBe(true);
    expect(chunks[0]).not.toContain('b');
  });

  it('honours maxChunks', () => {
    const text = 'x'.repeat(10_000);
    expect(chunkText(text, { chunkChars: 100, overlapChars: 0, maxChunks: 3 })).toHaveLength(3);
  });
});

describe('ranking helpers', () => {
  it('fuses both rankings with reciprocal rank fusion and rewards dual matches', () => {
    const fused = fuseRankings([10, 20, 30], [30, 40], true);
    // rank 0 in one list is 1/61; rank 2 in the other is 1/63.
    expect(fused.get(10)).toBeCloseTo(1 / 61, 6);
    expect(fused.get(30)).toBeCloseTo(1 / 63 + 1 / 61, 6);
    expect(fused.get(30)).toBeGreaterThan(fused.get(10));
    expect(fused.get(40)).toBeCloseTo(1 / 62, 6);
  });

  it('falls back to the only available ranking when hybrid is off', () => {
    const vectorOnly = fuseRankings([1, 2], [], false);
    expect([...vectorOnly.keys()]).toEqual([1, 2]);
    const lexicalFallback = fuseRankings([], [3], false);
    expect([...lexicalFallback.keys()]).toEqual([3]);
  });

  it('builds a quoted OR query from the meaningful tokens', () => {
    expect(buildFtsQuery('Kvartalsregnskab 2026')).toBe('"kvartalsregnskab" OR "2026"');
    expect(buildFtsQuery('a b')).toBeUndefined();
    expect(buildFtsQuery('')).toBeUndefined();
  });

  it('builds snippets around the first query match', () => {
    const text = `${'fyld '.repeat(200)}kvartalsregnskab og mere tekst`;
    const snippet = buildSnippet(text, 'kvartalsregnskab', 100);
    expect(snippet).toContain('kvartalsregnskab');
    expect(snippet.length).toBeLessThanOrEqual(102);
    expect(snippet.startsWith('…')).toBe(true);
  });

  it('normalises Graph delta links to version-relative paths', () => {
    expect(toGraphDeltaPath('https://graph.microsoft.com/v1.0/drives/d/root/delta?token=abc')).toBe(
      '/drives/d/root/delta?token=abc'
    );
    expect(toGraphDeltaPath('https://graph.microsoft.com/beta/drives/d/root/delta?token=abc')).toBe(
      '/drives/d/root/delta?token=abc'
    );
    expect(toGraphDeltaPath('/drives/d/root/delta?token=abc')).toBe(
      '/drives/d/root/delta?token=abc'
    );
  });
});

describe('embed client', () => {
  it('parses embeddings in index order and counts prompt tokens', async () => {
    const fakeFetch = vi.fn(async () =>
      Response.json({
        data: [
          { index: 1, embedding: [0, 1] },
          { index: 0, embedding: [1, 0] },
        ],
        usage: { prompt_tokens: 7 },
      })
    );
    const client = new EmbedClient({
      apiKey: 'k',
      dims: 2,
      fetchImpl: fakeFetch as unknown as typeof fetch,
    });
    const result = await client.embedBatch(['first', 'second']);
    expect([...result.vectors[0]]).toEqual([1, 0]);
    expect([...result.vectors[1]]).toEqual([0, 1]);
    expect(result.promptTokens).toBe(7);
  });

  it('retries a 500 response and succeeds on the second attempt', async () => {
    const fakeFetch = vi
      .fn()
      .mockResolvedValueOnce(new Response('boom', { status: 500 }))
      .mockResolvedValueOnce(
        Response.json({ data: [{ index: 0, embedding: [1, 0] }], usage: { prompt_tokens: 1 } })
      );
    const sleep = vi.fn(async () => undefined);
    const client = new EmbedClient({
      apiKey: 'k',
      dims: 2,
      fetchImpl: fakeFetch as unknown as typeof fetch,
      sleepImpl: sleep,
    });
    const result = await client.embedBatch(['x']);
    expect(fakeFetch).toHaveBeenCalledTimes(2);
    expect(sleep).toHaveBeenCalledTimes(1);
    expect(result.vectors).toHaveLength(1);
  });

  it('fails fast on a non-retriable error and never echoes the input', async () => {
    const fakeFetch = vi.fn(
      async () => new Response('{"error":{"message":"bad key"}}', { status: 401 })
    );
    const client = new EmbedClient({
      apiKey: 'k',
      dims: 2,
      fetchImpl: fakeFetch as unknown as typeof fetch,
      sleepImpl: async () => undefined,
    });
    await expect(client.embedBatch(['hemmelig tekst'])).rejects.toThrow(/HTTP 401/);
    expect(fakeFetch).toHaveBeenCalledTimes(1);
    await expect(client.embedBatch(['hemmelig tekst'])).rejects.not.toThrow(/hemmelig/);
  });

  it('rejects malformed responses', async () => {
    const fakeFetch = vi.fn(async () => Response.json({ data: [{ index: 0, embedding: [1] }] }));
    const client = new EmbedClient({
      apiKey: 'k',
      dims: 2,
      fetchImpl: fakeFetch as unknown as typeof fetch,
    });
    await expect(client.embedBatch(['x'])).rejects.toThrow(/malformed/i);
  });
});

describe('search index config', () => {
  it('defaults to disabled with the standard db path and model', () => {
    const config = loadSearchIndexConfig({});
    expect(config.enabled).toBe(false);
    expect(config.dbPath).toBe('/data/search-index.sqlite');
    expect(config.embedDims).toBe(1024);
    expect(config.syncMinutes).toBe(15);
  });

  it('requires drives and an API key when enabled', () => {
    const noDrives = loadSearchIndexConfig({ MS365_MCP_INDEX_ENABLED: '1' });
    expect(validateSearchIndexConfig(noDrives)).toMatch(/MS365_MCP_INDEX_DRIVES/);
    const noKey = loadSearchIndexConfig({
      MS365_MCP_INDEX_ENABLED: '1',
      MS365_MCP_INDEX_DRIVES: 'drv-1',
    });
    expect(validateSearchIndexConfig(noKey)).toMatch(/OPENAI_API_KEY/);
    const valid = loadSearchIndexConfig({
      MS365_MCP_INDEX_ENABLED: '1',
      MS365_MCP_INDEX_DRIVES: 'drv-1, drv-2',
      OPENAI_API_KEY: 'k',
    });
    expect(validateSearchIndexConfig(valid)).toBeUndefined();
    expect(valid.driveIds).toEqual(['drv-1', 'drv-2']);
  });
});

describe.skipIf(!sqliteAvailable)('search index store', () => {
  let directory: string;
  let store: SearchIndexStore;

  const record = {
    itemId: 'item-1',
    driveId: 'drv-1',
    name: 'budget.xlsx',
    path: '/Regnskab/budget.xlsx',
    webUrl: 'https://contoso.sharepoint.com/sites/team/Documents/budget.xlsx',
    mime: XLSX_MIME,
    size: 42,
    lastModified: '2026-10-11T10:00:00Z',
  };

  beforeEach(async () => {
    directory = mkdtempSync(join(tmpdir(), 'search-index-store-'));
    store = await openSearchIndexStore({
      dbPath: join(directory, 'index.sqlite'),
      dims: 4,
      embedModel: 'test-embed',
    });
  });

  afterEach(() => {
    store.close();
    rmSync(directory, { recursive: true, force: true });
  });

  it('stores chunks with FTS and vector rows and retrieves them both ways', () => {
    store.replaceItemChunks(record, [
      {
        text: 'Kvartalsregnskab for 2026 med omsætning',
        embedding: new Float32Array([1, 0, 0, 0]),
      },
      { text: 'Bilagsoversigt og noter', embedding: new Float32Array([0, 1, 0, 0]) },
    ]);
    expect(store.searchableChunkCount()).toBe(2);
    expect(store.counts().documents).toBe(1);

    const lexical = store.lexicalSearch(buildFtsQuery('kvartalsregnskab')!, 10);
    expect(lexical).toHaveLength(1);

    if (store.vectorEnabled) {
      const vector = store.vectorSearch(new Float32Array([1, 0, 0, 0]), 2);
      expect(vector).toHaveLength(2);
      expect(vector[0].distance).toBeLessThanOrEqual(vector[1].distance);
    }

    const rows = store.chunksWithItems(lexical.map((row) => row.rowid));
    expect(rows[0].name).toBe('budget.xlsx');
    expect(rows[0].path).toBe('/Regnskab/budget.xlsx');
  });

  it('replaces chunks atomically and cascades deletes', () => {
    store.replaceItemChunks(record, [
      { text: 'gammel tekst', embedding: new Float32Array([1, 0, 0, 0]) },
    ]);
    expect(store.searchableChunkCount()).toBe(1);
    store.replaceItemChunks(record, [
      { text: 'ny tekst', embedding: new Float32Array([1, 0, 0, 0]) },
      { text: 'mere tekst', embedding: new Float32Array([0, 1, 0, 0]) },
    ]);
    expect(store.searchableChunkCount()).toBe(2);
    expect(store.lexicalSearch(buildFtsQuery('gammel')!, 10)).toHaveLength(0);

    store.deleteItemWithChunks('item-1');
    expect(store.searchableChunkCount()).toBe(0);
    expect(store.counts().documents).toBe(0);
    expect(store.getItem('item-1')).toBeUndefined();
  });

  it('wipes derived data when the embedding model or dimensions change', async () => {
    store.replaceItemChunks(record, [{ text: 'noget', embedding: new Float32Array([1, 0, 0, 0]) }]);
    expect(store.counts().documents).toBe(1);
    store.close();

    store = await openSearchIndexStore({
      dbPath: join(directory, 'index.sqlite'),
      dims: 8,
      embedModel: 'test-embed',
    });
    expect(store.counts().documents).toBe(0);
    expect(store.searchableChunkCount()).toBe(0);
  });

  it('tracks drive state, delta tokens and item states', () => {
    expect(store.getDriveState('drv-1')).toBeUndefined();
    store.recordDriveRun('drv-1', 'ok');
    expect(store.getDriveState('drv-1')?.lastStatus).toBe('ok');
    store.saveDriveDelta(
      'drv-1',
      'https://graph.microsoft.com/v1.0/drives/drv-1/root/delta?token=1'
    );
    expect(store.getDriveState('drv-1')?.deltaLink).toContain('token=1');
    store.upsertItem({ ...record, state: 'error', skipReason: 'boom' });
    const errored = store.itemsByState('drv-1', 'error', 10);
    expect(errored).toHaveLength(1);
    expect(errored[0].quickxor).toBeUndefined();
    store.resetDrive('drv-1');
    expect(store.getDriveState('drv-1')).toBeUndefined();
    expect(store.getItem('item-1')).toBeUndefined();
  });
});

describe.skipIf(!sqliteAvailable)('index runtime with fake Graph and embeddings', () => {
  let directory: string;
  const driveId = 'drv-accounting';
  const files = new Map<string, FakeFile>();
  let graph: GraphAccess;
  let deps: SearchIndexDeps;
  let env: Record<string, string | undefined>;
  let extraItems: unknown[];

  const folderItem = {
    id: 'folder-general',
    name: 'General',
    folder: {},
    // Graph always sends a folder's lastModifiedDateTime; it changes whenever a
    // child changes, so folder records must not compare it.
    lastModifiedDateTime: '2026-10-10T10:00:00Z',
    parentReference: { id: 'root-1', driveId },
  };

  function defaultItems(): unknown[] {
    return [
      folderItem,
      deltaFile('item-notes', 'notes.txt', 'text/plain'),
      deltaFile('item-legacy', 'data.xls', 'application/vnd.ms-excel'),
      { id: 'item-gone', deleted: {} },
      ...extraItems,
    ];
  }

  beforeEach(() => {
    directory = mkdtempSync(join(tmpdir(), 'search-index-runtime-'));
    files.clear();
    extraItems = [];
    files.set('item-notes', {
      id: 'item-notes',
      name: 'notes.txt',
      mime: 'text/plain',
      bytes: Buffer.from('Alpha kvartalsregnskab og noter til revisionen', 'utf8'),
    });
    graph = makeFakeGraph(files, defaultItems);
    deps = { graph, embed: makeFakeEmbed() };
    env = {
      MS365_MCP_INDEX_ENABLED: '1',
      MS365_MCP_INDEX_DRIVES: driveId,
      MS365_MCP_INDEX_DB: join(directory, 'index.sqlite'),
      MS365_MCP_INDEX_SYNC_MINUTES: '0',
      OPENAI_EMBED_MODEL: 'test-embed',
      OPENAI_EMBED_DIMS: String(TEST_DIMS),
      OPENAI_API_KEY: 'test-key',
    };
  });

  afterEach(() => {
    resetSearchIndexForTests();
    rmSync(directory, { recursive: true, force: true });
  });

  it('indexes a delta page, skips unsupported types and deletes removed items', async () => {
    await initSearchIndex(deps, env);
    const result = await refreshSearchIndex(deps, {});

    const summary = result.drives[0];
    expect(summary.status).toBe('ok');
    expect(summary.folders).toBe(1);
    expect(summary.indexed).toBe(1);
    expect(summary.skipped).toBe(1);
    // The deletion refers to an item that was never indexed: nothing to remove.
    expect(summary.deleted).toBe(0);
    expect(summary.unchanged).toBe(1);
    expect(summary.moreRemaining).toBe(false);

    const status = searchIndexStatus(env);
    expect(status.ready).toBe(true);
    expect(status.enabled).toBe(true);
    expect(status.documents).toBe(1);
    expect(status.chunks).toBeGreaterThan(0);
    expect(status.driveStates[0]?.hasDeltaToken).toBe(true);
    expect(status.dbBytes).toBeGreaterThan(0);

    const search = await semanticSearch({ query: 'Alpha', topK: 5 });
    expect(search.hits.length).toBeGreaterThanOrEqual(1);
    expect(search.hits[0].name).toBe('notes.txt');
    expect(search.hits[0].webUrl).toContain('notes.txt');
    expect(search.hits[0].path).toBe('/Documents/General/notes.txt');
    expect(search.hits[0].snippet).toContain('Alpha');

    // A second run with the same page reports the folder, the notes file, the
    // unsupported file and the phantom deletion as unchanged and reuses the token.
    const second = await refreshSearchIndex(deps, {});
    expect(second.drives[0].unchanged).toBe(4);
    expect(second.drives[0].indexed).toBe(0);
    expect(second.drives[0].folders).toBe(0);
    expect(second.drives[0].skipped).toBe(0);
    expect(second.drives[0].deleted).toBe(0);
    const deltaCalls = (graph.makeRequest as unknown as { mock: { calls: unknown[][] } }).mock.calls
      .map((call) => String(call[0]))
      .filter((url) => url.includes('token-1'));
    expect(deltaCalls.length).toBe(1);
  });

  it('resumes a budget-limited cold run on the next run instead of restarting', async () => {
    env.MS365_MCP_INDEX_MAX_ITEMS_PER_RUN = '2';
    for (let i = 1; i <= 40; i += 1) {
      const id = `bulk-${String(i).padStart(2, '0')}`;
      files.set(id, {
        id,
        name: `${id}.txt`,
        mime: 'text/plain',
        bytes: Buffer.from(`Alpha indhold nummer ${i}`, 'utf8'),
      });
      extraItems.push(deltaFile(id, `${id}.txt`, 'text/plain'));
    }
    await initSearchIndex(deps, env);

    // Cold budget: 2 * COLD_RUN_MULTIPLIER = 20 work items. The folder, the
    // notes file and the unsupported legacy file count as work; the phantom
    // deletion and the re-walked unchanged prefix do not.
    const first = await refreshSearchIndex(deps, {});
    expect(first.drives[0].moreRemaining).toBe(true);
    expect(first.drives[0].seen).toBe(21);
    expect(first.drives[0].folders).toBe(1);
    expect(first.drives[0].indexed).toBe(18);
    expect(first.drives[0].skipped).toBe(1);
    expect(first.drives[0].unchanged).toBe(1);

    // The already-done prefix is free, so the second run advances into the
    // rest of the feed instead of re-spending its budget on the same items.
    const second = await refreshSearchIndex(deps, {});
    expect(second.drives[0].moreRemaining).toBe(true);
    expect(second.drives[0].indexed).toBe(20);
    expect(second.drives[0].folders).toBe(0);
    expect(second.drives[0].skipped).toBe(0);

    const third = await refreshSearchIndex(deps, {});
    expect(third.drives[0].moreRemaining).toBe(false);
    expect(third.drives[0].indexed).toBe(3);

    const status = searchIndexStatus(env);
    expect(status.documents).toBe(41);
    expect(status.driveStates[0]?.hasDeltaToken).toBe(true);

    // The feed now fits the warm delta budget and reports only unchanged items.
    const fourth = await refreshSearchIndex(deps, {});
    expect(fourth.drives[0].indexed).toBe(0);
    expect(fourth.drives[0].unchanged).toBe(44);
  });

  it('converges on a feed of folders, unsupported files and a persistent error', async () => {
    // Cold budget is 2 * 10 work items. The folders plus unsupported files alone
    // reach that budget, so without free re-walks this feed would never finish.
    env.MS365_MCP_INDEX_MAX_ITEMS_PER_RUN = '2';
    for (let i = 1; i <= 12; i += 1) {
      extraItems.push({
        id: `folder-${i}`,
        name: `Folder ${i}`,
        folder: {},
        lastModifiedDateTime: '2026-10-10T10:00:00Z',
        parentReference: { id: 'root-1', driveId },
      });
    }
    for (let i = 1; i <= 6; i += 1) {
      extraItems.push(deltaFile(`legacy-${i}`, `legacy-${i}.xls`, 'application/vnd.ms-excel'));
    }
    for (let i = 1; i <= 10; i += 1) {
      const id = `doc-${i}`;
      files.set(id, {
        id,
        name: `${id}.txt`,
        mime: 'text/plain',
        bytes: Buffer.from(`Alpha tekst nummer ${i}`, 'utf8'),
      });
      extraItems.push(deltaFile(id, `${id}.txt`, 'text/plain'));
    }
    // No bytes for this one: the download fails on every run.
    extraItems.push(deltaFile('doc-broken', 'broken.txt', 'text/plain'));
    await initSearchIndex(deps, env);

    // Cold budget is 3 * 10 work items. Folders and unsupported files must stop
    // counting once stored, so every run spends its budget on new documents.
    let runs = 0;
    let summary = await refreshSearchIndex(deps, {});
    runs += 1;
    while (summary.drives[0].moreRemaining && runs < 20) {
      summary = await refreshSearchIndex(deps, {});
      runs += 1;
    }
    expect(summary.drives[0].moreRemaining).toBe(false);
    expect(runs).toBeLessThanOrEqual(6);
    expect(searchIndexStatus(env).documents).toBe(11);
    expect(searchIndexStatus(env).driveStates[0]?.hasDeltaToken).toBe(true);

    // A warm run walks the whole feed for free: folders, unsupported files and
    // the unchanged error are all unchanged, and nothing is downloaded again.
    const warm = await refreshSearchIndex(deps, {});
    expect(warm.drives[0].moreRemaining).toBe(false);
    expect(warm.drives[0].folders).toBe(0);
    expect(warm.drives[0].skipped).toBe(0);
    expect(warm.drives[0].indexed).toBe(0);
    expect(warm.drives[0].errors).toBe(1);
    expect(warm.drives[0].unchanged).toBeGreaterThanOrEqual(31);
  });

  it('counts without writing on dryRun and walks past the item budget', async () => {
    // The item budget bounds real work only; the survey must still cover the
    // whole drive in a single call.
    env.MS365_MCP_INDEX_MAX_ITEMS_PER_RUN = '1';
    extraItems.push(deltaFile('item-budget', 'budget.xlsx', XLSX_MIME));
    await initSearchIndex(deps, env);
    const result = await refreshSearchIndex(deps, { dryRun: true, full: true });

    const summary = result.drives[0];
    expect(result.mode).toBe('dry-run');
    expect(summary.moreRemaining).toBe(false);
    expect(summary.folders).toBe(1);
    expect(summary.survey?.byKind.text?.files).toBe(1);
    expect(summary.survey?.byKind.xlsx?.files).toBe(1);
    expect(summary.survey?.unsupported.files).toBe(1);
    const status = searchIndexStatus(env);
    expect(status.documents).toBe(0);
    expect(status.chunks).toBe(0);
    expect(status.driveStates.length).toBe(0);
  });

  it('exposes cosine similarity per hit and filters on minSimilarity', async () => {
    extraItems.push(
      deltaFile('item-beta', 'beta.txt', 'text/plain'),
      deltaFile('item-gamma', 'gamma.txt', 'text/plain')
    );
    files.set('item-beta', {
      id: 'item-beta',
      name: 'beta.txt',
      mime: 'text/plain',
      bytes: Buffer.from('Beta uden nøgleordet', 'utf8'),
    });
    files.set('item-gamma', {
      id: 'item-gamma',
      name: 'gamma.txt',
      mime: 'text/plain',
      bytes: Buffer.from('Gamma indhold i dokumentet', 'utf8'),
    });
    await initSearchIndex(deps, env);
    await refreshSearchIndex(deps, {});

    // The fake embeddings are unit vectors at 0°, 45° and 90° to the query, so
    // the reported similarity must recover 1, cos 45° and 0 from the distance.
    const all = await semanticSearch({ query: 'Alpha', topK: 5 });
    const byName = new Map(all.hits.map((hit) => [hit.name, hit]));
    expect(byName.get('notes.txt')?.similarity).toBeCloseTo(1, 3);
    expect(byName.get('gamma.txt')?.similarity).toBeCloseTo(Math.SQRT1_2, 3);
    expect(byName.get('beta.txt')?.similarity).toBeCloseTo(0, 3);

    const filtered = await semanticSearch({ query: 'Alpha', topK: 5, minSimilarity: 0.6 });
    expect(filtered.hits.map((hit) => hit.name)).toEqual(['notes.txt', 'gamma.txt']);
    expect(filtered.notes.join(' ')).toMatch(/below minSimilarity/);

    const strict = await semanticSearch({ query: 'Alpha', topK: 5, minSimilarity: 0.9 });
    expect(strict.hits.map((hit) => hit.name)).toEqual(['notes.txt']);
  });

  it('records per-item errors and retries them on the next run', async () => {
    const working = graph.downloadStream;
    const failing = vi.fn(async () => {
      throw new Error('Download failed');
    }) as unknown as GraphAccess['downloadStream'];
    const failingGraph: GraphAccess = { ...graph, downloadStream: failing };
    await initSearchIndex({ graph: failingGraph, embed: deps.embed }, env);
    const first = await refreshSearchIndex({ graph: failingGraph, embed: deps.embed }, {});
    expect(first.drives[0].errors).toBe(1);
    expect(searchIndexStatus(env).byDrive[0]?.errors).toBe(1);

    // The retry pass runs before the delta feed, so the item is healed without
    // waiting for it to change again.
    const healedGraph: GraphAccess = { ...graph, downloadStream: working };
    const second = await refreshSearchIndex({ graph: healedGraph, embed: deps.embed }, {});
    expect(second.drives[0].indexed).toBe(1);
    expect(searchIndexStatus(env).byDrive[0]?.errors).toBe(0);
  });

  it('rejects unknown drives and uninitialized use', async () => {
    resetSearchIndexForTests();
    expect(searchIndexStatus({}).ready).toBe(false);
    await expect(semanticSearch({ query: 'x' })).rejects.toThrow(/search index/i);
    await initSearchIndex(deps, env);
    await expect(refreshSearchIndex(deps, { driveId: 'other-drive' })).rejects.toThrow(
      /MS365_MCP_INDEX_DRIVES/
    );
    expect(
      searchIndexStatus({ MS365_MCP_INDEX_ENABLED: '1', MS365_MCP_INDEX_DRIVES: 'x' }).ready
    ).toBe(true);
  });

  it('returns a lexical-only note when the embedding call fails but FTS still matches', async () => {
    const state = { fail: false };
    const fetchImpl = vi.fn(async (_url: unknown, init?: FetchInit) => {
      if (state.fail) return new Response('upstream boom', { status: 500 });
      const body = JSON.parse(String(init?.body)) as { input: string[] };
      const data = body.input.map((text, index) => ({ index, embedding: embeddingFor(text) }));
      return Response.json({ data, usage: { prompt_tokens: 1 } });
    });
    const embed = new EmbedClient({
      apiKey: 'test-key',
      model: 'test-embed',
      dims: TEST_DIMS,
      baseUrl: 'http://embeddings.test',
      fetchImpl: fetchImpl as unknown as typeof fetch,
      sleepImpl: async () => undefined,
      retries: 0,
    });
    await initSearchIndex({ graph, embed }, env);
    await refreshSearchIndex({ graph, embed }, {});
    // Sync embedded with the working client; only the query embedding fails now,
    // so the search must degrade to FTS instead of returning nothing.
    state.fail = true;
    const search = await semanticSearch({ query: 'Alpha' });
    expect(search.hits.length).toBeGreaterThanOrEqual(1);
    expect(search.notes.join(' ')).toMatch(/lexical/i);
  });
});
