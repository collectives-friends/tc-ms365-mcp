/**
 * Text extraction for the OOXML spreadsheet and presentation packages
 * (xlsx/xlsm, pptx/pptm).
 *
 * An OOXML file is a ZIP of XML parts. This module reads only the parts that
 * carry user text -- workbook sheets with their shared strings, and slides
 * with their speaker notes -- and flattens each into plain text. jszip and
 * @xmldom/xmldom are already in the dependency tree through mammoth, and both
 * are imported lazily like the docx/pdf parsers, so this payload is only paid
 * for when a spreadsheet or presentation is actually read.
 *
 * Resource limits are the caller's job: like document-text.ts this module
 * never reads from the network, and the callers cap the input bytes before
 * parsing. OOXML text is additionally capped at OOXML_MAX_TEXT_CHARS so a
 * pathological workbook cannot expand into unbounded text; the cap sits far
 * above the largest page any caller returns in one response.
 */

export const OOXML_MAX_TEXT_CHARS = 2_000_000;
/**
 * Default upper bound on one decompressed package part. The XML DOM for a part
 * costs many times the part's own bytes, so a part beyond this bound would risk
 * the whole process instead of one document; such packages are reported as
 * unreadable rather than parsed. Callers may raise or lower the bound (the
 * search index passes MS365_MCP_INDEX_MAX_PART_BYTES) when the deployment
 * accepts the larger memory cost.
 */
export const MAX_UNCOMPRESSED_PART_BYTES = 32 * 1024 * 1024;

const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Minimal structural types for the lazily imported dependencies, in the same
// spirit as document-text.ts's mammoth/unpdf shims: the server only relies on
// the small surface below, and the shims keep the dependencies' type graphs
// out of the build.
interface ZipEntryHandle {
  async(type: 'string'): Promise<string>;
  /**
   * The decompressed size jszip read from the zip central directory. Not part
   * of jszip's documented API, but stable across 3.x and the only way to reject
   * a part before it is decompressed into memory.
   */
  _data?: { uncompressedSize?: number };
}

interface ZipArchive {
  file(path: string): ZipEntryHandle | null;
  files: Record<string, { name: string; dir: boolean }>;
}

interface XmlNode {
  nodeType: number;
  localName?: string | null;
  childNodes: ArrayLike<XmlNode>;
  textContent?: string | null;
  getAttribute?(name: string): string | null;
  getAttributeNS?(namespace: string | null, localName: string): string | null;
}

type JSZipModule = { loadAsync(data: Buffer): Promise<ZipArchive> };
type XmldomModule = {
  DOMParser: new () => { parseFromString(xml: string, mime: string): XmlNode };
};

async function loadJSZip(): Promise<JSZipModule> {
  const imported: unknown = await import('jszip');
  const candidate = (imported as { default?: unknown }).default ?? imported;
  if (typeof (candidate as { loadAsync?: unknown }).loadAsync !== 'function') {
    throw new Error('jszip did not expose loadAsync');
  }
  return candidate as JSZipModule;
}

async function loadXmldom(): Promise<XmldomModule> {
  const imported: unknown = await import('@xmldom/xmldom');
  const candidate = (imported as { default?: unknown }).default ?? imported;
  const Parser = (candidate as { DOMParser?: unknown }).DOMParser;
  if (typeof Parser !== 'function') {
    throw new Error('@xmldom/xmldom did not expose DOMParser');
  }
  return candidate as XmldomModule;
}

async function openZip(bytes: Buffer, label: string): Promise<ZipArchive> {
  const jszip = await loadJSZip();
  try {
    return await jszip.loadAsync(bytes);
  } catch {
    throw new Error(`The ${label} could not be read; it may be password-protected or corrupt.`);
  }
}

async function readPart(
  zip: ZipArchive,
  partPath: string,
  maxPartBytes: number = MAX_UNCOMPRESSED_PART_BYTES
): Promise<string | undefined> {
  const entry = zip.file(partPath);
  if (!entry) return undefined;
  const uncompressedSize = entry._data?.uncompressedSize;
  if (typeof uncompressedSize === 'number' && uncompressedSize > maxPartBytes) {
    throw new Error('The document is too large to extract safely.');
  }
  return entry.async('string');
}

/**
 * Reject a Word package whose main document part would expand past the safe
 * part bound before mammoth parses it. An unreadable package passes through:
 * mammoth then reports the real read error.
 */
export async function assertDocxWithinPartBound(
  bytes: Buffer,
  maxPartBytes: number = MAX_UNCOMPRESSED_PART_BYTES
): Promise<void> {
  let uncompressedSize: number | undefined;
  try {
    const zip = await openZip(bytes, 'document');
    uncompressedSize = zip.file('word/document.xml')?._data?.uncompressedSize;
  } catch {
    return;
  }
  if (typeof uncompressedSize === 'number' && uncompressedSize > maxPartBytes) {
    throw new Error('The document is too large to extract safely.');
  }
}

async function parseXml(xml: string): Promise<XmlNode> {
  const xmldom = await loadXmldom();
  return new xmldom.DOMParser().parseFromString(xml, 'text/xml');
}

// -- DOM helpers (namespace-agnostic; OOXML parts name their own prefixes) --

function isElement(node: XmlNode): boolean {
  return node.nodeType === 1;
}

function elementChildren(node: XmlNode): XmlNode[] {
  const out: XmlNode[] = [];
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i];
    if (child && isElement(child)) out.push(child);
  }
  return out;
}

function descendants(node: XmlNode, name: string, out: XmlNode[] = []): XmlNode[] {
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i];
    if (!child || !isElement(child)) continue;
    if (child.localName === name) out.push(child);
    descendants(child, name, out);
  }
  return out;
}

function firstChild(node: XmlNode, name: string): XmlNode | undefined {
  return elementChildren(node).find((child) => child.localName === name);
}

/** Read a relationship-style attribute that is stored with a namespace prefix (r:id). */
function relationshipAttribute(node: XmlNode, localName: string): string | undefined {
  const value =
    node.getAttribute?.(`r:${localName}`) ??
    node.getAttributeNS?.(RELATIONSHIPS_NS, localName) ??
    null;
  return value ?? undefined;
}

interface Relationship {
  id: string;
  target: string;
  type: string;
}

async function readRelationships(
  zip: ZipArchive,
  relsPath: string,
  maxPartBytes: number = MAX_UNCOMPRESSED_PART_BYTES
): Promise<Relationship[]> {
  const xml = await readPart(zip, relsPath, maxPartBytes);
  if (!xml) return [];
  const relationships: Relationship[] = [];
  for (const rel of descendants(await parseXml(xml), 'Relationship')) {
    const id = rel.getAttribute?.('Id');
    const target = rel.getAttribute?.('Target');
    if (!id || !target) continue;
    if (rel.getAttribute?.('TargetMode') === 'External') continue;
    relationships.push({ id, target, type: rel.getAttribute?.('Type') ?? '' });
  }
  return relationships;
}

/** Resolve an OPC relationship target against the part that declares it. */
function resolvePart(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const baseDir = sourcePart.slice(0, sourcePart.lastIndexOf('/') + 1);
  const segments = `${baseDir}${target}`.split('/');
  const resolved: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === '.') continue;
    if (segment === '..') {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }
  return resolved.join('/');
}

/**
 * Append one section to the output under the global text cap. Returns false
 * when the cap has been reached (the caller then stops reading more parts).
 */
function appendSection(sections: string[], text: string, total: { value: number }): boolean {
  const separator = sections.length > 0 ? 2 : 0;
  const remaining = OOXML_MAX_TEXT_CHARS - total.value;
  if (text.length + separator > remaining) {
    if (remaining > 0) {
      sections.push(`${text.slice(0, remaining)}\n[... truncated]`);
    }
    total.value = OOXML_MAX_TEXT_CHARS;
    return false;
  }
  sections.push(text);
  total.value += text.length + separator;
  return true;
}

// -- XLSX -------------------------------------------------------------------

export async function extractXlsxText(
  bytes: Buffer,
  maxPartBytes: number = MAX_UNCOMPRESSED_PART_BYTES
): Promise<string> {
  const zip = await openZip(bytes, 'spreadsheet');
  const workbookXml = await readPart(zip, 'xl/workbook.xml', maxPartBytes);
  if (!workbookXml) {
    throw new Error('The spreadsheet package has no workbook part.');
  }
  const workbookRels = await readRelationships(zip, 'xl/_rels/workbook.xml.rels', maxPartBytes);
  const sharedStrings = await readSharedStrings(zip, maxPartBytes);

  const sections: string[] = [];
  const total = { value: 0 };
  for (const sheet of descendants(await parseXml(workbookXml), 'sheet')) {
    if (total.value >= OOXML_MAX_TEXT_CHARS) break;
    const sheetName = sheet.getAttribute?.('name') ?? 'Sheet';
    const relId = relationshipAttribute(sheet, 'id');
    const target = relId ? workbookRels.find((rel) => rel.id === relId)?.target : undefined;
    const partPath = target ? resolvePart('xl/workbook.xml', target) : undefined;
    const partXml = partPath ? await readPart(zip, partPath, maxPartBytes) : undefined;
    if (!partXml) continue;
    const lines = sheetTextLines(await parseXml(partXml), sharedStrings);
    if (lines.length === 0) continue;
    if (!appendSection(sections, `[Sheet: ${sheetName}]\n${lines.join('\n')}`, total)) break;
  }
  return sections.join('\n\n');
}

async function readSharedStrings(
  zip: ZipArchive,
  maxPartBytes: number = MAX_UNCOMPRESSED_PART_BYTES
): Promise<string[]> {
  const xml = await readPart(zip, 'xl/sharedStrings.xml', maxPartBytes);
  if (!xml) return [];
  const root = await parseXml(xml);
  return descendants(root, 'si').map((item) =>
    descendants(item, 't')
      .map((t) => t.textContent ?? '')
      .join('')
  );
}

const MAX_XLSX_COLUMNS = 256;

function sheetTextLines(worksheet: XmlNode, sharedStrings: string[]): string[] {
  // `worksheet` is the parsed document node, so <sheetData> sits one level
  // down under the <worksheet> root: search descendants, not direct children.
  const sheetData = descendants(worksheet, 'sheetData')[0];
  if (!sheetData) return [];
  const lines: string[] = [];
  for (const row of elementChildren(sheetData)) {
    if (row.localName !== 'row') continue;
    const values: Array<string | undefined> = [];
    let nextColumn = 0;
    for (const cell of elementChildren(row)) {
      if (cell.localName !== 'c') continue;
      const column = columnIndex(cell.getAttribute?.('r')) ?? nextColumn;
      nextColumn = column + 1;
      if (column >= MAX_XLSX_COLUMNS) continue;
      const text = cellText(cell, sharedStrings);
      if (text) values[column] = text;
    }
    const lastFilled = lastFilledIndex(values);
    if (lastFilled < 0) continue;
    lines.push(
      values
        .slice(0, lastFilled + 1)
        .map((value) => value ?? '')
        .join('\t')
    );
  }
  return lines;
}

function lastFilledIndex(values: Array<string | undefined>): number {
  for (let i = values.length - 1; i >= 0; i--) {
    const value = values[i];
    if (value !== undefined && value !== '') return i;
  }
  return -1;
}

/** Column index from an A1-style cell reference ("B3" -> 1); undefined when absent. */
function columnIndex(reference: string | null | undefined): number | undefined {
  const match = reference ? /^([A-Za-z]+)\d+$/.exec(reference) : null;
  if (!match) return undefined;
  let index = 0;
  for (const char of match[1].toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index - 1;
}

function cellText(cell: XmlNode, sharedStrings: string[]): string {
  const type = cell.getAttribute?.('t') ?? '';
  if (type === 's') {
    const index = Number.parseInt(firstChild(cell, 'v')?.textContent ?? '', 10);
    return Number.isInteger(index) ? (sharedStrings[index] ?? '') : '';
  }
  if (type === 'inlineStr') {
    const inline = firstChild(cell, 'is');
    return inline
      ? descendants(inline, 't')
          .map((t) => t.textContent ?? '')
          .join('')
      : '';
  }
  // Formulas contribute only their cached value; t="str"/"e"/"d" keep the raw
  // <v> text and booleans get a readable rendering.
  const value = firstChild(cell, 'v')?.textContent ?? '';
  if (type === 'b') return value === '1' ? 'TRUE' : value === '0' ? 'FALSE' : value;
  return value;
}

// -- PPTX -------------------------------------------------------------------

export async function extractPptxText(
  bytes: Buffer,
  maxPartBytes: number = MAX_UNCOMPRESSED_PART_BYTES
): Promise<string> {
  const zip = await openZip(bytes, 'presentation');
  const presentationXml = await readPart(zip, 'ppt/presentation.xml', maxPartBytes);
  const presentationRels = await readRelationships(
    zip,
    'ppt/_rels/presentation.xml.rels',
    maxPartBytes
  );

  let slideParts: string[] = [];
  if (presentationXml) {
    for (const slideId of descendants(await parseXml(presentationXml), 'sldId')) {
      const relId = relationshipAttribute(slideId, 'id');
      const target = relId ? presentationRels.find((rel) => rel.id === relId)?.target : undefined;
      if (target) slideParts.push(resolvePart('ppt/presentation.xml', target));
    }
  }
  if (slideParts.length === 0) {
    // Presentation rels are missing or empty: fall back to the slide parts in
    // numeric order, which is the order PowerPoint writes for a linear deck.
    slideParts = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort((a, b) => slideOrdinal(a) - slideOrdinal(b));
  }

  const sections: string[] = [];
  const total = { value: 0 };
  for (let index = 0; index < slideParts.length; index++) {
    if (total.value >= OOXML_MAX_TEXT_CHARS) break;
    const slideXml = await readPart(zip, slideParts[index], maxPartBytes);
    if (!slideXml) continue;
    const body = flattenDrawingText(await parseXml(slideXml));
    const notesPart = await notesPartFor(zip, slideParts[index], maxPartBytes);
    const notesXml = notesPart ? await readPart(zip, notesPart, maxPartBytes) : undefined;
    const notes = notesXml ? flattenDrawingText(await parseXml(notesXml)) : '';
    if (!body && !notes) continue;
    const parts = [`[Slide ${index + 1}]`];
    if (body) parts.push(body);
    if (notes) parts.push('[Notes]', notes);
    if (!appendSection(sections, parts.join('\n'), total)) break;
  }
  return sections.join('\n\n');
}

function slideOrdinal(path: string): number {
  const match = /slide(\d+)\.xml$/.exec(path);
  return match ? Number.parseInt(match[1], 10) : 0;
}

/** The notes slide attached to a slide, resolved through the slide's own rels. */
async function notesPartFor(
  zip: ZipArchive,
  slidePath: string,
  maxPartBytes: number = MAX_UNCOMPRESSED_PART_BYTES
): Promise<string | undefined> {
  const directory = slidePath.slice(0, slidePath.lastIndexOf('/') + 1);
  const fileName = slidePath.slice(slidePath.lastIndexOf('/') + 1);
  const rels = await readRelationships(zip, `${directory}_rels/${fileName}.rels`, maxPartBytes);
  const notes = rels.find((rel) => rel.type.endsWith('/notesSlide'));
  return notes ? resolvePart(slidePath, notes.target) : undefined;
}

/**
 * Flatten the text of a DrawingML part (slide or notes slide) in document
 * order: every <a:t> run contributes its text while <a:br> and every end of an
 * <a:p> paragraph start a new line.
 */
function flattenDrawingText(root: XmlNode): string {
  const lines: string[] = [];
  let line = '';
  const flush = (): void => {
    const trimmed = line.trim();
    if (trimmed) lines.push(trimmed);
    line = '';
  };
  const walk = (node: XmlNode): void => {
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (!child || !isElement(child)) continue;
      const name = child.localName;
      if (name === 't') {
        line += child.textContent ?? '';
        continue;
      }
      if (name === 'br') {
        flush();
        continue;
      }
      walk(child);
      if (name === 'p') flush();
    }
  };
  walk(root);
  flush();
  return lines.join('\n');
}
