import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { describe, expect, it } from 'vitest';
import {
  DEFAULT_MAX_DOCUMENT_CHARS,
  HARD_MAX_DOCUMENT_CHARS,
  detectDocumentKind,
  extractDocumentText,
  normalizeMimeType,
  pageDocumentText,
} from '../src/lib/document-text.js';

const fixture = (name: string): Buffer =>
  readFileSync(fileURLToPath(new URL(`./fixtures/${name}`, import.meta.url)));

describe('detectDocumentKind', () => {
  it('detects docx by MIME type or extension', () => {
    expect(
      detectDocumentKind(
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        undefined
      )
    ).toBe('docx');
    expect(detectDocumentKind(undefined, 'report.docx')).toBe('docx');
    expect(detectDocumentKind(undefined, 'REPORT.DOCM')).toBe('docx');
  });

  it('detects pdf by MIME type or extension', () => {
    expect(detectDocumentKind('application/pdf', undefined)).toBe('pdf');
    expect(detectDocumentKind('application/pdf; charset=binary', undefined)).toBe('pdf');
    expect(detectDocumentKind(undefined, 'scan.PDF')).toBe('pdf');
  });

  it('detects plain text by MIME type or extension', () => {
    expect(detectDocumentKind('text/plain', undefined)).toBe('text');
    expect(detectDocumentKind('text/csv', undefined)).toBe('text');
    expect(detectDocumentKind(undefined, 'notes.md')).toBe('text');
    expect(detectDocumentKind(undefined, 'export.csv')).toBe('text');
  });

  it('returns undefined for unsupported formats', () => {
    expect(detectDocumentKind('application/vnd.ms-powerpoint', 'deck.ppt')).toBeUndefined();
    expect(detectDocumentKind(undefined, 'legacy.doc')).toBeUndefined();
    expect(detectDocumentKind(undefined, 'legacy.xls')).toBeUndefined();
    expect(detectDocumentKind(undefined, undefined)).toBeUndefined();
  });
});

describe('normalizeMimeType', () => {
  it('lowercases and drops parameters', () => {
    expect(normalizeMimeType('Application/PDF; charset=binary')).toBe('application/pdf');
    expect(normalizeMimeType(undefined)).toBe('');
  });
});

describe('extractDocumentText', () => {
  it('extracts paragraph and table text from a minimal docx', async () => {
    const text = await extractDocumentText(fixture('minimal.docx'), 'docx');
    expect(text).toContain('Hello from the minimal docx fixture.');
    expect(text).toContain('Second paragraph: extraction must reach this line.');
    expect(text).toContain('Alpha');
    expect(text).toContain('Beta');
  });

  it('extracts the text layer of a minimal pdf', async () => {
    const text = await extractDocumentText(fixture('minimal.pdf'), 'pdf');
    expect(text).toContain('Hello from the minimal PDF fixture.');
    expect(text).toContain('Second PDF line with token 12345.');
  });

  it('passes plain text through as utf8', async () => {
    const text = await extractDocumentText(Buffer.from('hej med dig', 'utf8'), 'text');
    expect(text).toBe('hej med dig');
  });

  it('rejects empty documents', async () => {
    await expect(extractDocumentText(Buffer.alloc(0), 'text')).rejects.toThrow(/empty/i);
  });

  it('rejects corrupt docx payloads with a parser error', async () => {
    await expect(extractDocumentText(Buffer.from('not a zip'), 'docx')).rejects.toThrow(
      /zip|central directory/i
    );
  });
});

describe('pageDocumentText', () => {
  it('returns a bounded slice and a continuation offset', () => {
    const page = pageDocumentText('abcdefghij', 0, 4);
    expect(page.text).toBe('abcd');
    expect(page.truncated).toBe(true);
    expect(page.nextOffset).toBe(4);
  });

  it('returns the tail without truncation', () => {
    const page = pageDocumentText('abcdefghij', 8, 4);
    expect(page.text).toBe('ij');
    expect(page.truncated).toBe(false);
    expect(page.nextOffset).toBeUndefined();
  });

  it('clamps an offset past the end to empty text', () => {
    const page = pageDocumentText('abc', 99, 10);
    expect(page.text).toBe('');
    expect(page.truncated).toBe(false);
  });

  it('exposes sane default and hard caps', () => {
    expect(DEFAULT_MAX_DOCUMENT_CHARS).toBeLessThan(HARD_MAX_DOCUMENT_CHARS);
    expect(HARD_MAX_DOCUMENT_CHARS).toBe(200_000);
  });
});
