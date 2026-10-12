/**
 * PDF extraction lifecycle tests. unpdf is mocked so the tests can observe how
 * the walk touches pages: one at a time, cleaned up after use, with the
 * document destroyed on every exit path. The real-parser behaviour is covered
 * by the fixture tests in document-text.test.ts.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';

const state = vi.hoisted(() => ({
  numPages: 3,
  failOnPage: 0,
  getPageCalls: [] as number[],
  cleaned: [] as number[],
  destroyed: 0,
  inFlight: 0,
  maxInFlight: 0,
}));

vi.mock('unpdf', () => {
  const module = {
    getDocumentProxy: async () => ({
      get numPages() {
        return state.numPages;
      },
      async getPage(pageNumber: number) {
        state.getPageCalls.push(pageNumber);
        state.inFlight += 1;
        state.maxInFlight = Math.max(state.maxInFlight, state.inFlight);
        await Promise.resolve();
        return {
          async getTextContent() {
            state.inFlight -= 1;
            if (state.failOnPage === pageNumber) {
              throw new Error('page exploded');
            }
            return { items: [{ str: `Page ${pageNumber} body`, hasEOL: true }] };
          },
          cleanup() {
            state.cleaned.push(pageNumber);
          },
        };
      },
      async destroy() {
        state.destroyed += 1;
      },
    }),
  };
  // The loader reads the module's default export first, like the other lazy
  // parser loaders; vitest's mock proxy throws when that export is missing.
  return { default: module, ...module };
});

import { extractDocumentText } from '../src/lib/document-text.js';

beforeEach(() => {
  state.numPages = 3;
  state.failOnPage = 0;
  state.getPageCalls = [];
  state.cleaned = [];
  state.destroyed = 0;
  state.inFlight = 0;
  state.maxInFlight = 0;
});

describe('pdf extraction lifecycle', () => {
  it('walks pages one at a time and releases every page and the document', async () => {
    state.numPages = 4;
    const text = await extractDocumentText(Buffer.from('%PDF fake'), 'pdf');
    expect(text).toContain('Page 1 body');
    expect(text).toContain('Page 4 body');
    expect(state.getPageCalls).toEqual([1, 2, 3, 4]);
    // A page may not be requested before the previous page has been read.
    expect(state.maxInFlight).toBe(1);
    expect(state.cleaned).toEqual([1, 2, 3, 4]);
    expect(state.destroyed).toBe(1);
  });

  it('stops at maxChars without touching later pages and still destroys the document', async () => {
    state.numPages = 50;
    const text = await extractDocumentText(Buffer.from('%PDF fake'), 'pdf', { maxChars: 25 });
    expect(text.length).toBeLessThanOrEqual(25);
    expect(state.getPageCalls.length).toBeLessThan(50);
    expect(state.destroyed).toBe(1);
  });

  it('destroys the document when a page fails', async () => {
    state.failOnPage = 2;
    await expect(extractDocumentText(Buffer.from('%PDF fake'), 'pdf')).rejects.toThrow(/exploded/);
    expect(state.destroyed).toBe(1);
  });
});
