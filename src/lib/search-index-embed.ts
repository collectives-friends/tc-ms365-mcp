/**
 * OpenAI embeddings client for the semantic search index.
 *
 * This is the only outbound network call the index makes besides Microsoft
 * Graph: document chunks are embedded while indexing and the query text is
 * embedded once per search. The client is deliberately small -- POST
 * /v1/embeddings with bounded retries -- and never logs or echoes the input
 * text, only status codes and counts.
 */

export const DEFAULT_EMBED_MODEL = 'text-embedding-3-small';
export const DEFAULT_EMBED_DIMS = 1024;
export const DEFAULT_OPENAI_BASE_URL = 'https://api.openai.com';

const MAX_ERROR_BODY_CHARS = 300;
const DEFAULT_TIMEOUT_MS = 60_000;
const DEFAULT_RETRIES = 3;
const MAX_RETRY_DELAY_MS = 30_000;

export interface EmbedClientOptions {
  apiKey: string;
  model?: string;
  dims?: number;
  baseUrl?: string;
  fetchImpl?: typeof fetch;
  timeoutMs?: number;
  retries?: number;
  sleepImpl?: (ms: number) => Promise<void>;
}

export interface EmbedResult {
  vectors: Float32Array[];
  promptTokens: number;
}

interface EmbeddingsPayload {
  data?: Array<{ embedding?: unknown; index?: unknown }>;
  usage?: { prompt_tokens?: unknown };
}

function backoffMs(attempt: number): number {
  return Math.min(MAX_RETRY_DELAY_MS, 500 * 2 ** attempt);
}

function parseRetryAfter(header: string | null): number | undefined {
  if (!header) return undefined;
  const seconds = Number.parseFloat(header);
  if (Number.isFinite(seconds) && seconds >= 0) {
    return Math.min(Math.round(seconds * 1000), MAX_RETRY_DELAY_MS);
  }
  const date = Date.parse(header);
  if (Number.isFinite(date)) {
    return Math.min(Math.max(date - Date.now(), 0), MAX_RETRY_DELAY_MS);
  }
  return undefined;
}

function truncate(text: string, maxChars: number): string {
  return text.length > maxChars ? `${text.slice(0, maxChars)}…` : text;
}

export class EmbedClient {
  readonly model: string;
  readonly dims: number;

  private readonly apiKey: string;
  private readonly baseUrl: string;
  private readonly fetchImpl: typeof fetch;
  private readonly timeoutMs: number;
  private readonly retries: number;
  private readonly sleep: (ms: number) => Promise<void>;

  constructor(options: EmbedClientOptions) {
    this.apiKey = options.apiKey;
    this.model = options.model ?? DEFAULT_EMBED_MODEL;
    this.dims = options.dims ?? DEFAULT_EMBED_DIMS;
    this.baseUrl = (options.baseUrl ?? DEFAULT_OPENAI_BASE_URL).replace(/\/+$/, '');
    this.fetchImpl = options.fetchImpl ?? fetch;
    this.timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
    this.retries = options.retries ?? DEFAULT_RETRIES;
    this.sleep = options.sleepImpl ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
  }

  async embedBatch(texts: string[]): Promise<EmbedResult> {
    if (texts.length === 0) return { vectors: [], promptTokens: 0 };
    const body = JSON.stringify({
      model: this.model,
      input: texts,
      dimensions: this.dims,
      encoding_format: 'float',
    });

    let attempt = 0;
    for (;;) {
      let response: Response;
      try {
        response = await this.fetchImpl(`${this.baseUrl}/v1/embeddings`, {
          method: 'POST',
          headers: {
            'content-type': 'application/json',
            authorization: `Bearer ${this.apiKey}`,
          },
          body,
          signal: AbortSignal.timeout(this.timeoutMs),
        });
      } catch (error) {
        if (attempt >= this.retries) {
          const message = error instanceof Error ? error.message : String(error);
          throw new Error(`Embedding request failed: ${message}`);
        }
        await this.sleep(backoffMs(attempt));
        attempt += 1;
        continue;
      }

      if (!response.ok) {
        const errorBody = await response.text().catch(() => '');
        const retriable = response.status === 429 || response.status >= 500;
        if (retriable && attempt < this.retries) {
          const retryAfter = parseRetryAfter(response.headers.get('retry-after'));
          await this.sleep(retryAfter ?? backoffMs(attempt));
          attempt += 1;
          continue;
        }
        throw new Error(
          `Embedding request failed (HTTP ${response.status}): ${truncate(errorBody, MAX_ERROR_BODY_CHARS)}`
        );
      }

      const payload = (await response.json()) as EmbeddingsPayload;
      return parseEmbeddings(payload, texts.length, this.dims);
    }
  }
}

function parseEmbeddings(
  payload: EmbeddingsPayload,
  expectedCount: number,
  dims: number
): EmbedResult {
  const data = payload.data;
  if (!Array.isArray(data) || data.length !== expectedCount) {
    throw new Error(
      `Embedding response has ${Array.isArray(data) ? data.length : 'no'} items; expected ${expectedCount}.`
    );
  }
  const vectors = new Array<Float32Array>(data.length);
  for (let position = 0; position < data.length; position++) {
    const item = data[position];
    const index = typeof item?.index === 'number' ? item.index : position;
    const embedding = item?.embedding;
    if (
      !Number.isInteger(index) ||
      index < 0 ||
      index >= expectedCount ||
      !Array.isArray(embedding) ||
      embedding.length !== dims ||
      !embedding.every((value) => typeof value === 'number' && Number.isFinite(value))
    ) {
      throw new Error(`Embedding response item ${position} is malformed.`);
    }
    vectors[index] = Float32Array.from(embedding as number[]);
  }
  const promptTokens =
    typeof payload.usage?.prompt_tokens === 'number' ? payload.usage.prompt_tokens : 0;
  return { vectors, promptTokens };
}
