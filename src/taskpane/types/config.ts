export type InputMode = "selected" | "full" | "both";

export interface ApiCallConfig {
  readonly id: string;
  readonly name: string;
  readonly url: string;
  readonly method: "GET" | "POST";
  readonly inputMode: InputMode;
  /** Default prompt template with placeholders: {{prompt}}, {{text}}, {{documentName}}. */
  readonly promptTemplate?: string;
  readonly headers?: Readonly<Record<string, string>>;
  readonly bodyTemplate?: unknown;
  readonly responseField?: string;
  /** Request timeout in milliseconds. Required -- no fallback. */
  readonly timeout: number;
  readonly description?: string;
}

export interface ApiGroup {
  readonly id: string;
  readonly name: string;
  readonly description?: string;
  readonly groups?: readonly ApiGroup[];
  readonly apis?: readonly ApiCallConfig[];
}

export interface AddinConfiguration {
  readonly configVersion: string;
  readonly name: string;
  readonly description?: string;
  readonly groups: readonly ApiGroup[];
}

export interface CachedConfiguration {
  readonly config: AddinConfiguration;
  readonly url: string;
  readonly cachedAt: number;
  readonly ttlMs: number;
}
