export interface HistoryEntry {
  readonly id: string;
  readonly timestamp: number;
  readonly apiId: string;
  readonly apiName: string;
  readonly apiUrl: string;
  readonly prompt: string;
  readonly textSource: "selected" | "full";
  readonly textPreview: string;
  readonly documentName: string;
  readonly wasSuccessful: boolean;
  readonly responsePreview?: string;
  readonly durationMs: number;
}

export interface HistoryFilter {
  readonly apiId?: string;
  readonly documentName?: string;
  readonly fromTimestamp?: number;
  readonly toTimestamp?: number;
  readonly wasSuccessful?: boolean;
  readonly limit?: number;
}
