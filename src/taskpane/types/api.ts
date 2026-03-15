export interface ApiRequestPayload {
  readonly prompt: string;
  readonly text: string;
  readonly documentName: string;
  readonly textSource: "selected" | "full";
}

export interface ConstructedRequest {
  readonly url: string;
  readonly method: string;
  readonly headers: Record<string, string>;
  readonly body?: string;
}

export interface ApiSuccessResponse {
  readonly kind: "success";
  readonly data: unknown;
  readonly extractedText: string;
  readonly statusCode: number;
  readonly durationMs: number;
}

export interface ApiErrorResponse {
  readonly kind: "error";
  readonly errorType: "network" | "cors" | "timeout" | "http" | "parse" | "field_missing";
  readonly message: string;
  readonly statusCode?: number;
  readonly durationMs: number;
}

export type ApiResponse = ApiSuccessResponse | ApiErrorResponse;

export interface DocumentTextInfo {
  readonly selectedText: string;
  readonly hasSelection: boolean;
  readonly documentName: string;
}
