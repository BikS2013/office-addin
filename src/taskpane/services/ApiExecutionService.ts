import type {
  ApiCallConfig,
  ApiRequestPayload,
  ApiResponse,
  ApiSuccessResponse,
  ApiErrorResponse,
  ConstructedRequest,
} from "../types";
import { ConfigurationError } from "../types/errors";

export class ApiExecutionService {
  /**
   * Execute an API call using the given configuration and payload.
   * Resolves all placeholders in the body template and prompt.
   * @returns ApiResponse (success or error, never throws for HTTP errors).
   * @throws ConfigurationError if timeout is not configured.
   */
  async execute(
    apiConfig: ApiCallConfig,
    payload: ApiRequestPayload
  ): Promise<ApiResponse> {
    if (apiConfig.timeout === undefined || apiConfig.timeout === null) {
      throw new ConfigurationError(
        `apis.${apiConfig.id}.timeout`,
        "Required field 'timeout' is missing. Each API call must have an explicit timeout configured."
      );
    }

    const startTime = performance.now();

    try {
      const request = this.buildRequest(apiConfig, payload);

      const controller = new AbortController();
      const timeoutMs = apiConfig.timeout;
      const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

      let response: Response;
      try {
        response = await fetch(request.url, {
          method: request.method,
          headers: request.headers,
          body: request.body,
          signal: controller.signal,
        });
      } finally {
        clearTimeout(timeoutId);
      }

      const durationMs = Math.round(performance.now() - startTime);

      if (!response.ok) {
        const errorBody = await response.text().catch(() => "");
        return {
          kind: "error",
          errorType: "http",
          message: `HTTP ${response.status}: ${response.statusText}${errorBody ? ` - ${errorBody}` : ""}`,
          statusCode: response.status,
          durationMs,
        } as ApiErrorResponse;
      }

      let data: unknown;
      try {
        data = await response.json();
      } catch {
        return {
          kind: "error",
          errorType: "parse",
          message: "Failed to parse response as JSON",
          statusCode: response.status,
          durationMs,
        } as ApiErrorResponse;
      }

      let extractedText: string;
      if (apiConfig.responseField) {
        const fieldResult = this.extractResponseField(data, apiConfig.responseField);
        if (fieldResult === undefined) {
          return {
            kind: "error",
            errorType: "field_missing",
            message: `Response field '${apiConfig.responseField}' not found in response`,
            statusCode: response.status,
            durationMs,
          } as ApiErrorResponse;
        }
        extractedText = fieldResult;
      } else {
        extractedText = typeof data === "string" ? data : JSON.stringify(data);
      }

      return {
        kind: "success",
        data,
        extractedText,
        statusCode: response.status,
        durationMs,
      } as ApiSuccessResponse;
    } catch (error: unknown) {
      const durationMs = Math.round(performance.now() - startTime);
      const errorType = this.classifyError(error);
      const message =
        error instanceof Error ? error.message : "Unknown error occurred";

      return {
        kind: "error",
        errorType,
        message,
        durationMs,
      } as ApiErrorResponse;
    }
  }

  /**
   * Build the fully resolved HTTP request from configuration and payload.
   * Replaces all {{prompt}}, {{text}}, and {{documentName}} placeholders.
   */
  buildRequest(
    apiConfig: ApiCallConfig,
    payload: ApiRequestPayload
  ): ConstructedRequest {
    const method = apiConfig.method;
    const headers: Record<string, string> = { ...(apiConfig.headers ?? {}) };

    let body: string | undefined;

    if (method === "POST" && apiConfig.bodyTemplate !== undefined) {
      const resolved = this.resolvePlaceholders(apiConfig.bodyTemplate, payload);
      body = JSON.stringify(resolved);
      if (!headers["Content-Type"]) {
        headers["Content-Type"] = "application/json";
      }
    }

    return {
      url: apiConfig.url,
      method,
      headers,
      body,
    };
  }

  /**
   * Resolve all placeholder tokens recursively in a body template value.
   * - If string: replace {{prompt}}, {{text}}, {{documentName}} with actual values
   * - If object: recursively resolve each value
   * - If array: recursively resolve each element
   * - Otherwise: return as-is
   */
  private resolvePlaceholders(
    template: unknown,
    payload: ApiRequestPayload
  ): unknown {
    if (typeof template === "string") {
      return template
        .replace(/\{\{prompt\}\}/g, payload.prompt)
        .replace(/\{\{text\}\}/g, payload.text)
        .replace(/\{\{documentName\}\}/g, payload.documentName);
    }

    if (Array.isArray(template)) {
      return template.map((item) => this.resolvePlaceholders(item, payload));
    }

    if (typeof template === "object" && template !== null) {
      const resolved: Record<string, unknown> = {};
      for (const [key, value] of Object.entries(template as Record<string, unknown>)) {
        resolved[key] = this.resolvePlaceholders(value, payload);
      }
      return resolved;
    }

    return template;
  }

  /**
   * Extract the response text using the dot-notation responseField path.
   * Supports array indexing (e.g., "choices.0.message.content").
   * Returns undefined if the path does not resolve to a value.
   */
  private extractResponseField(data: unknown, fieldPath: string): string | undefined {
    const segments = fieldPath.split(".");
    let current: unknown = data;

    for (const segment of segments) {
      if (current === null || current === undefined) {
        return undefined;
      }

      if (typeof current !== "object") {
        return undefined;
      }

      const index = Number(segment);
      if (Array.isArray(current) && !isNaN(index)) {
        current = current[index];
      } else {
        current = (current as Record<string, unknown>)[segment];
      }
    }

    if (current === null || current === undefined) {
      return undefined;
    }

    return typeof current === "string" ? current : JSON.stringify(current);
  }

  /**
   * Classify fetch errors into error type categories.
   */
  private classifyError(
    error: unknown
  ): ApiErrorResponse["errorType"] {
    if (
      error instanceof TypeError &&
      (String(error.message).includes("Failed to fetch") ||
        String(error.message).includes("NetworkError"))
    ) {
      // CORS failures in browsers manifest as TypeErrors with "Failed to fetch".
      // Distinguishing CORS from true network errors is not reliably possible;
      // we label as "cors" following the design heuristic since this is the more
      // common cause in an add-in context.
      return "cors";
    }

    if (error instanceof DOMException && error.name === "AbortError") {
      return "timeout";
    }

    return "network";
  }
}
