import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { ApiExecutionService } from "../src/taskpane/services/ApiExecutionService";
import type {
  ApiCallConfig,
  ApiRequestPayload,
  ApiSuccessResponse,
  ApiErrorResponse,
} from "../src/taskpane/types";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makePayload(overrides?: Partial<ApiRequestPayload>): ApiRequestPayload {
  return {
    prompt: "Summarize this",
    text: "The quick brown fox jumps over the lazy dog.",
    documentName: "report.docx",
    textSource: "selected",
    ...overrides,
  };
}

function makePostConfig(overrides?: Partial<ApiCallConfig>): ApiCallConfig {
  return {
    id: "api-test",
    name: "Test API",
    url: "https://api.example.com/v1/complete",
    method: "POST",
    inputMode: "selected",
    timeout: 30000,
    bodyTemplate: {
      prompt: "{{prompt}}",
      context: "{{text}}",
      source: "{{documentName}}",
    },
    responseField: "result.text",
    ...overrides,
  };
}

function makeGetConfig(overrides?: Partial<ApiCallConfig>): ApiCallConfig {
  return {
    id: "api-get",
    name: "Get API",
    url: "https://api.example.com/v1/status",
    method: "GET",
    inputMode: "full",
    timeout: 15000,
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("ApiExecutionService", () => {
  let service: ApiExecutionService;

  beforeEach(() => {
    service = new ApiExecutionService();
    vi.restoreAllMocks();
    // Provide performance.now if not available in test env
    if (typeof performance === "undefined") {
      vi.stubGlobal("performance", { now: () => Date.now() });
    }
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // =========================================================================
  // buildRequest
  // =========================================================================

  describe("buildRequest()", () => {
    it("resolves {{prompt}}, {{text}}, {{documentName}} in body template", () => {
      const config = makePostConfig();
      const payload = makePayload();
      const request = service.buildRequest(config, payload);

      const parsedBody = JSON.parse(request.body!);
      expect(parsedBody.prompt).toBe("Summarize this");
      expect(parsedBody.context).toBe("The quick brown fox jumps over the lazy dog.");
      expect(parsedBody.source).toBe("report.docx");
    });

    it("resolves nested objects/arrays in body template", () => {
      const config = makePostConfig({
        bodyTemplate: {
          messages: [
            { role: "system", content: "You are a helpful assistant." },
            { role: "user", content: "{{prompt}}: {{text}}" },
          ],
          metadata: {
            document: "{{documentName}}",
          },
        },
      });
      const payload = makePayload();
      const request = service.buildRequest(config, payload);

      const parsedBody = JSON.parse(request.body!);
      expect(parsedBody.messages[1].content).toBe(
        "Summarize this: The quick brown fox jumps over the lazy dog."
      );
      expect(parsedBody.metadata.document).toBe("report.docx");
    });

    it("creates GET request without body", () => {
      const config = makeGetConfig();
      const payload = makePayload();
      const request = service.buildRequest(config, payload);

      expect(request.method).toBe("GET");
      expect(request.body).toBeUndefined();
    });

    it("creates POST request with JSON body and Content-Type header", () => {
      const config = makePostConfig();
      const payload = makePayload();
      const request = service.buildRequest(config, payload);

      expect(request.method).toBe("POST");
      expect(request.body).toBeDefined();
      expect(request.headers["Content-Type"]).toBe("application/json");
    });

    it("merges custom headers with Content-Type", () => {
      const config = makePostConfig({
        headers: {
          Authorization: "Bearer token123",
          "X-Custom": "value",
        },
      });
      const payload = makePayload();
      const request = service.buildRequest(config, payload);

      expect(request.headers["Authorization"]).toBe("Bearer token123");
      expect(request.headers["X-Custom"]).toBe("value");
      expect(request.headers["Content-Type"]).toBe("application/json");
    });

    it("does not override explicit Content-Type header", () => {
      const config = makePostConfig({
        headers: {
          "Content-Type": "text/plain",
        },
      });
      const payload = makePayload();
      const request = service.buildRequest(config, payload);

      expect(request.headers["Content-Type"]).toBe("text/plain");
    });
  });

  // =========================================================================
  // execute
  // =========================================================================

  describe("execute()", () => {
    it("returns ApiSuccessResponse on successful API call", async () => {
      vi.stubGlobal(
        "fetch",
        vi.fn().mockResolvedValue({
          ok: true,
          status: 200,
          statusText: "OK",
          json: () => Promise.resolve({ result: { text: "Generated summary" } }),
          text: () => Promise.resolve(""),
        })
      );

      const result = await service.execute(makePostConfig(), makePayload());

      expect(result.kind).toBe("success");
      const success = result as ApiSuccessResponse;
      expect(success.extractedText).toBe("Generated summary");
      expect(success.statusCode).toBe(200);
    });

    it("extracts nested response field with dot notation", async () => {
      vi.stubGlobal(
        "fetch",
        vi.fn().mockResolvedValue({
          ok: true,
          status: 200,
          statusText: "OK",
          json: () =>
            Promise.resolve({
              choices: [{ message: { content: "Deep nested value" } }],
            }),
          text: () => Promise.resolve(""),
        })
      );

      const config = makePostConfig({
        responseField: "choices.0.message.content",
      });

      const result = await service.execute(config, makePayload());

      expect(result.kind).toBe("success");
      expect((result as ApiSuccessResponse).extractedText).toBe("Deep nested value");
    });

    it('returns error with errorType "field_missing" when responseField not found', async () => {
      vi.stubGlobal(
        "fetch",
        vi.fn().mockResolvedValue({
          ok: true,
          status: 200,
          statusText: "OK",
          json: () => Promise.resolve({ data: "no matching field" }),
          text: () => Promise.resolve(""),
        })
      );

      const config = makePostConfig({ responseField: "nonexistent.path" });
      const result = await service.execute(config, makePayload());

      expect(result.kind).toBe("error");
      expect((result as ApiErrorResponse).errorType).toBe("field_missing");
    });

    it('returns error with errorType "timeout" on AbortError', async () => {
      vi.stubGlobal(
        "fetch",
        vi.fn().mockImplementation(() => {
          const err = new DOMException("The operation was aborted", "AbortError");
          return Promise.reject(err);
        })
      );

      const config = makePostConfig({ timeout: 1 });
      const result = await service.execute(config, makePayload());

      expect(result.kind).toBe("error");
      expect((result as ApiErrorResponse).errorType).toBe("timeout");
    });

    it('returns error with errorType "http" on 4xx/5xx responses', async () => {
      vi.stubGlobal(
        "fetch",
        vi.fn().mockResolvedValue({
          ok: false,
          status: 500,
          statusText: "Internal Server Error",
          text: () => Promise.resolve("Server error body"),
        })
      );

      const result = await service.execute(makePostConfig(), makePayload());

      expect(result.kind).toBe("error");
      const error = result as ApiErrorResponse;
      expect(error.errorType).toBe("http");
      expect(error.statusCode).toBe(500);
    });

    it('returns error with errorType "parse" on invalid JSON response', async () => {
      vi.stubGlobal(
        "fetch",
        vi.fn().mockResolvedValue({
          ok: true,
          status: 200,
          statusText: "OK",
          json: () => Promise.reject(new SyntaxError("Unexpected token")),
          text: () => Promise.resolve("not json"),
        })
      );

      const result = await service.execute(makePostConfig(), makePayload());

      expect(result.kind).toBe("error");
      expect((result as ApiErrorResponse).errorType).toBe("parse");
    });

    it("tracks durationMs", async () => {
      vi.stubGlobal(
        "fetch",
        vi.fn().mockResolvedValue({
          ok: true,
          status: 200,
          statusText: "OK",
          json: () => Promise.resolve({ result: { text: "ok" } }),
          text: () => Promise.resolve(""),
        })
      );

      const result = await service.execute(makePostConfig(), makePayload());

      expect(result.durationMs).toBeTypeOf("number");
      expect(result.durationMs).toBeGreaterThanOrEqual(0);
    });
  });
});
