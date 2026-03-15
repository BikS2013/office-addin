# Issues - Pending Items

**Project:** Word Add-in Sidebar
**Last Updated:** 2026-02-25

---

## Pending Items

### P1 - Open Questions from Implementation Plan (Critical)
**Date:** 2026-02-25
**Status:** Open
**Description:** The implementation plan (plan-001) identifies 5 open questions (Q1-Q5) that need decisions before the relevant phases begin:
- Q1: Authentication header support and token storage strategy (needed by Phase 2)
- Q2: Should API responses be insertable into the document? (needed by Phase 5B) -- Partially resolved: ResponsePanel includes "Insert at Cursor" action
- Q3: Maximum history size before auto-pruning (needed by Phase 3D) -- Resolved: MAX_ENTRIES=1000, MAX_AGE_DAYS=90 implemented in HistoryService
- Q4: Config URL hard-coded vs. user-provided (needed by Phase 5A) -- Resolved: user-provided via ConfigLoader
- Q5: Support for multiple simultaneous configurations (needed by Phase 3B) -- Not yet supported

### P2 - CORS Requirements Documentation for API Providers (Medium)
**Date:** 2026-02-25
**Status:** Open
**Description:** API providers serving the add-in must include CORS headers. A requirements document for API providers needs to be created listing the required headers (`Access-Control-Allow-Origin`, `Access-Control-Allow-Methods`, `Access-Control-Allow-Headers`).

### P3 - Icon Assets Creation (Low)
**Date:** 2026-02-25
**Status:** Open
**Description:** Ribbon button icons at 16x16, 32x32, and 80x80 pixel sizes need to be designed and created. These are required for the XML manifest.

### P4 - Implementation Type Divergences from Design Document (Medium)
**Date:** 2026-02-25
**Status:** Open
**Description:** The implementation's TypeScript types diverge from the technical design document in several ways. The implementation is internally consistent and compiles, but does not match the design spec exactly. Key divergences:
- **config.ts**: `configVersion`/`name` instead of design's `version`/`title`/`configUrl`; child groups field named `groups` instead of `children`; `CachedConfiguration` uses numeric `cachedAt`/`ttlMs` instead of ISO string format
- **api.ts**: Discriminated union uses `kind: "success"/"error"` instead of design's `success: true/false`; `ApiSuccessResponse` uses `data` instead of `rawBody`; `ConstructedRequest.body` is `string` instead of `Record<string, unknown>`; `DocumentTextInfo` missing `fullText` field; missing `TextSource` and `ApiErrorType` type exports
- **history.ts**: `timestamp` is `number` instead of ISO string; `textPreview` instead of `inputTextPreview`; missing `inputTextLength` and `responseText` fields
- **errors.ts**: Different constructor signatures (implementation uses `message`/`retryable`/`userMessage`/`cause?`; design uses `code`/`message`/`retryable`/`details?`)
- **state.ts**: `AsyncState` uses `loading: boolean` instead of design's `status: "idle"|"loading"|"success"|"error"`; `AppState` uses `prompt` instead of `currentPrompt`; missing `documentInfo` field; `AppAction` uses direct fields instead of `payload` pattern

**Impact:** The divergences mean the code does not match the design document. Either the design or the code should be updated to be in sync. The implementation is functional as-is.

### P5 - Placeholder Resolution Order Could Cause Unexpected Behavior (Low)
**Date:** 2026-02-25
**Status:** Open
**Description:** In `ApiExecutionService.resolvePlaceholders()`, placeholders `{{prompt}}`, `{{text}}`, and `{{documentName}}` are resolved sequentially. If the user's prompt contains the literal string `{{text}}` or `{{documentName}}`, those will be replaced during subsequent passes. This is not a security vulnerability (output is JSON-stringified), but could produce unexpected body content. Consider escaping placeholder-like patterns in user input or resolving all placeholders in a single pass using a replacement map.

### P6 - Office.context.partitionKey Not Used for localStorage Key Isolation (Low)
**Date:** 2026-02-25
**Status:** Open
**Description:** The design specifies using `Office.context.partitionKey` as a prefix for localStorage keys to ensure isolation in Office web environments. The current `ConfigService.getStorageKey()` uses a hardcoded prefix `"word-addin-sidebar"` without the partition key. This should be updated to include the partition key when available.

### P7 - Validate Script References Missing Package (Low)
**Date:** 2026-02-25
**Status:** Open
**Description:** The `validate` npm script (`office-addin-manifest validate -m manifest.xml`) references `office-addin-manifest`, which was previously available as a transitive dependency of `office-addin-debugging`. After removing `office-addin-debugging` (to eliminate deprecated transitive dependencies), this command is no longer available locally. The script will use npx to fetch it on-demand, but the `-m` flag syntax appears to have changed in newer versions. The script should be updated or the `office-addin-manifest` package should be added as an explicit devDependency if manifest validation is needed.

### P8 - Config Schema JSON File Not Created (Low)
**Date:** 2026-02-25
**Status:** Open
**Description:** The plan (Phase 2, Task 2.6) calls for a JSON Schema file at `src/taskpane/types/config.schema.json` for external configuration validation. This file was not created. The ConfigService performs manual validation instead, which is functional but lacks the formal JSON Schema that external tools could use for validation.

---

## Completed Items

### C1 - HistoryService Singleton Bug in useApiExecution (Critical)
**Date Fixed:** 2026-02-25
**Description:** `useApiExecution` hook created its own `HistoryService` instance that was never initialized (no `initialize()` call). This caused `ensureDb()` to throw on every history write attempt. Fixed by creating a shared singleton module (`historyServiceInstance.ts`) used by both `useHistory` and `useApiExecution`.

### C2 - Default Timeout Fallback Violated No-Fallback Policy (High)
**Date Fixed:** 2026-02-25
**Description:** `ApiExecutionService` had `DEFAULT_TIMEOUT_MS = 30000` used when `apiConfig.timeout` was undefined. Per project policy, no fallback values for configuration settings. Fixed by making `timeout` a required field in `ApiCallConfig` and throwing `ConfigurationError` if missing. `ConfigService.validateApiCall()` now validates timeout as required.

### C3 - HTTP Method Type Too Permissive (Medium)
**Date Fixed:** 2026-02-25
**Description:** `ApiCallConfig.method` allowed `"PUT" | "PATCH"` in addition to `"GET" | "POST"`. The design schema only specifies `"GET" | "POST"`. Fixed to match the design. `ConfigService` validation and `ApiExecutionService.buildRequest()` also updated to only accept GET/POST.

### C4 - Missing promptTemplate Field in ApiCallConfig (Medium)
**Date Fixed:** 2026-02-25
**Description:** The `ApiCallConfig` interface was missing the `promptTemplate` field specified in the design. This prevented the prompt editor from being pre-populated with the API's default template. Added the field and updated the `SELECT_API` reducer case to initialize the prompt from the template.

### C5 - HistoryPanel Replay Creates Incomplete ApiCallConfig (Medium)
**Date Fixed:** 2026-02-25
**Description:** The replay handler in `HistoryPanel` constructed a minimal `ApiCallConfig` object missing the now-required `timeout` field. Fixed to first search the loaded configuration for the matching API, falling back to a minimal object with a hardcoded timeout only when the config is not available.
