# Project Functions: Word Add-in Sidebar

**Date:** 2026-02-25
**Status:** Draft
**Plan Reference:** [Plan 001](plan-001-word-addin-sidebar.md)

---

## Functional Requirements

### FR-001: Load Configuration from URL

**Priority:** Critical
**Phase:** 3B, 5A
**Description:** The add-in must load a JSON configuration file from a user-provided URL using the `fetch()` API. The configuration defines the available API groups, their hierarchies, and individual API call definitions.

**Acceptance Criteria:**
- The user can enter a URL in an input field and click "Load" to fetch the configuration.
- The configuration is fetched over HTTPS using a GET request with `Accept: application/json` header.
- On success, the configuration is parsed, validated, and rendered in the sidebar.
- On failure (network error, HTTP error, invalid JSON), a clear error message is displayed.
- The last successfully used URL is persisted to localStorage and auto-loaded on next task pane open.
- Missing required configuration fields must throw an exception (no fallback values).

---

### FR-002: Display API Groups and Hierarchies in Sidebar

**Priority:** Critical
**Phase:** 5B
**Description:** The sidebar must render the API groups defined in the configuration as a hierarchical, collapsible tree structure. Groups can contain sub-groups (recursive nesting) and/or API definitions.

**Acceptance Criteria:**
- Groups render as collapsible accordion sections using Fluent UI Accordion.
- Sub-groups render recursively inside their parent group.
- Each group displays its name and optional icon.
- API definitions within a group are listed below the group header.
- The tree fits within the ~320px task pane width without horizontal scrolling.
- Empty groups (no apis and no children) are hidden or rendered with a "No APIs configured" message.

---

### FR-003: Capture Optional User Prompt

**Priority:** High
**Phase:** 5B
**Description:** For each API call, the user may optionally provide or edit a text prompt. If the API configuration includes a `promptTemplate`, it is pre-populated in the prompt editor as the default value. The user can modify it before execution.

**Acceptance Criteria:**
- A Textarea component is displayed for APIs that have a `promptTemplate` defined.
- The `promptTemplate` value is shown as the default editable content.
- The `{{text}}` placeholder in the template is replaced with the actual document text at execution time, not at display time.
- The user can clear or modify the prompt before clicking the execute button.
- If no `promptTemplate` is defined, the prompt editor is hidden or shows an empty, optional textarea.

---

### FR-004: Extract Selected Text from Document

**Priority:** Critical
**Phase:** 3A
**Description:** The add-in must extract the currently selected text from the active Word document using the Office.js Word API (`context.document.getSelection()`).

**Acceptance Criteria:**
- Selected text is retrieved using the `Word.run()` proxy object pattern with `selection.load("text")` and `context.sync()`.
- If no text is selected (cursor only), an empty string is returned.
- The extraction handles multi-paragraph selections correctly.
- Known issues with multi-row table selections are documented and handled gracefully.

---

### FR-005: Extract Full Document Text

**Priority:** Critical
**Phase:** 3A
**Description:** The add-in must extract the complete text content of the active Word document using the Office.js Word API (`context.document.body`).

**Acceptance Criteria:**
- Full document text is retrieved using `body.load("text")` and `context.sync()`.
- A loading indicator is shown during extraction for large documents.
- Very large documents (configurable threshold) display a warning about potential performance impact.

---

### FR-006: Dynamic Button Rendering Based on Configuration

**Priority:** Critical
**Phase:** 5B
**Description:** Each API definition specifies an `inputType` field that determines how many action buttons are rendered and their behavior.

**Rendering Rules:**
| `inputType` | Buttons | Behavior |
|-------------|---------|----------|
| `"selected"` | 1 button: "{name} (Selection)" | Sends selected text only |
| `"full"` | 1 button: "{name} (Full Doc)" | Sends full document text only |
| `"both"` | 2 buttons: "{name} (Selection)" + "{name} (Full Doc)" | User chooses which text to send |

**Acceptance Criteria:**
- APIs with `inputType: "selected"` render exactly 1 primary button labeled "{name} (Selection)".
- APIs with `inputType: "full"` render exactly 1 secondary button labeled "{name} (Full Doc)".
- APIs with `inputType: "both"` render 2 buttons side by side.
- Buttons show a loading spinner when the API call is in progress.
- Buttons are disabled during API execution to prevent duplicate calls.
- If `inputType: "selected"` and no text is selected, the button is disabled or shows a tooltip instructing the user to select text.

---

### FR-007: Execute API Calls with Prompt and Text Payload

**Priority:** Critical
**Phase:** 3C
**Description:** When the user clicks an action button, the add-in must execute the configured API call by sending an HTTP request with the prompt and extracted text as the payload.

**Acceptance Criteria:**
- The request is sent to the URL specified in the API configuration.
- The HTTP method matches the `method` field (GET or POST).
- For POST requests, the body is JSON with the structure: `{ "prompt": "<resolved prompt>", "text": "<extracted text>" }`.
- Custom headers from the API configuration `headers` field are included in the request.
- The `{{text}}` placeholder in the prompt template is replaced with the actual document text before sending.
- HTTP errors (4xx, 5xx) result in a user-visible error message.
- CORS failures result in a descriptive error message mentioning CORS.
- Network timeouts are handled with an appropriate error message.

---

### FR-008: Display API Responses

**Priority:** High
**Phase:** 5B
**Description:** After a successful API call, the response must be displayed in the sidebar below the action buttons.

**Acceptance Criteria:**
- The response text is displayed in a styled container with a monospace font and scroll support.
- If the API configuration specifies a `responseField`, the value at that JSON path is extracted and displayed.
- If no `responseField` is specified, the `result` field is used; if absent, the raw JSON is shown.
- The response area has a maximum height with vertical scrolling.
- A "Copy" button allows copying the response to the clipboard.
- An "Insert into Document" button inserts the response text at the current cursor position or replaces the selection.
- Error responses are displayed in a distinct style (red text or error banner).

---

### FR-009: Maintain Prompt and Document History

**Priority:** High
**Phase:** 3D, 5C
**Description:** The add-in must maintain a persistent history of all API interactions, including the prompt, document text (truncated), API name, and response. History survives task pane close/reopen.

**Acceptance Criteria:**
- Each API call creates a history entry with: timestamp, API name, API URL, prompt, input type (selected/full), input text (truncated to 500 characters), and response.
- History is stored in IndexedDB for large capacity and persistence.
- The history panel displays entries in reverse chronological order (newest first).
- Each entry is expandable to show full details.
- A "Clear History" action is available with a confirmation dialog.
- Individual entries can be deleted.
- A "Replay" action on a history entry re-populates the prompt editor and selects the same API.
- History retrieval is limited to the most recent 50 entries by default.

---

### FR-010: Configuration Validation

**Priority:** High
**Phase:** 3B
**Description:** The loaded configuration must be validated against the expected schema. Missing required fields must cause an exception (no fallback or default values per project policy).

**Validation Rules:**
- `version` field: required, must be a non-empty string.
- `title` field: required, must be a non-empty string.
- `groups` field: required, must be a non-empty array.
- Each `ApiGroup.name`: required, must be a non-empty string.
- Each `ApiGroup`: must have at least one of `children` or `apis` (both can be present).
- Each `ApiCallConfig.name`: required, non-empty string.
- Each `ApiCallConfig.url`: required, must be a valid HTTPS URL.
- Each `ApiCallConfig.method`: required, must be "GET" or "POST".
- Each `ApiCallConfig.inputType`: required, must be "selected", "full", or "both".

**Acceptance Criteria:**
- A configuration missing any required field throws a descriptive `ConfigurationError` with the field name and path.
- An invalid URL (non-HTTPS) throws a `ConfigurationError`.
- An invalid `inputType` or `method` value throws a `ConfigurationError`.
- No default values are substituted for missing fields.

---

### FR-011: Configuration Caching

**Priority:** Medium
**Phase:** 3B
**Description:** Successfully loaded configurations must be cached in localStorage with a time-to-live (TTL) to avoid unnecessary network requests.

**Acceptance Criteria:**
- The configuration is cached in localStorage with the source URL and timestamp.
- Cached configurations are served if the TTL (1 hour) has not expired and the URL matches.
- A "Reload" button in the UI forces a cache bypass and fetches fresh configuration.
- Cache uses `Office.context.partitionKey` prefix for storage isolation on Office for the web.

---

### FR-012: State Persistence Across Task Pane Sessions

**Priority:** Medium
**Phase:** 6
**Description:** The task pane webview is destroyed when closed. Critical user state must be persisted to survive close/reopen cycles.

**Persisted State:**
- Configuration URL (localStorage)
- Cached configuration data (localStorage)
- Prompt history (IndexedDB)

**Acceptance Criteria:**
- Reopening the task pane automatically restores the last configuration URL and loads the cached config.
- History entries from previous sessions are immediately available in the History tab.
- No user action is required to restore previous session state.

---

### FR-013: Responsive Design for Task Pane Dimensions

**Priority:** Medium
**Phase:** 6
**Description:** The UI must render correctly within the narrow task pane viewport (default ~320px width, variable height).

**Acceptance Criteria:**
- Single-column layout throughout.
- No horizontal scrollbar at default task pane width.
- Buttons render full-width within their container.
- The personality menu area (top-right corner: 12x32px on Windows, 34x32px on Mac) is not obscured.
- Collapsible sections (Accordion) are used to conserve vertical space.
- Text overflow uses ellipsis or wrapping, not clipping.

---

### FR-014: Error Handling and User Feedback

**Priority:** High
**Phase:** 6
**Description:** All error conditions must be communicated to the user through visible UI elements. The add-in must not fail silently.

**Error Conditions to Handle:**
- Configuration URL unreachable (network error)
- Configuration URL returns non-200 HTTP status
- Configuration JSON is malformed
- Configuration fails schema validation
- API call network error
- API call HTTP error (4xx, 5xx)
- API call CORS failure
- API call timeout
- Office.js API error (e.g., document not available)
- IndexedDB initialization failure

**Acceptance Criteria:**
- Each error condition displays a descriptive message in the UI.
- Error messages include actionable guidance (e.g., "Check that the URL is correct and accessible").
- Errors do not crash the task pane; the React error boundary catches unhandled exceptions.
- A retry option is available for transient errors (network, timeout).

---

### FR-015: Office Ribbon Button Integration

**Priority:** Low
**Phase:** 1
**Description:** The add-in must add a button to the Word ribbon (Home tab) that opens the task pane sidebar.

**Acceptance Criteria:**
- A button labeled "Open Sidebar" appears in the Home tab ribbon under a group named "API Sidebar".
- Clicking the button opens the task pane on the right side of the Word window.
- The button has icons at 16x16, 32x32, and 80x80 pixel sizes.
- A tooltip describes the button's purpose.

---

### FR-016: Prompt Template Resolution

**Priority:** High
**Phase:** 3C
**Description:** API configurations may include a `promptTemplate` string containing the placeholder `{{text}}`. Before sending the API request, the placeholder must be replaced with the actual document text.

**Acceptance Criteria:**
- All occurrences of `{{text}}` in the prompt are replaced with the extracted document text.
- If the prompt contains no `{{text}}` placeholder, the text is appended to the prompt in the request payload.
- If no prompt is provided and no `promptTemplate` exists, only the extracted text is sent in the `text` field.
- Template resolution occurs at execution time, not when the prompt editor is rendered.

---

## Non-Functional Requirements

### NFR-001: HTTPS Only

All network communication (configuration loading, API calls) must use HTTPS. HTTP URLs must be rejected with a validation error.

### NFR-002: No Fallback Configuration Values

Per project policy, missing configuration values must raise exceptions. No default or fallback values are permitted.

### NFR-003: Cross-Platform Compatibility

The add-in must function on Windows (Edge Chromium), macOS (WebKit), and Office for the web (browser iframe).

### NFR-004: Performance

- Configuration loading must complete within 5 seconds (network dependent).
- Text extraction from documents up to 50 pages should complete within 2 seconds.
- UI interactions (button clicks, tab switches) must respond within 100ms.

### NFR-005: Storage Limits

- localStorage usage must stay within 5MB.
- IndexedDB storage for history should auto-prune entries older than 90 days or when exceeding 1000 entries.

### NFR-006: Security

- No sensitive data (API keys, tokens) should be stored in localStorage without encryption.
- CORS is enforced by the browser; the add-in must not attempt to bypass CORS restrictions.
- All external communication uses HTTPS.
