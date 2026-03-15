# Plan 001: Word Add-in Sidebar Implementation

**Date:** 2026-02-25
**Status:** Draft
**Reference:** [Investigation Document](../reference/investigation-word-addin-sidebar.md)

---

## 1. Overview

This plan describes the phased implementation of a Microsoft Word add-in delivered as a task pane (sidebar). The add-in dynamically loads an API configuration from a remote URL, displays API call groups in a hierarchical tree, and allows users to execute configured API calls using optional prompts combined with selected text or full document text. A history of all interactions is maintained locally via IndexedDB.

### Technology Stack (from Investigation)

| Layer | Technology |
|-------|-----------|
| Scaffolding | Yeoman (`yo office`) with React + TypeScript |
| Manifest | XML (production-ready for Word) |
| UI Framework | Fluent UI React v9 |
| State Management | React Context + hooks |
| Storage (history) | IndexedDB via `idb` wrapper library |
| Storage (settings) | localStorage (with partition key) |
| HTTP Client | Native `fetch()` API |
| Build Tool | Webpack (Yeoman default) |

---

## 2. Implementation Phases

### Phase 1: Project Scaffolding and Foundation

**Objective:** Generate the base Office Add-in project, configure the development environment, and establish the project structure.

**Tasks:**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 1.1 | Scaffold project using `yo office --projectType react --name "Word API Sidebar" --host word --ts true` | Entire project skeleton |
| 1.2 | Install additional dependencies: `@fluentui/react-components`, `idb` | `package.json` |
| 1.3 | Configure `tsconfig.json` with strict mode and appropriate compiler options | `tsconfig.json` |
| 1.4 | Customize the XML manifest: set add-in identity, ribbon button label ("Open Sidebar"), task pane URL, permissions (`ReadWriteDocument`) | `manifest.xml` |
| 1.5 | Create the directory structure under `src/taskpane/` for `components/`, `services/`, `types/`, `hooks/` | Directory structure |
| 1.6 | Verify `npm start` launches Word with the sideloaded add-in and the empty task pane renders | N/A (manual test) |

**Dependencies:** None (starting point).

**Acceptance Criteria:**
- Running `npm start` opens Word with the add-in task pane visible.
- The task pane renders a "Hello World" Fluent UI component.
- No TypeScript compilation errors.
- The XML manifest passes validation (`npx office-addin-manifest validate manifest.xml`).

**Estimated Effort:** 0.5 day

---

### Phase 2: TypeScript Type Definitions and Configuration Schema

**Objective:** Define all TypeScript interfaces for the configuration schema, history entries, and API payloads.

**Tasks:**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 2.1 | Define `ApiCallConfig` interface (name, url, method, inputType, promptTemplate, headers) | `src/taskpane/types/config.ts` |
| 2.2 | Define `ApiGroup` interface (name, icon, children, apis) for recursive hierarchy | `src/taskpane/types/config.ts` |
| 2.3 | Define `AddinConfiguration` interface (version, title, groups) | `src/taskpane/types/config.ts` |
| 2.4 | Define `HistoryEntry` interface (id, timestamp, apiName, apiUrl, prompt, inputType, inputText, response) | `src/taskpane/types/history.ts` |
| 2.5 | Define `ApiRequestPayload` and `ApiResponsePayload` interfaces | `src/taskpane/types/api.ts` |
| 2.6 | Create a JSON Schema file for external configuration validation | `src/taskpane/types/config.schema.json` |

**Dependencies:** Phase 1 (project exists).

**Acceptance Criteria:**
- All interfaces compile without errors.
- The JSON schema validates the example configuration from the investigation document.
- Types are exported and importable from other modules.

**Estimated Effort:** 0.5 day

---

### Phase 3: Core Services Layer

**Objective:** Implement the service classes that handle Word document interaction, configuration loading, API execution, and history persistence.

**Sub-phase 3A: Office API Service**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 3A.1 | Implement `getSelectedText()` using `Word.run()` + proxy pattern | `src/taskpane/services/officeApi.ts` |
| 3A.2 | Implement `getFullDocumentText()` using `context.document.body` | `src/taskpane/services/officeApi.ts` |
| 3A.3 | Implement `replaceSelection(text)` for inserting API responses | `src/taskpane/services/officeApi.ts` |
| 3A.4 | Implement `appendText(text)` for appending to document | `src/taskpane/services/officeApi.ts` |
| 3A.5 | Implement `getAvailableInput()` returning `{ hasSelection, selectedText, fullText }` | `src/taskpane/services/officeApi.ts` |

**Sub-phase 3B: Configuration Service**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 3B.1 | Implement `loadConfig(url)` with `fetch()` and JSON parsing | `src/taskpane/services/configService.ts` |
| 3B.2 | Implement localStorage-based caching with TTL (1 hour) and partition key support | `src/taskpane/services/configService.ts` |
| 3B.3 | Implement `reloadConfig(url)` to bypass cache | `src/taskpane/services/configService.ts` |
| 3B.4 | Implement `saveConfigUrl()` and `getConfigUrl()` for URL persistence | `src/taskpane/services/configService.ts` |
| 3B.5 | Implement configuration validation against the JSON schema (raise exceptions for invalid/missing required fields -- no fallbacks) | `src/taskpane/services/configService.ts` |

**Sub-phase 3C: API Execution Service**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 3C.1 | Implement `executeApiCall(config, prompt, text)` using `fetch()` | `src/taskpane/services/apiService.ts` |
| 3C.2 | Implement prompt template resolution (replace `{{text}}` and `{{prompt}}` placeholders) | `src/taskpane/services/apiService.ts` |
| 3C.3 | Implement error handling for HTTP errors, CORS failures, and timeouts | `src/taskpane/services/apiService.ts` |
| 3C.4 | Implement configurable request headers from API config | `src/taskpane/services/apiService.ts` |

**Sub-phase 3D: History Service**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 3D.1 | Implement IndexedDB initialization using `idb` library | `src/taskpane/services/historyService.ts` |
| 3D.2 | Implement `addEntry(entry)` to store history records | `src/taskpane/services/historyService.ts` |
| 3D.3 | Implement `getRecentEntries(limit)` with descending timestamp ordering | `src/taskpane/services/historyService.ts` |
| 3D.4 | Implement `clearHistory()` for user-initiated cleanup | `src/taskpane/services/historyService.ts` |
| 3D.5 | Implement `deleteEntry(id)` for individual record removal | `src/taskpane/services/historyService.ts` |

**Dependencies:** Phase 2 (type definitions must exist).

**Parallelization:** Sub-phases 3A, 3B, 3C, and 3D are independent and can be developed in parallel.

**Acceptance Criteria:**
- `officeApi.ts`: Successfully extracts selected text and full document text when tested inside Word.
- `configService.ts`: Loads and caches a remote JSON configuration; throws on invalid URL or missing required fields.
- `apiService.ts`: Executes a POST request with prompt + text payload and returns the response body.
- `historyService.ts`: Stores, retrieves, and deletes history entries from IndexedDB.

**Estimated Effort:** 2 days

---

### Phase 4: React Hooks Layer

**Objective:** Create custom React hooks that bridge services with UI components, managing loading states, error handling, and data flow.

**Tasks:**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 4.1 | Implement `useConfig(url)` hook: manages configuration loading state, error state, and cached config data | `src/taskpane/hooks/useConfig.ts` |
| 4.2 | Implement `useHistory()` hook: provides history entries, addEntry, deleteEntry, and clearHistory functions | `src/taskpane/hooks/useHistory.ts` |
| 4.3 | Implement `useDocumentText()` hook: provides `getSelectedText`, `getFullText`, and `getAvailableInput` with loading states | `src/taskpane/hooks/useDocumentText.ts` |
| 4.4 | Implement `useApiExecution(config)` hook: manages API call execution state, results, and error handling | `src/taskpane/hooks/useApiExecution.ts` |

**Dependencies:** Phase 3 (services must exist).

**Parallelization:** All hooks can be developed in parallel once the services are ready.

**Acceptance Criteria:**
- Hooks correctly manage loading/error/success states.
- Hooks properly initialize and clean up service connections (especially IndexedDB).
- State updates trigger re-renders in consuming components.

**Estimated Effort:** 1 day

---

### Phase 5: UI Components

**Objective:** Build the Fluent UI React v9 components for the task pane sidebar.

**Sub-phase 5A: Layout and Navigation**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 5A.1 | Implement `App.tsx` with tab navigation (APIs, History), config URL input bar, title display, and error banner | `src/taskpane/App.tsx` |
| 5A.2 | Implement `ConfigLoader.tsx` with URL input field, Load button, and reload button | `src/taskpane/components/ConfigLoader.tsx` |
| 5A.3 | Implement `StatusBar.tsx` for displaying loading, error, and success messages | `src/taskpane/components/StatusBar.tsx` |
| 5A.4 | Update `taskpane.html` and `index.tsx` with FluentProvider and Office.onReady bootstrap | `src/taskpane/taskpane.html`, `src/taskpane/index.tsx` |

**Sub-phase 5B: API Group Tree and Actions**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 5B.1 | Implement `ApiGroupTree.tsx` using Fluent UI Accordion for groups with recursive rendering of child groups | `src/taskpane/components/ApiGroupTree.tsx` |
| 5B.2 | Implement `ApiButton.tsx` with dynamic button rendering: 1 button for `selected` or `full` inputType; 2 buttons for `both` inputType | `src/taskpane/components/ApiButton.tsx` |
| 5B.3 | Implement `PromptEditor.tsx` with Textarea for optional prompt editing, showing the promptTemplate as default value | `src/taskpane/components/PromptEditor.tsx` |
| 5B.4 | Implement `ResponseDisplay.tsx` for showing API call results with copy-to-clipboard and insert-into-document actions | `src/taskpane/components/ResponseDisplay.tsx` |

**Sub-phase 5C: History Panel**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 5C.1 | Implement `HistoryPanel.tsx` with a scrollable list of history entries (timestamp, API name, truncated prompt) | `src/taskpane/components/HistoryPanel.tsx` |
| 5C.2 | Implement `HistoryEntry.tsx` as an expandable card showing full details (prompt, input text, response) | `src/taskpane/components/HistoryEntry.tsx` |
| 5C.3 | Add "Clear History" button with confirmation dialog | `src/taskpane/components/HistoryPanel.tsx` |
| 5C.4 | Add "Replay" action on history entries to re-populate the prompt editor and API selection | `src/taskpane/components/HistoryEntry.tsx` |

**Dependencies:** Phase 4 (hooks must exist). Sub-phase 5A must complete before 5B and 5C for layout integration.

**Parallelization:** Sub-phases 5B and 5C can be developed in parallel once 5A is done.

**Acceptance Criteria:**
- ConfigLoader: User can enter a URL, click Load, and see the configuration title update.
- ApiGroupTree: Groups render as collapsible accordion sections; nested groups render recursively.
- ApiButton: APIs with `inputType: "both"` show 2 buttons; `selected` or `full` show 1 button.
- PromptEditor: Shows the `promptTemplate` as editable default text; user edits persist during the session.
- HistoryPanel: Displays recent entries in reverse chronological order; entries expand to show details.
- All components render correctly within the ~320px task pane width constraint.

**Estimated Effort:** 3 days

---

### Phase 6: Integration, Styling, and Responsive Design

**Objective:** Wire all components together, apply responsive styles for the narrow task pane viewport, and polish the user experience.

**Tasks:**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 6.1 | Wire App.tsx to connect ConfigLoader, ApiGroupTree, and HistoryPanel via hooks and state | `src/taskpane/App.tsx` |
| 6.2 | Implement responsive CSS for ~320px viewport: single-column layout, full-width buttons, compact spacing | `src/taskpane/taskpane.css` |
| 6.3 | Add loading spinners during config load and API execution | Various components |
| 6.4 | Add error boundaries to prevent task pane crashes from unhandled errors | `src/taskpane/components/ErrorBoundary.tsx` |
| 6.5 | Handle edge cases: empty selection, empty document, network failures, invalid config | Various services and components |
| 6.6 | Verify personality menu clearance (top-right corner 12x32px Windows, 34x32px Mac) | `src/taskpane/taskpane.css` |

**Dependencies:** Phase 5 (all UI components must exist).

**Acceptance Criteria:**
- End-to-end flow works: load config from URL, select text in Word, click API button, see response.
- No horizontal scrolling in the task pane at default 320px width.
- Error states are displayed clearly to the user (not silent failures).
- The add-in gracefully handles network failures with retry-able error messages.

**Estimated Effort:** 1.5 days

---

### Phase 7: Testing

**Objective:** Write unit tests and integration tests. Create manual test scripts.

**Tasks:**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 7.1 | Unit tests for `configService.ts`: config loading, caching, cache expiry, validation failures | `test_scripts/configService.test.ts` |
| 7.2 | Unit tests for `apiService.ts`: request construction, template resolution, error handling | `test_scripts/apiService.test.ts` |
| 7.3 | Unit tests for `historyService.ts`: CRUD operations on IndexedDB (using fake-indexeddb) | `test_scripts/historyService.test.ts` |
| 7.4 | Unit tests for type validation: configuration schema validation | `test_scripts/configValidation.test.ts` |
| 7.5 | Create a mock API server (Express.js or similar) for end-to-end testing of API calls | `test_scripts/mockApiServer.ts` |
| 7.6 | Create a sample configuration JSON for testing | `test_scripts/sample-config.json` |
| 7.7 | Write manual test plan document for in-Word testing scenarios | `test_scripts/manual-test-plan.md` |

**Dependencies:** Phase 6 (integration must be complete).

**Parallelization:** Tests 7.1-7.4 can be written in parallel. 7.5 is independent.

**Acceptance Criteria:**
- All unit tests pass.
- The mock API server responds correctly to configured endpoints.
- The manual test plan covers all functional requirements.

**Estimated Effort:** 2 days

---

### Phase 8: Documentation and Deployment Preparation

**Objective:** Document the add-in, create deployment artifacts, and prepare for distribution.

**Tasks:**

| # | Task | Files Created/Modified |
|---|------|----------------------|
| 8.1 | Document all tools in `CLAUDE.md` using the required XML format | `CLAUDE.md` |
| 8.2 | Create configuration guide explaining all settings, their purpose, and how to obtain/manage them | `docs/design/configuration-guide.md` |
| 8.3 | Create icon assets (16x16, 32x32, 80x80) for the ribbon button | `assets/icon-*.png` |
| 8.4 | Update manifest.xml with production URLs (replace localhost references) | `manifest.xml` |
| 8.5 | Create deployment script for building production bundle | `scripts/build-prod.sh` |
| 8.6 | Write an example configuration JSON with documentation of all fields | `docs/reference/example-config.json` |
| 8.7 | Update `Issues - Pending Items.md` with any remaining items | `Issues - Pending Items.md` |

**Dependencies:** Phase 7 (testing must be complete).

**Acceptance Criteria:**
- `CLAUDE.md` documents all project tools in the required XML format.
- The configuration guide covers all settings per the configuration-guide template.
- The production build produces a deployable bundle without errors.
- The manifest validates for production deployment.

**Estimated Effort:** 1 day

---

## 3. Dependency Graph

```
Phase 1: Scaffolding
    |
    v
Phase 2: Type Definitions
    |
    v
Phase 3: Services (3A, 3B, 3C, 3D in parallel)
    |
    v
Phase 4: React Hooks (all in parallel)
    |
    v
Phase 5: UI Components (5A first, then 5B + 5C in parallel)
    |
    v
Phase 6: Integration & Styling
    |
    v
Phase 7: Testing
    |
    v
Phase 8: Documentation & Deployment
```

**Critical Path:** Phase 1 -> 2 -> 3 -> 4 -> 5A -> 5B -> 6 -> 7 -> 8

**Total Estimated Effort:** 11.5 days (approximately 2.5 weeks)

---

## 4. Files Index

### Files to Be Created

| File | Phase | Purpose |
|------|-------|---------|
| `manifest.xml` | 1, 8 | Office Add-in XML manifest |
| `src/taskpane/types/config.ts` | 2 | Configuration type definitions |
| `src/taskpane/types/history.ts` | 2 | History entry type definitions |
| `src/taskpane/types/api.ts` | 2 | API payload type definitions |
| `src/taskpane/types/config.schema.json` | 2 | JSON Schema for configuration validation |
| `src/taskpane/services/officeApi.ts` | 3A | Word document interaction service |
| `src/taskpane/services/configService.ts` | 3B | Remote configuration loading and caching |
| `src/taskpane/services/apiService.ts` | 3C | API call execution service |
| `src/taskpane/services/historyService.ts` | 3D | IndexedDB history management |
| `src/taskpane/hooks/useConfig.ts` | 4 | Configuration state hook |
| `src/taskpane/hooks/useHistory.ts` | 4 | History state hook |
| `src/taskpane/hooks/useDocumentText.ts` | 4 | Document text extraction hook |
| `src/taskpane/hooks/useApiExecution.ts` | 4 | API execution state hook |
| `src/taskpane/components/ConfigLoader.tsx` | 5A | Configuration URL input component |
| `src/taskpane/components/StatusBar.tsx` | 5A | Status/error display component |
| `src/taskpane/components/ApiGroupTree.tsx` | 5B | API group hierarchy tree component |
| `src/taskpane/components/ApiButton.tsx` | 5B | Dynamic API action button component |
| `src/taskpane/components/PromptEditor.tsx` | 5B | Prompt text editing component |
| `src/taskpane/components/ResponseDisplay.tsx` | 5B | API response display component |
| `src/taskpane/components/HistoryPanel.tsx` | 5C | History list component |
| `src/taskpane/components/HistoryEntry.tsx` | 5C | Individual history entry component |
| `src/taskpane/components/ErrorBoundary.tsx` | 6 | React error boundary component |
| `src/taskpane/App.tsx` | 5A, 6 | Main application component |
| `src/taskpane/index.tsx` | 5A | React entry point with Office.onReady |
| `src/taskpane/taskpane.html` | 5A | HTML entry point |
| `src/taskpane/taskpane.css` | 6 | Global responsive styles |

### Files Modified from Yeoman Scaffold

| File | Phase | Modification |
|------|-------|-------------|
| `package.json` | 1 | Add dependencies (`@fluentui/react-components`, `idb`) |
| `tsconfig.json` | 1 | Strict mode, path aliases |
| `webpack.config.js` | 1 | Any build customizations if needed |

### Test Files

| File | Phase | Purpose |
|------|-------|---------|
| `test_scripts/configService.test.ts` | 7 | Config service unit tests |
| `test_scripts/apiService.test.ts` | 7 | API service unit tests |
| `test_scripts/historyService.test.ts` | 7 | History service unit tests |
| `test_scripts/configValidation.test.ts` | 7 | Configuration validation tests |
| `test_scripts/mockApiServer.ts` | 7 | Mock API server for E2E testing |
| `test_scripts/sample-config.json` | 7 | Sample configuration for testing |
| `test_scripts/manual-test-plan.md` | 7 | Manual in-Word test scenarios |

---

## 5. Risks and Mitigation

| # | Risk | Impact | Probability | Mitigation |
|---|------|--------|-------------|------------|
| R1 | **CORS blocking on API calls.** Target API servers may not include required CORS headers (`Access-Control-Allow-Origin`). | High -- API calls will fail entirely. | Medium | Document CORS requirements for API providers. Provide a CORS proxy option in configuration (optional `proxyUrl` field). Test with mock server that includes proper CORS headers. |
| R2 | **macOS WebKit differences.** The task pane uses WebKit on macOS vs Edge Chromium on Windows, which may cause rendering or API behavior differences. | Medium -- UI may look different or break. | Medium | Test on both platforms during Phase 7. Use Fluent UI components (which are cross-browser tested) rather than custom CSS where possible. |
| R3 | **IndexedDB storage limits or unavailability.** Some corporate environments may restrict IndexedDB access. | Medium -- history feature would break. | Low | Implement a fallback detection: check if IndexedDB is available at startup and show a warning if history cannot be persisted. |
| R4 | **Large document text extraction performance.** Documents with 100+ pages may cause slow `body.load("text")` calls. | Medium -- UI may freeze during text extraction. | Medium | Show a loading indicator during text extraction. Consider implementing a character limit warning. Add a configurable `maxTextLength` to truncate very large documents. |
| R5 | **Configuration URL unavailability.** The remote config server may be down or unreachable. | High -- add-in is non-functional without config. | Low | Cache the last successful configuration in localStorage. Show the cached version with a warning when the remote URL is unreachable. Allow manual reload. |
| R6 | **Office.js API version incompatibility.** Older Office installations may not support all Word.js API features used. | Medium -- some features may not work. | Low | Check `Office.context.requirements.isSetSupported("WordApi", "1.3")` at startup. Document minimum Office version requirements. |
| R7 | **Yeoman generator produces outdated scaffold.** The generator may not include the latest webpack or Fluent UI versions. | Low -- requires manual updates. | Medium | Pin dependency versions in `package.json` after scaffolding. Update Fluent UI to v9 if the generator includes an older version. |
| R8 | **Task pane state loss on close/reopen.** The task pane webview is destroyed when closed. | Medium -- user loses current session state. | High (by design) | Persist critical state (config URL, current API selection) to localStorage. Restore state on task pane re-open via `Office.onReady`. |

---

## 6. Configuration Schema Design

The remote configuration JSON must follow this structure. All fields marked as required must be present; missing required fields cause a validation exception (no fallbacks per project policy).

```typescript
// Required fields are non-optional in the interface.
// The configService MUST throw an error if any required field is missing.

interface AddinConfiguration {
  version: string;           // Required: Schema version (e.g., "1.0")
  title: string;             // Required: Display title for the sidebar
  groups: ApiGroup[];        // Required: At least one group
}

interface ApiGroup {
  name: string;              // Required: Group display name
  icon?: string;             // Optional: Fluent UI icon name
  children?: ApiGroup[];     // Optional: Nested sub-groups
  apis?: ApiCallConfig[];    // Optional: API definitions in this group
  // A group must have at least one of: children or apis
}

interface ApiCallConfig {
  name: string;                        // Required: API display name
  url: string;                         // Required: API endpoint URL (must be HTTPS)
  method: "GET" | "POST";             // Required: HTTP method
  inputType: "selected" | "full" | "both";  // Required: Determines button count
  promptTemplate?: string;             // Optional: Default prompt with {{text}} placeholder
  headers?: Record<string, string>;    // Optional: Additional HTTP headers
  responseField?: string;              // Optional: JSON path to extract response text (default: "result")
}
```

### Button Rendering Logic

| `inputType` Value | Buttons Rendered | Button Labels |
|-------------------|-----------------|---------------|
| `"selected"` | 1 button | "{name} (Selection)" |
| `"full"` | 1 button | "{name} (Full Doc)" |
| `"both"` | 2 buttons | "{name} (Selection)" + "{name} (Full Doc)" |

---

## 7. Open Questions

| # | Question | Impact | Decision Needed By |
|---|----------|--------|--------------------|
| Q1 | Should the add-in support authentication headers (Bearer tokens) in API calls? If so, where should tokens be stored? | Config schema design | Phase 2 |
| Q2 | Should API responses be insertable into the document (replace selection or append), or display-only? | UI design | Phase 5B |
| Q3 | What is the maximum acceptable history size before auto-pruning? | History service design | Phase 3D |
| Q4 | Should the config URL be hard-coded in the manifest or always user-provided? | UX flow | Phase 5A |
| Q5 | Is there a need to support multiple simultaneous configurations (switching between config URLs)? | Config service design | Phase 3B |
