# Prompt 001 - Word Add-in API Sidebar: Full Development Lifecycle

## Purpose
This prompt guides a Claude instance through the complete development lifecycle of a generic Microsoft Word add-in implemented as a sidebar panel. The add-in retrieves its configuration from a remote URL, renders API call groups with hierarchies, accepts user prompts, sends selected or complete document text to configured API endpoints, and maintains an execution history.

---

<prompt>

<role>
You are a senior full-stack developer specializing in Microsoft Office Add-ins, TypeScript, and modern web development. You will research, plan, design, implement, and test a Microsoft Word sidebar add-in from scratch. You must follow every instruction precisely and produce production-quality code.
</role>

<context>
The project lives at the root of a repository dedicated to this Office Add-in. All code must be in TypeScript. The project uses UV for Python (if any Python is needed) and follows strict configuration rules: no fallback/default values for configuration settings -- missing config must raise exceptions. All tools developed must be documented in the project's CLAUDE.md. Test scripts go in the `test_scripts/` folder. Plans go in `docs/design/`. Reference material goes in `docs/reference/`.
</context>

<!-- ============================================================ -->
<!-- PHASE 1: RESEARCH AND INVESTIGATION                          -->
<!-- ============================================================ -->

<phase id="1" name="Research and Investigation">

<objective>
Investigate the current best practices, frameworks, and tooling for building Microsoft Word add-ins that render as sidebar (task pane) panels. Produce a research summary document.
</objective>

<tasks>

<task id="1.1" name="Investigate Office Add-in Platforms">
Research the following topics and document your findings:

1. **Office Add-in Architecture**: Understand the Office Add-in platform model (manifest-based, web-technology task panes, Office.js API). Identify whether the add-in runs inside a WebView, how it communicates with Word, and what security/sandboxing constraints exist.

2. **Yeoman Generator vs. Teams Toolkit vs. Manual Setup**: Compare the official scaffolding approaches:
   - `yo office` (Yeoman generator for Office Add-ins)
   - Teams Toolkit in VS Code
   - Manual project setup with a bundler (Webpack, Vite)

   Evaluate which approach gives the most control and is best suited for a TypeScript + React (or plain TS) sidebar add-in.

3. **Manifest Formats**: Research the two manifest formats:
   - XML manifest (classic)
   - Unified manifest (JSON-based, newer)

   Determine which is more appropriate for a Word-only task pane add-in and why.

4. **Office.js Word API Surface**: Investigate the Word JavaScript API capabilities:
   - Reading the full document body text (`Word.Range`, `body.getRange()`)
   - Reading the current selection (`context.document.getSelection()`)
   - Any limitations on text extraction (e.g., headers, footers, embedded objects)

5. **Sideloading and Debugging**: Research how to sideload the add-in for local development on macOS and Windows, and how to debug the task pane (browser dev tools, Office dev tools).

6. **Deployment Options**: Briefly document deployment paths (centralized deployment via Microsoft 365 admin center, AppSource, SharePoint catalog, direct sideloading).
</task>

<task id="1.2" name="Investigate UI Frameworks for Task Panes">
Research UI framework options for the sidebar:

1. **Fluent UI (React)**: The Microsoft-recommended component library for Office Add-ins. Investigate components relevant to our needs: buttons, text areas, tree views (for API group hierarchies), history lists.

2. **Plain HTML/CSS/TypeScript**: Evaluate whether a framework-free approach is viable and what trade-offs it introduces.

3. **State Management**: Determine what state management approach is appropriate (React Context, Zustand, or simple module-level state) given the sidebar's scope.

Recommend the UI approach with justification.
</task>

<task id="1.3" name="Investigate Configuration Schema Design">
Research patterns for describing API call groups and hierarchies in a JSON/YAML configuration file. Consider:

1. How to represent a tree of API groups (nested groups, ordering).
2. How to represent individual API calls within groups (name, URL, input mode).
3. How to represent the input mode: `selected-text-only`, `complete-text-only`, or `both`.
4. How to support HTTP method, headers, authentication tokens, and request body templates.
5. How to version the configuration schema for future evolution.

Produce a draft JSON schema.
</task>

<task id="1.4" name="Document Research Findings">
Compile all findings into a document at `docs/reference/office-addin-research.md` structured with clear sections for each topic above. Include links to official Microsoft documentation.
</task>

</tasks>

<output>
- `docs/reference/office-addin-research.md` -- comprehensive research document
- A clear recommendation on: scaffolding tool, manifest format, UI framework, and configuration format
</output>

</phase>

<!-- ============================================================ -->
<!-- PHASE 2: PLANNING                                            -->
<!-- ============================================================ -->

<phase id="2" name="Planning">

<objective>
Create a detailed development plan based on the Phase 1 findings. The plan must break the work into milestones with clear deliverables.
</objective>

<tasks>

<task id="2.1" name="Create Development Plan">
Write a plan document at `docs/design/plan-001-word-addin-sidebar.md` covering:

1. **Milestone 1 -- Project Scaffolding and Baseline**
   - Initialize the project (chosen scaffolding approach)
   - Configure TypeScript, bundler, linting, and formatting
   - Set up the manifest file targeting Word task pane
   - Verify sideloading works with a "Hello World" sidebar

2. **Milestone 2 -- Configuration System**
   - Define the JSON configuration schema (API groups, hierarchy, API call definitions)
   - Implement configuration fetching from a remote URL
   - Implement configuration parsing and validation
   - Implement error handling for missing/invalid configuration (no fallbacks -- raise exceptions)
   - Implement a configuration URL input/settings UI in the sidebar

3. **Milestone 3 -- Sidebar UI Core**
   - Implement the sidebar layout (header, navigation tree, prompt area, action buttons, output area)
   - Render API groups and sub-groups as a navigable tree/accordion
   - Render individual API calls with their name and description
   - Implement the prompt text input area
   - Implement conditional button rendering based on the API call's input mode configuration:
     - If `selected-text-only`: show only "Send Selected Text" button
     - If `complete-text-only`: show only "Send Complete Document" button
     - If `both`: show both buttons

4. **Milestone 4 -- Document Text Extraction**
   - Implement selected text extraction via Office.js Word API
   - Implement complete document text extraction via Office.js Word API
   - Handle edge cases: no selection, empty document, very large documents

5. **Milestone 5 -- API Invocation Engine**
   - Implement the HTTP client for calling configured APIs
   - Construct request payloads combining: prompt text + document text (selected or complete)
   - Handle API responses (display results in the sidebar output area)
   - Implement error handling for network failures, timeouts, non-2xx responses
   - Implement loading/progress indicators during API calls

6. **Milestone 6 -- History System**
   - Design the history data model (prompt text, document name/identifier, timestamp, API call name, input mode used, response summary)
   - Implement persistent storage for history (localStorage or IndexedDB)
   - Implement the history view in the sidebar (list, search, re-execute)
   - Implement history export functionality

7. **Milestone 7 -- Testing**
   - Unit tests for configuration parsing and validation
   - Unit tests for API invocation logic
   - Unit tests for history management
   - Integration tests for the sidebar UI
   - End-to-end test plan for manual verification with Word

8. **Milestone 8 -- Documentation and Polish**
   - Update CLAUDE.md with all tools developed
   - Create configuration guide at `docs/design/configuration-guide.md`
   - Update `docs/design/project-design.md`
   - Update `docs/design/project-functions.md`
   - Review and update `Issues - Pending Items.md`
</task>

<task id="2.2" name="Define Functional Requirements">
Create or update `docs/design/project-functions.md` with all functional requirements:

- FR-001: The add-in must load as a Word sidebar (task pane)
- FR-002: The add-in must retrieve its configuration from a user-provided URL
- FR-003: The configuration must define API call groups in a hierarchical structure
- FR-004: Each API call must have a name and a target URL
- FR-005: Each API call must accept an optional user-typed prompt from the sidebar
- FR-006: Each API call must be configurable to accept selected text, complete document text, or both
- FR-007: The sidebar must render one or two action buttons per API call based on the input mode configuration
- FR-008: The add-in must extract selected text from the active Word document
- FR-009: The add-in must extract the complete text from the active Word document
- FR-010: The add-in must send the prompt and document text to the configured API endpoint
- FR-011: The add-in must display API responses in the sidebar
- FR-012: The add-in must maintain a history of executions (prompt, document name, timestamp, API call used, input mode)
- FR-013: The history must persist across sidebar sessions (using browser storage)
- FR-014: The add-in must handle configuration errors with clear error messages (no silent fallbacks)
- FR-015: The add-in must handle API errors gracefully with user-visible feedback
</task>

</tasks>

<output>
- `docs/design/plan-001-word-addin-sidebar.md`
- `docs/design/project-functions.md`
</output>

</phase>

<!-- ============================================================ -->
<!-- PHASE 3: DESIGN AND ARCHITECTURE                             -->
<!-- ============================================================ -->

<phase id="3" name="Design and Architecture">

<objective>
Produce a complete architectural design for the Word add-in, including the configuration schema, component architecture, data flow, and storage design.
</objective>

<tasks>

<task id="3.1" name="Define Configuration Schema">
Design the JSON configuration schema. The schema must support:

```json
{
  "schemaVersion": "1.0",
  "configName": "My API Configuration",
  "groups": [
    {
      "id": "group-1",
      "name": "Text Analysis",
      "description": "APIs for analyzing document text",
      "groups": [
        {
          "id": "group-1-1",
          "name": "Sentiment",
          "description": "Sentiment analysis APIs",
          "apiCalls": [
            {
              "id": "api-1",
              "name": "Analyze Sentiment",
              "description": "Analyzes the sentiment of the provided text",
              "url": "https://api.example.com/sentiment",
              "method": "POST",
              "headers": {
                "Authorization": "Bearer {{AUTH_TOKEN}}",
                "Content-Type": "application/json"
              },
              "inputMode": "both",
              "requestBodyTemplate": {
                "prompt": "{{PROMPT}}",
                "text": "{{DOCUMENT_TEXT}}"
              },
              "responseMapping": {
                "displayField": "result"
              }
            }
          ]
        }
      ],
      "apiCalls": []
    }
  ]
}
```

Key design decisions to document:
- `inputMode` values: `"selected"`, `"complete"`, `"both"`
- Template variables: `{{PROMPT}}`, `{{DOCUMENT_TEXT}}`, `{{AUTH_TOKEN}}`
- `responseMapping.displayField`: JSON path to the field to display from the API response
- Groups can nest recursively (groups within groups)
- API calls can exist at any level in the group hierarchy

Create a formal JSON Schema definition file and place it in the project as `src/schema/config-schema.json`.
</task>

<task id="3.2" name="Design Component Architecture">
Design the React component tree for the sidebar:

```
App
+-- ConfigLoader (handles URL input and config fetching)
+-- Sidebar (main layout, shown after config is loaded)
|   +-- Header (config name, settings gear icon)
|   +-- NavigationTree (renders groups and sub-groups)
|   |   +-- GroupNode (expandable/collapsible group)
|   |       +-- GroupNode (nested sub-group, recursive)
|   |       +-- ApiCallItem (leaf node for an API call)
|   +-- ApiCallPanel (shown when an API call is selected)
|   |   +-- PromptInput (text area for the optional prompt)
|   |   +-- ActionButtons (one or two buttons based on inputMode)
|   |   +-- ResponseDisplay (shows API response or error)
|   |   +-- LoadingIndicator
|   +-- HistoryPanel (toggle-able view)
|       +-- HistoryList
|           +-- HistoryItem
+-- ErrorBoundary
```

Document this in `docs/design/project-design.md`.
</task>

<task id="3.3" name="Design Data Flow">
Document the data flow for a typical API call execution:

1. User opens sidebar --> ConfigLoader prompts for configuration URL (or uses previously saved URL)
2. Add-in fetches and validates the configuration JSON
3. Sidebar renders the API group tree from the configuration
4. User navigates to and selects an API call
5. User optionally types a prompt in the PromptInput
6. User clicks "Send Selected Text" or "Send Complete Document"
7. Add-in extracts the appropriate text from Word via Office.js
8. Add-in constructs the HTTP request using the API call's template, substituting `{{PROMPT}}` and `{{DOCUMENT_TEXT}}`
9. Add-in sends the request and displays the response
10. Add-in records the execution in history

Document this in `docs/design/project-design.md`.
</task>

<task id="3.4" name="Design History Data Model">
Design the history entry structure:

```typescript
interface HistoryEntry {
  id: string;                    // UUID
  timestamp: string;             // ISO 8601
  apiCallId: string;             // Reference to the API call config
  apiCallName: string;           // Human-readable name
  apiCallUrl: string;            // The URL that was invoked
  inputMode: "selected" | "complete";  // Which mode was used
  prompt: string;                // The user's prompt (may be empty)
  documentName: string;          // Name of the Word document
  documentTextSnippet: string;   // First N characters of the sent text
  responseStatus: number;        // HTTP status code
  responseSummary: string;       // First N characters of the response
  success: boolean;              // Whether the call succeeded
}
```

Storage: Use IndexedDB via a lightweight wrapper (e.g., `idb` library) for structured storage with indexing and querying capabilities. Fall back considerations: do NOT implement fallbacks -- if IndexedDB is unavailable, raise an error.

Document this in `docs/design/project-design.md`.
</task>

<task id="3.5" name="Design Error Handling Strategy">
Define the error handling approach:

1. **Configuration errors**: Missing URL, unreachable URL, invalid JSON, schema validation failure --> display specific error message in ConfigLoader, do not proceed to Sidebar.
2. **Document text extraction errors**: No active document, empty selection when "Send Selected Text" is clicked --> display inline warning near the action buttons.
3. **API call errors**: Network failure, timeout (configurable, default 30s), non-2xx response --> display error in ResponseDisplay with status code and message.
4. **Storage errors**: IndexedDB unavailable or write failure --> display non-blocking warning, allow the user to continue using the add-in but warn that history will not be saved.

No silent fallbacks. Every error must be surfaced to the user with an actionable message.

Document this in `docs/design/project-design.md`.
</task>

<task id="3.6" name="Architectural Decisions Record">
Document the following architectural decisions in `docs/design/project-design.md`:

- **AD-001**: Use the Yeoman Office generator or manual Vite setup (based on Phase 1 findings)
- **AD-002**: Use XML manifest or unified JSON manifest (based on Phase 1 findings)
- **AD-003**: Use Fluent UI React for the sidebar UI components
- **AD-004**: Use TypeScript strict mode throughout
- **AD-005**: Use IndexedDB (via `idb` library) for history persistence
- **AD-006**: Use `fetch` API for HTTP calls to configured endpoints
- **AD-007**: Configuration schema versioned with `schemaVersion` field
- **AD-008**: No environment variable fallbacks -- all configuration must be explicitly provided
- **AD-009**: The configuration URL is stored in localStorage so it persists across sessions
- **AD-010**: API response display supports plain text and JSON (pretty-printed)
</task>

</tasks>

<output>
- `docs/design/project-design.md` -- complete architectural design
- `src/schema/config-schema.json` -- formal JSON Schema for the configuration file
</output>

</phase>

<!-- ============================================================ -->
<!-- PHASE 4: IMPLEMENTATION                                      -->
<!-- ============================================================ -->

<phase id="4" name="Implementation">

<objective>
Implement the Word add-in according to the design from Phase 3. Follow the milestones from the plan. All code must be TypeScript.
</objective>

<tasks>

<task id="4.1" name="Project Scaffolding">
Steps:
1. Initialize the project using the chosen scaffolding approach (from Phase 1/2 findings).
2. Configure TypeScript with strict mode enabled.
3. Configure the bundler (Webpack or Vite) for development and production builds.
4. Set up ESLint and Prettier with sensible defaults.
5. Create the Word manifest file (XML or unified JSON) that registers a task pane command.
6. Verify the "Hello World" sidebar loads when sideloaded into Word.

Key files to create:
- `package.json`
- `tsconfig.json`
- `.eslintrc.json` or `eslint.config.js`
- `.prettierrc`
- `manifest.xml` (or `manifest.json`)
- `src/taskpane/taskpane.html` -- entry HTML for the sidebar
- `src/taskpane/index.tsx` -- React entry point
- `webpack.config.js` or `vite.config.ts`
</task>

<task id="4.2" name="Implement Configuration System">
Create the following modules:

1. **`src/config/configSchema.ts`** -- TypeScript types for the configuration:
   ```typescript
   export type InputMode = "selected" | "complete" | "both";

   export interface ApiCallConfig {
     id: string;
     name: string;
     description?: string;
     url: string;
     method: "GET" | "POST" | "PUT" | "PATCH";
     headers?: Record<string, string>;
     inputMode: InputMode;
     requestBodyTemplate?: Record<string, unknown>;
     responseMapping?: {
       displayField?: string;
     };
   }

   export interface ApiGroupConfig {
     id: string;
     name: string;
     description?: string;
     groups?: ApiGroupConfig[];
     apiCalls?: ApiCallConfig[];
   }

   export interface AppConfig {
     schemaVersion: string;
     configName: string;
     groups: ApiGroupConfig[];
   }
   ```

2. **`src/config/configLoader.ts`** -- Fetches and validates configuration:
   ```typescript
   export async function loadConfig(url: string): Promise<AppConfig> {
     // Fetch from URL
     // Validate against schema
     // Throw descriptive errors on failure
     // Return parsed config
   }
   ```

3. **`src/config/configValidator.ts`** -- Validates the configuration object against the JSON schema. Use `ajv` library for JSON Schema validation.

4. **`src/config/configStore.ts`** -- Manages the configuration URL in localStorage:
   ```typescript
   export function getStoredConfigUrl(): string | null;
   export function storeConfigUrl(url: string): void;
   ```

Important: If the configuration URL is not provided and not stored, display the ConfigLoader UI. Do NOT fall back to any default URL.
</task>

<task id="4.3" name="Implement Sidebar UI">
Create the React components as designed in Phase 3:

1. **`src/components/App.tsx`** -- Root component. Manages the top-level state: config loaded vs. not loaded, selected API call, history panel visibility.

2. **`src/components/ConfigLoader.tsx`** -- Shown when no config is loaded. Contains:
   - A text input for the configuration URL
   - A "Load Configuration" button
   - Error display area
   - A link/button to load a previously stored URL (if one exists in localStorage)

3. **`src/components/Sidebar.tsx`** -- Main layout after config is loaded. Contains the NavigationTree, ApiCallPanel, and HistoryPanel.

4. **`src/components/NavigationTree.tsx`** -- Renders the hierarchical API groups:
   - Use Fluent UI `Tree` or `Accordion` component
   - Groups are expandable/collapsible
   - API calls are leaf nodes that, when clicked, populate the ApiCallPanel
   - Support recursive nesting

5. **`src/components/ApiCallPanel.tsx`** -- Shown when an API call is selected:
   - Display the API call name and description
   - Render the PromptInput
   - Render ActionButtons based on `inputMode`
   - Render ResponseDisplay

6. **`src/components/PromptInput.tsx`** -- A multi-line text area for the user's optional prompt. Include a character count and clear button.

7. **`src/components/ActionButtons.tsx`** -- Renders buttons conditionally:
   - If `inputMode === "selected"`: one button labeled "Send Selected Text"
   - If `inputMode === "complete"`: one button labeled "Send Complete Document"
   - If `inputMode === "both"`: two buttons, one for each option
   - Buttons must be disabled while an API call is in progress

8. **`src/components/ResponseDisplay.tsx`** -- Displays the API response:
   - If JSON: pretty-print with syntax highlighting
   - If plain text: display in a scrollable pre block
   - If error: display with a red/warning style
   - Include a "Copy to Clipboard" button

9. **`src/components/HistoryPanel.tsx`** -- Toggle-able panel showing execution history:
   - List of HistoryItem components, newest first
   - Search/filter input
   - "Clear History" button with confirmation

10. **`src/components/HistoryItem.tsx`** -- Single history entry display:
    - Timestamp, API call name, document name, input mode, success/failure indicator
    - Expandable to show prompt and response summary
    - "Re-execute" button to reload the same API call and prompt
</task>

<task id="4.4" name="Implement Document Text Extraction">
Create `src/services/documentService.ts`:

```typescript
export async function getSelectedText(): Promise<string> {
  // Use Office.js Word API: context.document.getSelection()
  // Throw if no selection or empty selection
}

export async function getCompleteDocumentText(): Promise<string> {
  // Use Office.js Word API: context.document.body
  // Throw if document is empty
}

export async function getDocumentName(): Promise<string> {
  // Use Office.js to get the document file name/title
  // Return "Untitled" if not available (this is an acceptable display-only default)
}
```

Handle the Office.js async context pattern correctly:
```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();
  return selection.text;
});
```
</task>

<task id="4.5" name="Implement API Invocation Engine">
Create `src/services/apiService.ts`:

```typescript
export interface ApiCallResult {
  status: number;
  headers: Record<string, string>;
  body: unknown;
  rawText: string;
  isJson: boolean;
}

export async function executeApiCall(
  config: ApiCallConfig,
  prompt: string,
  documentText: string
): Promise<ApiCallResult> {
  // 1. Clone the requestBodyTemplate
  // 2. Replace {{PROMPT}} with the actual prompt
  // 3. Replace {{DOCUMENT_TEXT}} with the actual document text
  // 4. Replace any other {{VARIABLE}} placeholders in headers and body
  // 5. Execute the HTTP request using fetch()
  // 6. Parse the response
  // 7. Return the structured result
  // 8. On network error or timeout, throw with a descriptive message
}
```

Template variable substitution must be recursive through the entire request body object, handling nested objects and arrays. Header values must also support template variables.
</task>

<task id="4.6" name="Implement History System">
Create `src/services/historyService.ts`:

```typescript
import { openDB, IDBPDatabase } from "idb";

export async function initHistoryDb(): Promise<void>;
export async function addHistoryEntry(entry: Omit<HistoryEntry, "id">): Promise<string>;
export async function getHistoryEntries(limit?: number, offset?: number): Promise<HistoryEntry[]>;
export async function searchHistory(query: string): Promise<HistoryEntry[]>;
export async function clearHistory(): Promise<void>;
export async function deleteHistoryEntry(id: string): Promise<void>;
```

The IndexedDB database must:
- Be named `word-addin-history`
- Have an object store named `entries` with `id` as the key path
- Have indexes on `timestamp`, `apiCallName`, and `documentName`
- Use auto-generated UUIDs for entry IDs
</task>

<task id="4.7" name="Implement State Management">
Create `src/state/appState.ts` using React Context or Zustand:

```typescript
interface AppState {
  // Configuration
  config: AppConfig | null;
  configUrl: string | null;
  configLoading: boolean;
  configError: string | null;

  // Selected API call
  selectedApiCall: ApiCallConfig | null;

  // Prompt
  currentPrompt: string;

  // API execution
  apiLoading: boolean;
  apiResult: ApiCallResult | null;
  apiError: string | null;

  // History
  historyVisible: boolean;
  historyEntries: HistoryEntry[];

  // Actions
  loadConfig: (url: string) => Promise<void>;
  selectApiCall: (apiCall: ApiCallConfig) => void;
  setPrompt: (prompt: string) => void;
  executeWithSelectedText: () => Promise<void>;
  executeWithCompleteText: () => Promise<void>;
  toggleHistory: () => void;
  refreshHistory: () => Promise<void>;
}
```
</task>

<task id="4.8" name="Implement Styling">
Style the sidebar to match the Office/Fluent design language:
- Use Fluent UI theme tokens for colors, spacing, and typography
- Ensure the sidebar is responsive within the task pane's narrow width (typically 300-400px)
- Ensure proper scrolling behavior for long content (navigation tree, response display, history)
- Support both light and dark themes (Office supports both)
- Use CSS modules or styled-components for component-scoped styles
</task>

<task id="4.9" name="Wire Everything Together">
1. In `src/taskpane/index.tsx`, initialize Office.js and render the App component:
   ```typescript
   Office.onReady((info) => {
     if (info.host === Office.HostType.Word) {
       const root = createRoot(document.getElementById("root")!);
       root.render(<App />);
     }
   });
   ```

2. Ensure the complete data flow works end-to-end:
   - Config URL entry --> fetch --> validate --> render tree
   - Select API call --> enter prompt --> click button --> extract text --> call API --> display result --> save to history

3. Test each flow manually and fix any issues.
</task>

<task id="4.10" name="Create Development Tools">
Create the following utility tools for the project and document them in CLAUDE.md:

1. **Config Validator Tool** (`tools/validate-config.ts`): A CLI tool that takes a configuration JSON file path or URL and validates it against the schema, reporting any errors. This helps users verify their configuration files before loading them in the add-in.
   - Command: `npx ts-node tools/validate-config.ts <path-or-url>`
   - Output: Validation result with detailed error messages if invalid

2. **Config Generator Tool** (`tools/generate-config.ts`): A CLI tool that generates a sample/template configuration file with example API groups and calls, which users can customize.
   - Command: `npx ts-node tools/generate-config.ts --output <path>`
   - Output: A well-commented sample configuration JSON file

3. **History Export Tool** (`tools/export-history.ts`): A CLI-accessible module (also used by the UI) that exports history entries to JSON or CSV format.
   - Command: `npx ts-node tools/export-history.ts --format json|csv --output <path>`
</task>

</tasks>

<output>
- Complete, functional Word add-in source code under `src/`
- Development tools under `tools/`
- Updated CLAUDE.md with tool documentation
- Manifest file for Word task pane
- Build configuration (webpack/vite)
</output>

</phase>

<!-- ============================================================ -->
<!-- PHASE 5: TESTING                                             -->
<!-- ============================================================ -->

<phase id="5" name="Testing">

<objective>
Create a comprehensive test suite covering unit tests, integration tests, and a manual end-to-end test plan. All test scripts must reside in the `test_scripts/` folder or the standard test directory configured by the test runner.
</objective>

<tasks>

<task id="5.1" name="Set Up Test Infrastructure">
1. Install and configure a test runner: **Vitest** or **Jest** with TypeScript support.
2. Configure test paths, coverage thresholds, and module resolution.
3. Set up mocking utilities:
   - Mock for `Office.js` (since tests run outside Word): create `test_scripts/mocks/officeMock.ts`
   - Mock for `fetch`: use `msw` (Mock Service Worker) or a simple fetch mock
   - Mock for `IndexedDB`: use `fake-indexeddb`
4. Add test scripts to `package.json`:
   ```json
   {
     "scripts": {
       "test": "vitest run",
       "test:watch": "vitest",
       "test:coverage": "vitest run --coverage"
     }
   }
   ```
</task>

<task id="5.2" name="Unit Tests - Configuration System">
Create `test_scripts/config/configLoader.test.ts`:
- Test successful configuration loading from a URL
- Test configuration loading with invalid URL (should throw)
- Test configuration loading with unreachable URL (should throw with network error)
- Test configuration loading with invalid JSON response (should throw with parse error)
- Test configuration loading with valid JSON but invalid schema (should throw with validation errors)

Create `test_scripts/config/configValidator.test.ts`:
- Test validation of a fully valid configuration
- Test validation with missing required fields (schemaVersion, configName, groups)
- Test validation with invalid inputMode values
- Test validation with empty groups array
- Test validation of deeply nested groups (3+ levels)
- Test validation with duplicate API call IDs
- Test validation with invalid HTTP methods

Create `test_scripts/config/configStore.test.ts`:
- Test storing and retrieving a config URL
- Test retrieving when no URL is stored (returns null)
- Test overwriting a previously stored URL
</task>

<task id="5.3" name="Unit Tests - API Invocation">
Create `test_scripts/services/apiService.test.ts`:
- Test successful API call with all template variables substituted
- Test API call with empty prompt ({{PROMPT}} replaced with empty string)
- Test API call with large document text
- Test template substitution in nested request body objects
- Test template substitution in headers
- Test handling of JSON response
- Test handling of plain text response
- Test handling of non-2xx response (should include status code in result)
- Test handling of network timeout
- Test handling of network error (DNS failure, connection refused)
- Test that the correct HTTP method is used
</task>

<task id="5.4" name="Unit Tests - Document Service">
Create `test_scripts/services/documentService.test.ts`:
- Test getSelectedText returns the selected text (mocked Office.js)
- Test getSelectedText throws when no text is selected
- Test getCompleteDocumentText returns full document body (mocked Office.js)
- Test getCompleteDocumentText throws when document is empty
- Test getDocumentName returns the document name
- Test getDocumentName returns "Untitled" when name is unavailable
</task>

<task id="5.5" name="Unit Tests - History Service">
Create `test_scripts/services/historyService.test.ts`:
- Test adding a history entry and retrieving it
- Test retrieving entries in reverse chronological order
- Test pagination (limit and offset)
- Test searching history by API call name
- Test searching history by document name
- Test clearing all history
- Test deleting a single history entry
- Test that entry IDs are unique
</task>

<task id="5.6" name="Integration Tests - UI Components">
Create `test_scripts/components/` with the following test files using React Testing Library:

**`ConfigLoader.test.tsx`**:
- Test rendering the config URL input and load button
- Test loading a valid configuration (mocked fetch)
- Test displaying an error for an invalid configuration
- Test loading from a previously stored URL

**`NavigationTree.test.tsx`**:
- Test rendering a flat list of API groups
- Test rendering nested groups
- Test expanding and collapsing groups
- Test selecting an API call updates the panel

**`ApiCallPanel.test.tsx`**:
- Test rendering with inputMode "selected" (one button)
- Test rendering with inputMode "complete" (one button)
- Test rendering with inputMode "both" (two buttons)
- Test prompt input updates state
- Test button click triggers API call flow
- Test loading state disables buttons
- Test response display after successful API call
- Test error display after failed API call

**`HistoryPanel.test.tsx`**:
- Test rendering history entries
- Test searching/filtering history
- Test clearing history with confirmation
- Test re-executing a history entry
</task>

<task id="5.7" name="End-to-End Test Plan">
Create `test_scripts/e2e-test-plan.md` -- a manual test plan for verifying the add-in in a real Word environment:

1. **Sideloading Test**: Verify the add-in loads in Word on Windows and macOS
2. **Config Loading Test**: Enter a config URL, verify the sidebar populates correctly
3. **Invalid Config Test**: Enter an invalid URL, verify error is shown
4. **Selected Text Test**: Select text in Word, click "Send Selected Text", verify the correct text is sent
5. **Complete Document Test**: Click "Send Complete Document", verify all document text is sent
6. **Prompt Test**: Type a prompt, execute an API call, verify the prompt is included in the request
7. **Empty Prompt Test**: Leave prompt empty, execute, verify it works without a prompt
8. **API Error Test**: Configure an API call with an unreachable URL, verify error handling
9. **History Test**: Execute several API calls, open history, verify all are recorded
10. **History Persistence Test**: Close and reopen the sidebar, verify history is preserved
11. **History Search Test**: Search for a specific entry in history
12. **Re-execute Test**: Click re-execute on a history entry, verify it reloads the API call and prompt
13. **Large Document Test**: Open a large document (50+ pages), test complete document extraction
14. **Theme Test**: Switch Office to dark mode, verify the sidebar renders correctly
15. **Multiple Documents Test**: Open two documents, verify each execution records the correct document name
</task>

<task id="5.8" name="Create Mock Configuration Server">
Create `test_scripts/mock-config-server.ts`:
- A simple HTTP server (using Node.js `http` module or Express) that serves sample configuration files
- Useful for local development and testing
- Supports multiple configuration files at different paths (e.g., `/config/basic.json`, `/config/complex.json`, `/config/invalid.json`)
- Command: `npx ts-node test_scripts/mock-config-server.ts --port 3001`

Document this tool in CLAUDE.md.
</task>

</tasks>

<output>
- Complete test suite under `test_scripts/`
- E2E test plan document
- Mock configuration server
- Test coverage report configuration
</output>

</phase>

<!-- ============================================================ -->
<!-- CROSS-CUTTING CONCERNS                                       -->
<!-- ============================================================ -->

<cross-cutting-concerns>

<concern name="Configuration Management">
All configuration settings (the config URL, any local preferences) must be explicitly provided. No fallback values. If a setting is missing, raise an exception or display a clear error to the user. The only exception is the document name, where "Untitled" is acceptable as a display label (not a configuration fallback).
</concern>

<concern name="Error Handling">
Every error must be surfaced to the user. Use typed error classes:
```typescript
export class ConfigurationError extends Error { }
export class ApiInvocationError extends Error { }
export class DocumentExtractionError extends Error { }
export class StorageError extends Error { }
```
</concern>

<concern name="Security">
- The configuration file may contain API keys in header values. These should not be logged or displayed in the UI.
- Consider adding a note about CORS: the APIs called from the sidebar's WebView must support CORS, or the add-in must document this requirement.
- Template variable `{{AUTH_TOKEN}}` should be resolvable from a secure input (e.g., a settings panel where the user enters tokens), not hardcoded in the config.
</concern>

<concern name="Documentation">
At the end of implementation, ensure:
1. `CLAUDE.md` documents all tools with the required XML format
2. `docs/design/project-design.md` is complete and current
3. `docs/design/project-functions.md` lists all functional requirements
4. `docs/design/configuration-guide.md` covers all configuration options per the configuration guide template in the global instructions
5. `Issues - Pending Items.md` is updated at the project root
</concern>

<concern name="Accessibility">
The sidebar must be keyboard-navigable. Use appropriate ARIA roles and labels. Fluent UI components handle much of this by default, but custom components must be reviewed for accessibility.
</concern>

</cross-cutting-concerns>

<!-- ============================================================ -->
<!-- EXECUTION INSTRUCTIONS                                       -->
<!-- ============================================================ -->

<execution-instructions>

<instruction priority="critical">
Execute the phases in order: Phase 1 -> Phase 2 -> Phase 3 -> Phase 4 -> Phase 5. Do not skip phases. Each phase builds on the outputs of the previous phase.
</instruction>

<instruction priority="critical">
All code must be TypeScript with strict mode enabled. No JavaScript files. No `any` types unless absolutely unavoidable (and if used, add a comment explaining why).
</instruction>

<instruction priority="critical">
Do not create fallback values for configuration settings. If a configuration value is missing, raise an exception. This is a hard project rule.
</instruction>

<instruction priority="high">
After completing Phase 4, run the test suite (Phase 5) and fix any failures before considering the implementation complete.
</instruction>

<instruction priority="high">
Document every tool you create in CLAUDE.md using the required XML format:
```xml
<toolName>
    <objective>what the tool does</objective>
    <command>the exact command to run</command>
    <info>detailed description, parameters, examples</info>
</toolName>
```
</instruction>

<instruction priority="high">
Keep `Issues - Pending Items.md` updated throughout development. Add items as you discover them. Remove items as you resolve them. Pending items go on top, completed items below.
</instruction>

<instruction priority="medium">
Use Fluent UI React v9 (the latest stable version) for UI components. Import from `@fluentui/react-components`.
</instruction>

<instruction priority="medium">
The sidebar width is constrained (typically 300-400px). Design all UI to work within this constraint. Use vertical scrolling generously. Avoid horizontal scrolling.
</instruction>

<instruction priority="medium">
When implementing the history system, generate UUIDs using `crypto.randomUUID()` (available in modern browsers and Node.js).
</instruction>

</execution-instructions>

</prompt>
