# Technical Design: Word Add-in Sidebar

**Date:** 2026-02-25
**Status:** Draft
**Plan Reference:** [Plan 001](plan-001-word-addin-sidebar.md)
**Functional Requirements:** [Project Functions](project-functions.md)
**Investigation:** [Investigation](../reference/investigation-word-addin-sidebar.md)

---

## Table of Contents

1. [Overview](#1-overview)
2. [Configuration Schema (JSON)](#2-configuration-schema-json)
3. [TypeScript Type Definitions](#3-typescript-type-definitions)
4. [Component Architecture (React)](#4-component-architecture-react)
5. [Service Layer](#5-service-layer)
6. [Data Flow](#6-data-flow)
7. [Error Handling](#7-error-handling)
8. [File Structure](#8-file-structure)
9. [Parallel Implementation Units](#9-parallel-implementation-units)

---

## 1. Overview

This document defines the technical design for a Microsoft Word add-in implemented as a task pane (sidebar). The add-in dynamically loads API configurations from a remote URL, renders API groups as a hierarchical tree, and allows users to execute API calls with optional prompts combined with selected text or full document text. A persistent history of all interactions is maintained via IndexedDB.

### Technology Stack

| Layer | Technology |
|-------|-----------|
| Runtime | Office.js Word API (WordApi 1.3+) |
| Framework | React 18 + TypeScript (strict mode) |
| UI Library | Fluent UI React v9 (`@fluentui/react-components`) |
| State Management | React Context + `useReducer` |
| Persistent Storage | IndexedDB via `idb` library |
| Session Storage | `localStorage` (with `Office.context.partitionKey` prefix) |
| HTTP Client | Native `fetch()` API |
| Build | Webpack (Yeoman default) |
| Manifest | XML format (production-ready for Word) |

---

## 2. Configuration Schema (JSON)

### 2.1 JSON Schema Definition

The remote configuration file must conform to the following JSON Schema. The `configService` validates all loaded configurations against this schema and throws `ConfigurationError` for any violation. No fallback or default values are permitted.

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "https://schema.word-addin-sidebar/config/v1",
  "title": "AddinConfiguration",
  "description": "Configuration schema for the Word Add-in Sidebar. All required fields must be present; missing fields cause a validation exception.",
  "type": "object",
  "required": ["version", "title", "configUrl", "groups"],
  "additionalProperties": false,
  "properties": {
    "version": {
      "type": "string",
      "minLength": 1,
      "description": "Schema version identifier (e.g., '1.0'). Required, non-empty."
    },
    "title": {
      "type": "string",
      "minLength": 1,
      "description": "Display title rendered in the sidebar header. Required, non-empty."
    },
    "configUrl": {
      "type": "string",
      "format": "uri",
      "pattern": "^https://",
      "description": "The canonical URL from which this configuration was loaded. Used for cache keying."
    },
    "groups": {
      "type": "array",
      "minItems": 1,
      "description": "Top-level API groups. At least one group is required.",
      "items": { "$ref": "#/definitions/ApiGroup" }
    }
  },
  "definitions": {
    "ApiGroup": {
      "type": "object",
      "required": ["id", "name"],
      "additionalProperties": false,
      "properties": {
        "id": {
          "type": "string",
          "minLength": 1,
          "description": "Unique identifier for the group within the configuration."
        },
        "name": {
          "type": "string",
          "minLength": 1,
          "description": "Display name for the group. Required, non-empty."
        },
        "icon": {
          "type": "string",
          "description": "Fluent UI icon name (e.g., 'Document', 'Settings'). Optional."
        },
        "description": {
          "type": "string",
          "description": "Optional description displayed as a subtitle or tooltip."
        },
        "children": {
          "type": "array",
          "items": { "$ref": "#/definitions/ApiGroup" },
          "description": "Nested sub-groups for hierarchical organization."
        },
        "apis": {
          "type": "array",
          "items": { "$ref": "#/definitions/ApiCallConfig" },
          "description": "API call definitions belonging to this group."
        }
      },
      "anyOf": [
        { "required": ["children"] },
        { "required": ["apis"] }
      ]
    },
    "ApiCallConfig": {
      "type": "object",
      "required": ["id", "name", "url", "method", "inputMode"],
      "additionalProperties": false,
      "properties": {
        "id": {
          "type": "string",
          "minLength": 1,
          "description": "Unique identifier for the API call within the configuration."
        },
        "name": {
          "type": "string",
          "minLength": 1,
          "description": "Display name for the API call. Required, non-empty."
        },
        "description": {
          "type": "string",
          "description": "Optional description shown as a tooltip or subtitle."
        },
        "url": {
          "type": "string",
          "format": "uri",
          "pattern": "^https://",
          "description": "API endpoint URL. Must be HTTPS. Required."
        },
        "method": {
          "type": "string",
          "enum": ["GET", "POST"],
          "description": "HTTP method for the API call. Required."
        },
        "inputMode": {
          "type": "string",
          "enum": ["selected", "full", "both"],
          "description": "Determines button rendering: 'selected' = 1 button (selection only), 'full' = 1 button (full document), 'both' = 2 buttons. Required."
        },
        "promptTemplate": {
          "type": "string",
          "description": "Default prompt text with placeholders: {{prompt}}, {{text}}, {{documentName}}. Shown pre-populated in the prompt editor."
        },
        "headers": {
          "type": "object",
          "additionalProperties": { "type": "string" },
          "description": "Additional HTTP headers to include in the request."
        },
        "bodyTemplate": {
          "type": "object",
          "description": "JSON body template for POST requests. String values may contain placeholders: {{prompt}}, {{text}}, {{documentName}}. If not provided, a default structure is used: { \"prompt\": \"{{prompt}}\", \"text\": \"{{text}}\" }."
        },
        "responseField": {
          "type": "string",
          "description": "Dot-notation JSON path to extract the response text (e.g., 'data.result.text'). If not specified, 'result' is used. If 'result' is absent, raw JSON is displayed."
        },
        "timeout": {
          "type": "integer",
          "minimum": 1000,
          "maximum": 300000,
          "description": "Request timeout in milliseconds. Required in configuration (no default)."
        }
      }
    }
  }
}
```

### 2.2 Example Configuration

```json
{
  "version": "1.0",
  "title": "AI Writing Assistant",
  "configUrl": "https://api.example.com/config/writing-tools.json",
  "groups": [
    {
      "id": "text-analysis",
      "name": "Text Analysis",
      "icon": "DocumentSearch",
      "description": "Tools for analyzing document text",
      "apis": [
        {
          "id": "summarize",
          "name": "Summarize",
          "description": "Generate a summary of the text",
          "url": "https://api.example.com/v1/summarize",
          "method": "POST",
          "inputMode": "both",
          "promptTemplate": "Summarize the following text:\n\n{{text}}",
          "headers": {
            "Authorization": "Bearer {{token}}",
            "X-Api-Version": "2024-01"
          },
          "bodyTemplate": {
            "model": "gpt-4",
            "messages": [
              {
                "role": "user",
                "content": "{{prompt}}"
              }
            ],
            "context": {
              "document": "{{documentName}}",
              "text": "{{text}}"
            }
          },
          "responseField": "choices.0.message.content",
          "timeout": 30000
        },
        {
          "id": "grammar-check",
          "name": "Grammar Check",
          "url": "https://api.example.com/v1/grammar",
          "method": "POST",
          "inputMode": "selected",
          "bodyTemplate": {
            "text": "{{text}}"
          },
          "responseField": "corrections",
          "timeout": 15000
        }
      ]
    },
    {
      "id": "ai-tools",
      "name": "AI Tools",
      "icon": "Bot",
      "children": [
        {
          "id": "generation",
          "name": "Content Generation",
          "apis": [
            {
              "id": "expand",
              "name": "Expand Text",
              "url": "https://api.example.com/v1/expand",
              "method": "POST",
              "inputMode": "selected",
              "promptTemplate": "Expand the following text into a detailed paragraph:\n\n{{text}}",
              "bodyTemplate": {
                "prompt": "{{prompt}}",
                "text": "{{text}}"
              },
              "responseField": "result",
              "timeout": 30000
            }
          ]
        },
        {
          "id": "translation",
          "name": "Translation",
          "apis": [
            {
              "id": "translate-en-fr",
              "name": "English to French",
              "url": "https://api.example.com/v1/translate",
              "method": "POST",
              "inputMode": "both",
              "promptTemplate": "Translate the following English text to French:\n\n{{text}}",
              "bodyTemplate": {
                "source_lang": "en",
                "target_lang": "fr",
                "text": "{{text}}"
              },
              "responseField": "translation",
              "timeout": 20000
            }
          ]
        }
      ]
    }
  ]
}
```

### 2.3 Placeholder Resolution Rules

The configuration supports three placeholders in `promptTemplate` and in string values of `bodyTemplate`:

| Placeholder | Resolved To | Context |
|-------------|-------------|---------|
| `{{prompt}}` | The user-edited prompt text from the PromptEditor | The final prompt after the user has optionally modified the `promptTemplate` |
| `{{text}}` | The extracted document text (selected or full, depending on which button was clicked) | Raw text from `Word.run()` |
| `{{documentName}}` | The active document filename | From `Office.context.document` properties |

**Resolution order:**
1. The `promptTemplate` is displayed in the PromptEditor with `{{text}}` and `{{documentName}}` as literal strings (not resolved at display time).
2. When the user clicks an action button, the text is extracted from Word.
3. The user-edited prompt replaces `{{prompt}}` in the `bodyTemplate`.
4. `{{text}}` in both the resolved prompt and `bodyTemplate` is replaced with the extracted text.
5. `{{documentName}}` in both the resolved prompt and `bodyTemplate` is replaced with the document name.

---

## 3. TypeScript Type Definitions

### 3.1 Configuration Types (`src/taskpane/types/config.ts`)

```typescript
/**
 * Input mode determining how many action buttons render and their behavior.
 * - "selected": 1 button, sends selected text only
 * - "full": 1 button, sends full document text only
 * - "both": 2 buttons, user chooses which text to send
 */
export type InputMode = "selected" | "full" | "both";

/**
 * HTTP methods supported by API call configurations.
 */
export type HttpMethod = "GET" | "POST";

/**
 * Represents a single API call definition within a group.
 * All required fields must be present in the configuration JSON.
 */
export interface ApiCallConfig {
  /** Unique identifier for the API call within the configuration. */
  readonly id: string;
  /** Display name for the API call. */
  readonly name: string;
  /** Optional description shown as tooltip or subtitle. */
  readonly description?: string;
  /** API endpoint URL. Must be HTTPS. */
  readonly url: string;
  /** HTTP method for the request. */
  readonly method: HttpMethod;
  /** Determines button rendering and text extraction behavior. */
  readonly inputMode: InputMode;
  /** Default prompt with placeholders: {{prompt}}, {{text}}, {{documentName}}. */
  readonly promptTemplate?: string;
  /** Additional HTTP headers. */
  readonly headers?: Readonly<Record<string, string>>;
  /** JSON body template for POST requests. String values may contain placeholders. */
  readonly bodyTemplate?: Readonly<Record<string, unknown>>;
  /** Dot-notation path to extract response text (e.g., "data.result.text"). */
  readonly responseField?: string;
  /** Request timeout in milliseconds. */
  readonly timeout?: number;
}

/**
 * Represents a group of API calls, potentially containing nested sub-groups.
 * A group must have at least one of: children or apis.
 */
export interface ApiGroup {
  /** Unique identifier for the group. */
  readonly id: string;
  /** Display name for the group. */
  readonly name: string;
  /** Fluent UI icon name (e.g., "Document", "Settings"). */
  readonly icon?: string;
  /** Optional description displayed as subtitle or tooltip. */
  readonly description?: string;
  /** Nested sub-groups for hierarchical organization. */
  readonly children?: readonly ApiGroup[];
  /** API call definitions belonging to this group. */
  readonly apis?: readonly ApiCallConfig[];
}

/**
 * Root configuration object loaded from the remote URL.
 * All required fields must be present; missing fields throw ConfigurationError.
 */
export interface AddinConfiguration {
  /** Schema version identifier (e.g., "1.0"). */
  readonly version: string;
  /** Display title rendered in the sidebar header. */
  readonly title: string;
  /** The canonical URL from which this configuration was loaded. */
  readonly configUrl: string;
  /** Top-level API groups. At least one group required. */
  readonly groups: readonly ApiGroup[];
}

/**
 * Cached configuration with metadata for TTL-based cache invalidation.
 */
export interface CachedConfiguration {
  /** The configuration data. */
  readonly config: AddinConfiguration;
  /** ISO 8601 timestamp of when the configuration was cached. */
  readonly cachedAt: string;
  /** The URL from which the configuration was loaded. */
  readonly sourceUrl: string;
}
```

### 3.2 API Types (`src/taskpane/types/api.ts`)

```typescript
import { InputMode, HttpMethod } from "./config";

/**
 * The text source that was used for the API call.
 */
export type TextSource = "selected" | "full";

/**
 * Resolved request payload after placeholder substitution.
 */
export interface ApiRequestPayload {
  /** The resolved prompt text (after user edits, before placeholder substitution in body). */
  readonly prompt: string;
  /** The extracted document text. */
  readonly text: string;
  /** The document filename. */
  readonly documentName: string;
  /** Whether the text came from selection or full document. */
  readonly textSource: TextSource;
}

/**
 * Fully constructed HTTP request ready for execution.
 */
export interface ConstructedRequest {
  /** Target URL. */
  readonly url: string;
  /** HTTP method. */
  readonly method: HttpMethod;
  /** Resolved HTTP headers. */
  readonly headers: Readonly<Record<string, string>>;
  /** Resolved JSON body (for POST requests). Undefined for GET. */
  readonly body?: Readonly<Record<string, unknown>>;
  /** Timeout in milliseconds. */
  readonly timeout: number;
}

/**
 * Successful API response.
 */
export interface ApiSuccessResponse {
  readonly success: true;
  /** HTTP status code. */
  readonly statusCode: number;
  /** The raw response body (parsed JSON). */
  readonly rawBody: unknown;
  /** The extracted response text (using responseField path). */
  readonly extractedText: string;
  /** Response time in milliseconds. */
  readonly durationMs: number;
}

/**
 * Failed API response.
 */
export interface ApiErrorResponse {
  readonly success: false;
  /** Error category for UI display. */
  readonly errorType: ApiErrorType;
  /** Human-readable error message. */
  readonly message: string;
  /** HTTP status code (if applicable). */
  readonly statusCode?: number;
  /** Raw error details for debugging. */
  readonly details?: string;
  /** Response time in milliseconds. */
  readonly durationMs: number;
}

/**
 * Discriminated union of API response types.
 */
export type ApiResponse = ApiSuccessResponse | ApiErrorResponse;

/**
 * Classification of API errors for appropriate UI display and retry logic.
 */
export type ApiErrorType =
  | "NETWORK_ERROR"
  | "CORS_ERROR"
  | "TIMEOUT_ERROR"
  | "HTTP_CLIENT_ERROR"
  | "HTTP_SERVER_ERROR"
  | "PARSE_ERROR"
  | "RESPONSE_EXTRACTION_ERROR";

/**
 * Information about the current document text availability.
 */
export interface DocumentTextInfo {
  /** Whether any text is currently selected. */
  readonly hasSelection: boolean;
  /** The selected text (empty string if no selection). */
  readonly selectedText: string;
  /** The full document body text. */
  readonly fullText: string;
  /** The document filename. */
  readonly documentName: string;
}
```

### 3.3 History Types (`src/taskpane/types/history.ts`)

```typescript
import { TextSource } from "./api";

/**
 * A single history entry representing one API interaction.
 * Stored in IndexedDB for persistence across task pane sessions.
 */
export interface HistoryEntry {
  /** Auto-generated unique identifier (UUID v4). */
  readonly id: string;
  /** ISO 8601 timestamp of when the API call was executed. */
  readonly timestamp: string;
  /** The API call identifier from the configuration. */
  readonly apiId: string;
  /** The API display name at the time of execution. */
  readonly apiName: string;
  /** The API endpoint URL. */
  readonly apiUrl: string;
  /** The user-entered or template-derived prompt text. */
  readonly prompt: string;
  /** Whether selected text or full document text was used. */
  readonly textSource: TextSource;
  /** The document text sent to the API (truncated to 500 characters for storage). */
  readonly inputTextPreview: string;
  /** The full length of the original input text (before truncation). */
  readonly inputTextLength: number;
  /** The document name at the time of execution. */
  readonly documentName: string;
  /** The extracted response text (or error message if the call failed). */
  readonly responseText: string;
  /** Whether the API call was successful. */
  readonly wasSuccessful: boolean;
  /** Response time in milliseconds. */
  readonly durationMs: number;
}

/**
 * Filter criteria for history queries.
 */
export interface HistoryFilter {
  /** Filter by API identifier. */
  readonly apiId?: string;
  /** Filter by text source type. */
  readonly textSource?: TextSource;
  /** Filter entries after this date (ISO 8601). */
  readonly after?: string;
  /** Filter entries before this date (ISO 8601). */
  readonly before?: string;
  /** Maximum number of entries to return. */
  readonly limit?: number;
}

/**
 * IndexedDB database schema version info.
 */
export interface HistoryDbSchema {
  readonly dbName: "word-addin-sidebar-history";
  readonly version: 1;
  readonly storeName: "history";
  readonly indexes: {
    readonly byTimestamp: "timestamp";
    readonly byApiId: "apiId";
  };
}
```

### 3.4 Application State Types (`src/taskpane/types/state.ts`)

```typescript
import { AddinConfiguration, ApiCallConfig } from "./config";
import { ApiResponse, DocumentTextInfo } from "./api";
import { HistoryEntry } from "./history";

/**
 * Active tab in the sidebar navigation.
 */
export type ActiveTab = "apis" | "history";

/**
 * Loading state for async operations.
 */
export interface AsyncState<T> {
  readonly status: "idle" | "loading" | "success" | "error";
  readonly data: T | null;
  readonly error: AppError | null;
}

/**
 * Structured application error.
 */
export interface AppError {
  readonly code: string;
  readonly message: string;
  readonly details?: string;
  readonly retryable: boolean;
}

/**
 * Root application state managed by the top-level reducer.
 */
export interface AppState {
  /** Current active tab. */
  readonly activeTab: ActiveTab;
  /** Configuration loading state. */
  readonly config: AsyncState<AddinConfiguration>;
  /** The persisted configuration URL. */
  readonly configUrl: string;
  /** Currently selected API call (if any). */
  readonly selectedApi: ApiCallConfig | null;
  /** Current prompt text in the editor. */
  readonly currentPrompt: string;
  /** API execution state. */
  readonly execution: AsyncState<ApiResponse>;
  /** History entries (most recent first). */
  readonly history: AsyncState<readonly HistoryEntry[]>;
  /** Document text information. */
  readonly documentInfo: AsyncState<DocumentTextInfo>;
}

/**
 * All possible actions dispatched to the application reducer.
 */
export type AppAction =
  | { readonly type: "SET_ACTIVE_TAB"; readonly payload: ActiveTab }
  | { readonly type: "SET_CONFIG_URL"; readonly payload: string }
  | { readonly type: "CONFIG_LOAD_START" }
  | { readonly type: "CONFIG_LOAD_SUCCESS"; readonly payload: AddinConfiguration }
  | { readonly type: "CONFIG_LOAD_ERROR"; readonly payload: AppError }
  | { readonly type: "SELECT_API"; readonly payload: ApiCallConfig | null }
  | { readonly type: "SET_PROMPT"; readonly payload: string }
  | { readonly type: "EXECUTION_START" }
  | { readonly type: "EXECUTION_SUCCESS"; readonly payload: ApiResponse }
  | { readonly type: "EXECUTION_ERROR"; readonly payload: AppError }
  | { readonly type: "HISTORY_LOAD_START" }
  | { readonly type: "HISTORY_LOAD_SUCCESS"; readonly payload: readonly HistoryEntry[] }
  | { readonly type: "HISTORY_LOAD_ERROR"; readonly payload: AppError }
  | { readonly type: "HISTORY_ENTRY_ADDED"; readonly payload: HistoryEntry }
  | { readonly type: "HISTORY_ENTRY_DELETED"; readonly payload: string }
  | { readonly type: "HISTORY_CLEARED" }
  | { readonly type: "DOCUMENT_INFO_LOAD_START" }
  | { readonly type: "DOCUMENT_INFO_LOAD_SUCCESS"; readonly payload: DocumentTextInfo }
  | { readonly type: "DOCUMENT_INFO_LOAD_ERROR"; readonly payload: AppError };
```

### 3.5 Error Types (`src/taskpane/types/errors.ts`)

```typescript
/**
 * Base error class for all application-specific errors.
 */
export class AppBaseError extends Error {
  public readonly code: string;
  public readonly retryable: boolean;
  public readonly details?: string;

  constructor(code: string, message: string, retryable: boolean, details?: string) {
    super(message);
    this.name = this.constructor.name;
    this.code = code;
    this.retryable = retryable;
    this.details = details;
  }
}

/**
 * Thrown when configuration validation fails.
 * NEVER recoverable with fallback values -- the configuration must be fixed.
 */
export class ConfigurationError extends AppBaseError {
  public readonly fieldPath: string;

  constructor(fieldPath: string, message: string) {
    super(
      "CONFIGURATION_ERROR",
      `Configuration error at '${fieldPath}': ${message}`,
      false,
      `Field path: ${fieldPath}`
    );
    this.fieldPath = fieldPath;
  }
}

/**
 * Thrown when a configuration URL cannot be fetched.
 */
export class ConfigFetchError extends AppBaseError {
  public readonly url: string;
  public readonly statusCode?: number;

  constructor(url: string, message: string, statusCode?: number) {
    super("CONFIG_FETCH_ERROR", message, true, `URL: ${url}, Status: ${statusCode ?? "N/A"}`);
    this.url = url;
    this.statusCode = statusCode;
  }
}

/**
 * Thrown when an API call fails.
 */
export class ApiExecutionError extends AppBaseError {
  public readonly apiId: string;
  public readonly url: string;
  public readonly statusCode?: number;

  constructor(
    apiId: string,
    url: string,
    message: string,
    retryable: boolean,
    statusCode?: number
  ) {
    super("API_EXECUTION_ERROR", message, retryable, `API: ${apiId}, URL: ${url}`);
    this.apiId = apiId;
    this.url = url;
    this.statusCode = statusCode;
  }
}

/**
 * Thrown when Office.js document operations fail.
 */
export class OfficeApiError extends AppBaseError {
  public readonly operation: string;

  constructor(operation: string, message: string) {
    super("OFFICE_API_ERROR", message, true, `Operation: ${operation}`);
    this.operation = operation;
  }
}

/**
 * Thrown when IndexedDB operations fail.
 */
export class HistoryStorageError extends AppBaseError {
  public readonly operation: string;

  constructor(operation: string, message: string) {
    super("HISTORY_STORAGE_ERROR", message, false, `Operation: ${operation}`);
    this.operation = operation;
  }
}
```

---

## 4. Component Architecture (React)

### 4.1 Component Tree

```
FluentProvider (theme: webLightTheme)
└── ErrorBoundary
    └── AppProvider (React Context + useReducer)
        └── App
            ├── ConfigLoader
            │   ├── Input (Fluent UI) -- config URL
            │   ├── Button "Load"
            │   └── Button "Reload" (icon only)
            ├── StatusBar
            │   └── MessageBar (Fluent UI) -- error/success/loading messages
            ├── TabList (Fluent UI)
            │   ├── Tab "APIs"
            │   └── Tab "History"
            ├── [Tab: APIs] Sidebar
            │   └── ApiTree
            │       └── ApiGroup (recursive)
            │           ├── Accordion / AccordionItem (Fluent UI)
            │           ├── ApiGroup (nested children, recursive)
            │           └── ApiItem
            │               ├── Card (Fluent UI) -- API name + description
            │               ├── PromptInput (Textarea)
            │               ├── ActionButtons
            │               │   ├── Button "{name} (Selection)"
            │               │   └── Button "{name} (Full Doc)" (conditional)
            │               └── ResponsePanel
            │                   ├── Text (response content, scrollable)
            │                   ├── Button "Copy"
            │                   └── Button "Insert into Document"
            └── [Tab: History] HistoryPanel
                ├── Button "Clear History"
                └── HistoryEntry[] (scrollable list)
                    └── HistoryEntry
                        ├── Card (timestamp, API name, prompt preview)
                        ├── [Expanded] Full details (prompt, text, response)
                        ├── Button "Replay"
                        └── Button "Delete"
```

### 4.2 Component Specifications

#### App (`src/taskpane/App.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | None (root component) |
| **State** | Consumed from `AppContext` via `useAppState()` |
| **Responsibilities** | Top-level layout, tab navigation, orchestration |
| **Children** | ConfigLoader, StatusBar, TabList, Sidebar, HistoryPanel |

```typescript
const App: React.FC = () => {
  const { state, dispatch } = useAppState();
  // Renders ConfigLoader at top, TabList for navigation,
  // conditionally renders Sidebar or HistoryPanel based on activeTab
};
```

#### ConfigLoader (`src/taskpane/components/ConfigLoader.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `onLoad: (url: string) => void`, `onReload: () => void`, `isLoading: boolean`, `currentUrl: string` |
| **Local State** | `inputUrl: string` (controlled input) |
| **Responsibilities** | URL input, load/reload triggers |
| **Fluent UI** | `Input`, `Button`, `Spinner` |

#### StatusBar (`src/taskpane/components/StatusBar.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `error: AppError | null`, `configTitle: string | null`, `isLoading: boolean` |
| **Responsibilities** | Display current status: loading indicator, error banner, or config title |
| **Fluent UI** | `MessageBar`, `MessageBarBody`, `Spinner` |

#### ApiTree (`src/taskpane/components/ApiTree.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `groups: readonly ApiGroup[]` |
| **Responsibilities** | Render top-level groups, delegate recursion to ApiGroup |
| **Fluent UI** | `Accordion` |

#### ApiGroup (`src/taskpane/components/ApiGroup.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `group: ApiGroup`, `depth: number` |
| **Responsibilities** | Render group as collapsible accordion, recursively render children, render API items |
| **Fluent UI** | `AccordionItem`, `AccordionHeader`, `AccordionPanel` |

```typescript
interface ApiGroupProps {
  readonly group: ApiGroup;
  readonly depth: number; // for indentation control
}
```

#### ApiItem (`src/taskpane/components/ApiItem.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `api: ApiCallConfig`, `onExecute: (api: ApiCallConfig, textSource: TextSource) => void`, `isExecuting: boolean`, `response: ApiResponse | null`, `hasSelection: boolean` |
| **Local State** | `prompt: string` (initialized from `api.promptTemplate`) |
| **Responsibilities** | Display API name, prompt editor, action buttons, response panel |
| **Children** | PromptInput, ActionButtons, ResponsePanel |

#### PromptInput (`src/taskpane/components/PromptInput.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `value: string`, `onChange: (value: string) => void`, `placeholder: string`, `disabled: boolean` |
| **Responsibilities** | Editable textarea for prompt text |
| **Fluent UI** | `Textarea` |

#### ActionButtons (`src/taskpane/components/ActionButtons.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `apiName: string`, `inputMode: InputMode`, `onExecute: (textSource: TextSource) => void`, `isExecuting: boolean`, `hasSelection: boolean` |
| **Responsibilities** | Render 1 or 2 buttons based on inputMode, disable during execution |
| **Fluent UI** | `Button`, `Spinner` |

**Button rendering logic:**

```typescript
interface ActionButtonsProps {
  readonly apiName: string;
  readonly inputMode: InputMode;
  readonly onExecute: (textSource: TextSource) => void;
  readonly isExecuting: boolean;
  readonly hasSelection: boolean;
}

// Rendering rules:
// inputMode === "selected" -> 1 button: "{apiName} (Selection)", disabled if !hasSelection
// inputMode === "full"     -> 1 button: "{apiName} (Full Doc)"
// inputMode === "both"     -> 2 buttons: both of the above
```

#### ResponsePanel (`src/taskpane/components/ResponsePanel.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `response: ApiResponse | null`, `onCopy: () => void`, `onInsert: () => void` |
| **Responsibilities** | Display API response or error, copy/insert actions |
| **Fluent UI** | `Card`, `Text`, `Button`, `MessageBar` |

#### HistoryPanel (`src/taskpane/components/HistoryPanel.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `entries: readonly HistoryEntry[]`, `onReplay: (entry: HistoryEntry) => void`, `onDelete: (id: string) => void`, `onClear: () => void`, `isLoading: boolean` |
| **Responsibilities** | Scrollable history list, clear all action |
| **Fluent UI** | `Button`, `Dialog` (confirmation), scrollable container |

#### HistoryEntry (`src/taskpane/components/HistoryEntry.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `entry: HistoryEntry`, `onReplay: (entry: HistoryEntry) => void`, `onDelete: (id: string) => void` |
| **Local State** | `isExpanded: boolean` |
| **Responsibilities** | Expandable card showing history details, replay/delete actions |
| **Fluent UI** | `Card`, `Text`, `Button`, `Badge` |

#### ErrorBoundary (`src/taskpane/components/ErrorBoundary.tsx`)

| Aspect | Detail |
|--------|--------|
| **Props** | `children: React.ReactNode`, `fallback?: React.ReactNode` |
| **State** | `hasError: boolean`, `error: Error | null` |
| **Responsibilities** | Catch unhandled React rendering errors, display recovery UI |

### 4.3 State Management

The application uses **React Context + `useReducer`** for global state management. This provides a single source of truth without the complexity of external state management libraries.

#### AppContext Provider (`src/taskpane/context/AppContext.tsx`)

```typescript
import { AppState, AppAction } from "../types/state";

interface AppContextValue {
  readonly state: AppState;
  readonly dispatch: React.Dispatch<AppAction>;
}

const AppContext = React.createContext<AppContextValue | undefined>(undefined);

/**
 * Custom hook to access application state. Throws if used outside AppProvider.
 */
export function useAppState(): AppContextValue {
  const context = React.useContext(AppContext);
  if (context === undefined) {
    throw new Error("useAppState must be used within an AppProvider");
  }
  return context;
}
```

#### Application Reducer (`src/taskpane/context/appReducer.ts`)

```typescript
export function appReducer(state: AppState, action: AppAction): AppState {
  switch (action.type) {
    case "SET_ACTIVE_TAB":
      return { ...state, activeTab: action.payload };
    case "SET_CONFIG_URL":
      return { ...state, configUrl: action.payload };
    case "CONFIG_LOAD_START":
      return { ...state, config: { status: "loading", data: null, error: null } };
    case "CONFIG_LOAD_SUCCESS":
      return {
        ...state,
        config: { status: "success", data: action.payload, error: null },
        selectedApi: null,
        currentPrompt: "",
      };
    case "CONFIG_LOAD_ERROR":
      return { ...state, config: { status: "error", data: null, error: action.payload } };
    case "SELECT_API":
      return {
        ...state,
        selectedApi: action.payload,
        currentPrompt: action.payload?.promptTemplate ?? "",
        execution: { status: "idle", data: null, error: null },
      };
    case "SET_PROMPT":
      return { ...state, currentPrompt: action.payload };
    case "EXECUTION_START":
      return { ...state, execution: { status: "loading", data: null, error: null } };
    case "EXECUTION_SUCCESS":
      return { ...state, execution: { status: "success", data: action.payload, error: null } };
    case "EXECUTION_ERROR":
      return { ...state, execution: { status: "error", data: null, error: action.payload } };
    // ... history and document actions follow the same pattern
    default:
      return state;
  }
}
```

#### Initial State

```typescript
export const initialAppState: AppState = {
  activeTab: "apis",
  config: { status: "idle", data: null, error: null },
  configUrl: "",
  selectedApi: null,
  currentPrompt: "",
  execution: { status: "idle", data: null, error: null },
  history: { status: "idle", data: null, error: null },
  documentInfo: { status: "idle", data: null, error: null },
};
```

---

## 5. Service Layer

### 5.1 ConfigService (`src/taskpane/services/ConfigService.ts`)

**Responsibility:** Fetch, validate, and cache the remote configuration JSON.

```typescript
export class ConfigService {
  private static readonly CACHE_TTL_MS = 3600000; // 1 hour
  private static readonly STORAGE_KEY_PREFIX = "word-addin-sidebar";

  /**
   * Load configuration from the given URL.
   * Returns cached configuration if the cache is valid and not expired.
   * @throws ConfigFetchError if the URL is unreachable or returns non-200.
   * @throws ConfigurationError if the configuration fails validation.
   */
  async loadConfig(url: string): Promise<AddinConfiguration>;

  /**
   * Bypass cache and fetch fresh configuration from the URL.
   * @throws ConfigFetchError if the URL is unreachable or returns non-200.
   * @throws ConfigurationError if the configuration fails validation.
   */
  async reloadConfig(url: string): Promise<AddinConfiguration>;

  /**
   * Validate the configuration object against the JSON schema.
   * @throws ConfigurationError with the field path for any validation failure.
   * No fallback values are ever substituted.
   */
  validateConfig(config: unknown): asserts config is AddinConfiguration;

  /**
   * Persist the last-used configuration URL to localStorage.
   */
  saveConfigUrl(url: string): void;

  /**
   * Retrieve the last-used configuration URL from localStorage.
   * Returns null if no URL was previously saved.
   */
  getConfigUrl(): string | null;

  /**
   * Cache configuration to localStorage with TTL metadata.
   */
  private cacheConfig(url: string, config: AddinConfiguration): void;

  /**
   * Retrieve cached configuration if TTL has not expired.
   * Returns null if no cache exists or if the cache is expired.
   */
  private getCachedConfig(url: string): AddinConfiguration | null;

  /**
   * Generate a localStorage key with partition key prefix for Office web isolation.
   */
  private getStorageKey(suffix: string): string;
}
```

**Validation implementation notes:**

- The `validateConfig` method performs a recursive, depth-first traversal of the configuration object.
- For each required field, if the field is missing or empty, a `ConfigurationError` is thrown with the full dot-notation path (e.g., `"groups[0].apis[1].url"`).
- URL fields are validated with a regex check for `https://` prefix.
- `method` values are checked against `["GET", "POST"]`.
- `inputMode` values are checked against `["selected", "full", "both"]`.
- Groups are validated to have at least one of `children` or `apis`.
- No `try/catch` silences validation errors; they propagate to the caller.

### 5.2 OfficeApiService (`src/taskpane/services/OfficeApiService.ts`)

**Responsibility:** Interact with the Word document through Office.js.

```typescript
export class OfficeApiService {
  /**
   * Extract the currently selected text from the active Word document.
   * Returns an empty string if the cursor has no selection.
   * @throws OfficeApiError if the Word API call fails.
   */
  async getSelectedText(): Promise<string>;

  /**
   * Extract the full text content of the active Word document body.
   * @throws OfficeApiError if the Word API call fails.
   */
  async getFullDocumentText(): Promise<string>;

  /**
   * Get the active document filename.
   * @throws OfficeApiError if the document properties cannot be read.
   */
  async getDocumentName(): Promise<string>;

  /**
   * Get complete document text information including selection status.
   * @throws OfficeApiError if any Word API call fails.
   */
  async getDocumentTextInfo(): Promise<DocumentTextInfo>;

  /**
   * Insert text at the current cursor position, replacing any selection.
   * @throws OfficeApiError if the insert operation fails.
   */
  async insertAtCursor(text: string): Promise<void>;

  /**
   * Append text at the end of the document body.
   * @throws OfficeApiError if the append operation fails.
   */
  async appendToDocument(text: string): Promise<void>;
}
```

**Implementation pattern (Office.js proxy model):**

```typescript
async getSelectedText(): Promise<string> {
  try {
    return await Word.run(async (context: Word.RequestContext) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text;
    });
  } catch (error) {
    throw new OfficeApiError(
      "getSelectedText",
      `Failed to extract selected text: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
```

### 5.3 ApiExecutionService (`src/taskpane/services/ApiExecutionService.ts`)

**Responsibility:** Build HTTP requests from API configurations and templates, execute them, and parse responses.

```typescript
export class ApiExecutionService {
  /**
   * Execute an API call using the given configuration and payload.
   * Resolves all placeholders in the body template and prompt.
   * @returns ApiResponse (success or error, never throws for HTTP errors).
   */
  async execute(
    apiConfig: ApiCallConfig,
    payload: ApiRequestPayload
  ): Promise<ApiResponse>;

  /**
   * Build the fully resolved HTTP request from configuration and payload.
   * Replaces all {{prompt}}, {{text}}, and {{documentName}} placeholders.
   */
  buildRequest(
    apiConfig: ApiCallConfig,
    payload: ApiRequestPayload
  ): ConstructedRequest;

  /**
   * Resolve all placeholder strings in a body template object (recursive).
   * Traverses objects and arrays, replacing placeholders in string values.
   */
  private resolveBodyTemplate(
    template: Readonly<Record<string, unknown>>,
    payload: ApiRequestPayload
  ): Record<string, unknown>;

  /**
   * Replace placeholder tokens in a string.
   */
  private resolvePlaceholders(
    template: string,
    payload: ApiRequestPayload
  ): string;

  /**
   * Extract the response text using the dot-notation responseField path.
   * Supports array indexing (e.g., "choices.0.message.content").
   * @throws if the path does not resolve to a string value.
   */
  private extractResponseField(body: unknown, fieldPath: string): string;

  /**
   * Classify fetch errors into ApiErrorType categories.
   */
  private classifyError(error: unknown): ApiErrorType;
}
```

**Placeholder resolution implementation:**

```typescript
private resolvePlaceholders(template: string, payload: ApiRequestPayload): string {
  return template
    .replace(/\{\{prompt\}\}/g, payload.prompt)
    .replace(/\{\{text\}\}/g, payload.text)
    .replace(/\{\{documentName\}\}/g, payload.documentName);
}

private resolveBodyTemplate(
  template: Readonly<Record<string, unknown>>,
  payload: ApiRequestPayload
): Record<string, unknown> {
  const resolved: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(template)) {
    if (typeof value === "string") {
      resolved[key] = this.resolvePlaceholders(value, payload);
    } else if (Array.isArray(value)) {
      resolved[key] = value.map((item) =>
        typeof item === "object" && item !== null
          ? this.resolveBodyTemplate(item as Record<string, unknown>, payload)
          : typeof item === "string"
            ? this.resolvePlaceholders(item, payload)
            : item
      );
    } else if (typeof value === "object" && value !== null) {
      resolved[key] = this.resolveBodyTemplate(
        value as Record<string, unknown>,
        payload
      );
    } else {
      resolved[key] = value;
    }
  }
  return resolved;
}
```

**Error classification:**

```typescript
private classifyError(error: unknown): ApiErrorType {
  if (error instanceof TypeError && (error.message.includes("Failed to fetch") ||
      error.message.includes("NetworkError"))) {
    // CORS failures in browsers manifest as TypeErrors with "Failed to fetch"
    // Distinguishing CORS from network errors is not reliably possible,
    // so we use a heuristic based on the error message.
    return "CORS_ERROR";
  }
  if (error instanceof DOMException && error.name === "AbortError") {
    return "TIMEOUT_ERROR";
  }
  return "NETWORK_ERROR";
}
```

### 5.4 HistoryService (`src/taskpane/services/HistoryService.ts`)

**Responsibility:** CRUD operations on IndexedDB for persistent history storage.

```typescript
import { openDB, IDBPDatabase } from "idb";

export class HistoryService {
  private static readonly DB_NAME = "word-addin-sidebar-history";
  private static readonly DB_VERSION = 1;
  private static readonly STORE_NAME = "history";
  private static readonly MAX_ENTRIES = 1000;
  private static readonly MAX_AGE_DAYS = 90;

  private db: IDBPDatabase | null = null;

  /**
   * Initialize the IndexedDB database connection.
   * Creates the object store and indexes if they do not exist.
   * @throws HistoryStorageError if IndexedDB is unavailable or initialization fails.
   */
  async initialize(): Promise<void>;

  /**
   * Add a new history entry.
   * Triggers auto-pruning if entry count exceeds MAX_ENTRIES.
   * @throws HistoryStorageError if the write operation fails.
   */
  async addEntry(entry: HistoryEntry): Promise<void>;

  /**
   * Retrieve history entries ordered by timestamp descending.
   * Applies optional filters.
   * @throws HistoryStorageError if the read operation fails.
   */
  async getEntries(filter?: HistoryFilter): Promise<readonly HistoryEntry[]>;

  /**
   * Delete a single history entry by ID.
   * @throws HistoryStorageError if the delete operation fails.
   */
  async deleteEntry(id: string): Promise<void>;

  /**
   * Delete all history entries.
   * @throws HistoryStorageError if the clear operation fails.
   */
  async clearAll(): Promise<void>;

  /**
   * Remove entries older than MAX_AGE_DAYS and entries exceeding MAX_ENTRIES.
   * Called automatically after addEntry when thresholds are exceeded.
   */
  private async prune(): Promise<void>;

  /**
   * Ensure the database connection is initialized.
   * @throws HistoryStorageError if not initialized.
   */
  private ensureDb(): IDBPDatabase;
}
```

**IndexedDB schema setup:**

```typescript
async initialize(): Promise<void> {
  try {
    this.db = await openDB(HistoryService.DB_NAME, HistoryService.DB_VERSION, {
      upgrade(db) {
        const store = db.createObjectStore(HistoryService.STORE_NAME, {
          keyPath: "id",
        });
        store.createIndex("byTimestamp", "timestamp");
        store.createIndex("byApiId", "apiId");
      },
    });
  } catch (error) {
    throw new HistoryStorageError(
      "initialize",
      `Failed to initialize IndexedDB: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
```

---

## 6. Data Flow

### 6.1 Application Initialization Flow

```
1. taskpane.html loads
2. index.tsx calls Office.onReady()
3. Office.onReady resolves -> React app mounts
4. FluentProvider wraps the app with webLightTheme
5. AppProvider initializes:
   a. Read persisted configUrl from localStorage
   b. If configUrl exists:
      i.  Dispatch CONFIG_LOAD_START
      ii. ConfigService.loadConfig(url) -- checks cache first
      iii. On success: dispatch CONFIG_LOAD_SUCCESS
      iv. On failure: dispatch CONFIG_LOAD_ERROR
   c. Initialize HistoryService (IndexedDB)
   d. Load recent history entries -> dispatch HISTORY_LOAD_SUCCESS
6. App renders with restored state
```

### 6.2 API Execution Flow (Primary User Journey)

```
 User Action                  System Response
 -----------                  ---------------
 1. Select API in tree    --> dispatch SELECT_API
                              PromptInput populated with promptTemplate
                              ActionButtons rendered based on inputMode

 2. Edit prompt (optional) -> dispatch SET_PROMPT

 3. Click action button    --> dispatch EXECUTION_START
    (e.g., "Summarize         |
     (Selection)")             |
                               v
                          4. OfficeApiService.getSelectedText()
                             or getFullDocumentText()
                               |
                               v
                          5. OfficeApiService.getDocumentName()
                               |
                               v
                          6. Build ApiRequestPayload:
                             {
                               prompt: <user-edited prompt>,
                               text: <extracted text>,
                               documentName: <filename>,
                               textSource: "selected" | "full"
                             }
                               |
                               v
                          7. ApiExecutionService.execute(apiConfig, payload)
                             a. buildRequest() -- resolve placeholders
                             b. fetch() with timeout via AbortController
                             c. Parse response JSON
                             d. Extract responseField
                               |
                               v
                          8. On success:
                             a. dispatch EXECUTION_SUCCESS
                             b. ResponsePanel displays extracted text
                             c. Create HistoryEntry
                             d. HistoryService.addEntry(entry)
                             e. dispatch HISTORY_ENTRY_ADDED

                          8. On failure:
                             a. dispatch EXECUTION_ERROR
                             b. ResponsePanel displays error with retry option
                             c. Create HistoryEntry (wasSuccessful: false)
                             d. HistoryService.addEntry(entry)
                             e. dispatch HISTORY_ENTRY_ADDED
```

### 6.3 History Replay Flow

```
 User Action                  System Response
 -----------                  ---------------
 1. Click "Replay"        --> Read HistoryEntry
    on history entry           |
                               v
                          2. Find matching ApiCallConfig by apiId
                             in current configuration
                               |
                               v
                          3a. If found:
                              dispatch SELECT_API(matchedConfig)
                              dispatch SET_PROMPT(entry.prompt)
                              dispatch SET_ACTIVE_TAB("apis")
                              -> User sees the API selected with
                                 the historical prompt pre-filled

                          3b. If not found (config changed):
                              Show warning: "This API is no longer
                              available in the current configuration"
```

### 6.4 Configuration Reload Flow

```
 User Action                  System Response
 -----------                  ---------------
 1. Click "Reload"        --> dispatch CONFIG_LOAD_START
                               |
                               v
                          2. ConfigService.reloadConfig(url)
                             (bypasses cache, fetches fresh)
                               |
                               v
                          3. validateConfig(response)
                               |
                               v
                          4a. Valid:
                              dispatch CONFIG_LOAD_SUCCESS
                              Cache updated in localStorage
                              UI re-renders with new config

                          4b. Invalid:
                              throw ConfigurationError
                              dispatch CONFIG_LOAD_ERROR
                              StatusBar shows error with field path
```

---

## 7. Error Handling

### 7.1 Error Handling Principles

1. **No fallback values for configuration.** Missing required fields throw `ConfigurationError` with the field path. No defaults are ever substituted. This is a strict project policy.
2. **All errors are typed.** Every error is an instance of a specific error class extending `AppBaseError`.
3. **Errors are user-visible.** Every error condition produces a visible message in the StatusBar or ResponsePanel.
4. **Retryable vs. non-retryable.** Each error type declares whether the operation can be retried. The UI shows a "Retry" button only for retryable errors.
5. **Error boundaries prevent crashes.** The React `ErrorBoundary` catches unhandled rendering errors and displays a recovery UI.

### 7.2 Error Matrix

| Error Condition | Error Class | Retryable | User Message | UI Location |
|----------------|-------------|-----------|--------------|-------------|
| Config URL unreachable | `ConfigFetchError` | Yes | "Cannot reach configuration URL. Check your network connection and try again." | StatusBar |
| Config URL returns 4xx | `ConfigFetchError` | No | "Configuration URL returned HTTP {status}. Verify the URL is correct." | StatusBar |
| Config URL returns 5xx | `ConfigFetchError` | Yes | "Configuration server error (HTTP {status}). Try reloading later." | StatusBar |
| Config JSON malformed | `ConfigurationError` | No | "Configuration file contains invalid JSON." | StatusBar |
| Config missing required field | `ConfigurationError` | No | "Configuration error at '{fieldPath}': {message}" | StatusBar |
| Config invalid URL (non-HTTPS) | `ConfigurationError` | No | "Configuration error at '{fieldPath}': URL must use HTTPS." | StatusBar |
| Config invalid inputMode | `ConfigurationError` | No | "Configuration error at '{fieldPath}': inputMode must be 'selected', 'full', or 'both'." | StatusBar |
| API call network error | `ApiExecutionError` | Yes | "Network error calling {apiName}. Check your connection." | ResponsePanel |
| API call CORS error | `ApiExecutionError` | No | "CORS error calling {apiName}. The API server must allow cross-origin requests from Office add-ins." | ResponsePanel |
| API call timeout | `ApiExecutionError` | Yes | "Request to {apiName} timed out after {timeout}ms." | ResponsePanel |
| API call HTTP 4xx | `ApiExecutionError` | No | "{apiName} returned HTTP {status}: {message}" | ResponsePanel |
| API call HTTP 5xx | `ApiExecutionError` | Yes | "{apiName} server error (HTTP {status}). Try again later." | ResponsePanel |
| API response parse error | `ApiExecutionError` | No | "Failed to parse response from {apiName}." | ResponsePanel |
| Response field not found | `ApiExecutionError` | No | "Response from {apiName} does not contain field '{responseField}'." | ResponsePanel |
| Word API selection error | `OfficeApiError` | Yes | "Failed to read selected text. Make sure the document is open and try again." | StatusBar |
| Word API body read error | `OfficeApiError` | Yes | "Failed to read document text. The document may be too large or locked." | StatusBar |
| IndexedDB init failure | `HistoryStorageError` | No | "History storage unavailable. History will not be persisted this session." | StatusBar (warning) |
| IndexedDB write failure | `HistoryStorageError` | No | "Failed to save history entry." | StatusBar (warning) |

### 7.3 Error Boundary Implementation

```typescript
interface ErrorBoundaryProps {
  readonly children: React.ReactNode;
  readonly fallback?: React.ReactNode;
}

interface ErrorBoundaryState {
  readonly hasError: boolean;
  readonly error: Error | null;
}

class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  state: ErrorBoundaryState = { hasError: false, error: null };

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    console.error("ErrorBoundary caught:", error, errorInfo);
  }

  private handleReset = (): void => {
    this.setState({ hasError: false, error: null });
  };

  render(): React.ReactNode {
    if (this.state.hasError) {
      return this.props.fallback ?? (
        <div>
          <MessageBar intent="error">
            <MessageBarBody>
              An unexpected error occurred: {this.state.error?.message}
            </MessageBarBody>
          </MessageBar>
          <Button onClick={this.handleReset}>Try Again</Button>
        </div>
      );
    }
    return this.props.children;
  }
}
```

---

## 8. File Structure

```
office-addin/
├── docs/
│   ├── design/
│   │   ├── plan-001-word-addin-sidebar.md
│   │   ├── project-design.md                    # This document
│   │   ├── project-functions.md
│   │   └── configuration-guide.md               # Phase 8
│   └── reference/
│       ├── investigation-word-addin-sidebar.md
│       └── example-config.json                   # Phase 8
├── test_scripts/
│   ├── configService.test.ts                     # Phase 7
│   ├── apiService.test.ts                        # Phase 7
│   ├── historyService.test.ts                    # Phase 7
│   ├── configValidation.test.ts                  # Phase 7
│   ├── mockApiServer.ts                          # Phase 7
│   ├── sample-config.json                        # Phase 7
│   └── manual-test-plan.md                       # Phase 7
├── assets/
│   ├── icon-16.png                               # Phase 8
│   ├── icon-32.png                               # Phase 8
│   └── icon-80.png                               # Phase 8
├── src/
│   └── taskpane/
│       ├── index.tsx                             # React entry point + Office.onReady
│       ├── App.tsx                               # Root component with tab navigation
│       ├── taskpane.html                         # HTML entry point
│       ├── taskpane.css                          # Global responsive styles
│       ├── types/
│       │   ├── config.ts                         # AddinConfiguration, ApiGroup, ApiCallConfig
│       │   ├── api.ts                            # ApiRequestPayload, ApiResponse, DocumentTextInfo
│       │   ├── history.ts                        # HistoryEntry, HistoryFilter, HistoryDbSchema
│       │   ├── state.ts                          # AppState, AppAction, AsyncState, ActiveTab
│       │   ├── errors.ts                         # Error classes: ConfigurationError, ApiExecutionError, etc.
│       │   └── config.schema.json                # JSON Schema for configuration validation
│       ├── context/
│       │   ├── AppContext.tsx                     # React Context provider + useAppState hook
│       │   └── appReducer.ts                     # Application state reducer
│       ├── services/
│       │   ├── ConfigService.ts                  # Fetch, validate, cache configuration
│       │   ├── OfficeApiService.ts               # Word document interaction (Office.js)
│       │   ├── ApiExecutionService.ts            # HTTP request construction and execution
│       │   └── HistoryService.ts                 # IndexedDB CRUD for history
│       ├── hooks/
│       │   ├── useConfig.ts                      # Configuration loading state management
│       │   ├── useHistory.ts                     # History CRUD operations
│       │   ├── useDocumentText.ts                # Document text extraction
│       │   └── useApiExecution.ts                # API call execution state
│       └── components/
│           ├── ConfigLoader.tsx                  # URL input + Load/Reload buttons
│           ├── StatusBar.tsx                     # Loading/error/success message bar
│           ├── ApiTree.tsx                       # Top-level accordion for API groups
│           ├── ApiGroup.tsx                      # Recursive group rendering
│           ├── ApiItem.tsx                       # Single API call card with prompt/buttons/response
│           ├── PromptInput.tsx                   # Editable textarea for prompts
│           ├── ActionButtons.tsx                 # Dynamic 1-or-2 button rendering
│           ├── ResponsePanel.tsx                 # API response display with copy/insert
│           ├── HistoryPanel.tsx                  # History list with clear action
│           ├── HistoryEntry.tsx                  # Expandable history entry card
│           └── ErrorBoundary.tsx                 # React error boundary wrapper
├── manifest.xml                                  # Office Add-in XML manifest
├── webpack.config.js                             # Webpack build configuration
├── tsconfig.json                                 # TypeScript configuration (strict mode)
├── package.json                                  # Dependencies and scripts
├── CLAUDE.md                                     # Tool documentation
└── Issues - Pending Items.md                     # Issues and pending items tracker
```

---

## 9. Parallel Implementation Units

### 9.1 Implementation Units and Dependencies

The project decomposes into **10 implementation units** across 4 dependency tiers. Units within the same tier can be built in parallel.

```
Tier 0 (Foundation)
  [U0] Project Scaffolding + Manifest
    |
Tier 1 (Types -- no runtime dependencies, all parallel)
  [U1] Type Definitions (config.ts, api.ts, history.ts, state.ts, errors.ts)
    |
Tier 2 (Services -- depend on types, all parallel within tier)
  [U2a] ConfigService        [U2b] OfficeApiService
  [U2c] ApiExecutionService  [U2d] HistoryService
    |
Tier 3 (UI + Hooks -- depend on services)
  [U3a] Context + Reducer    [U3b] Custom Hooks (all 4 in parallel)
    |                           |
  [U3c] Layout Components    [U3d] API Components    [U3e] History Components
    |                           |                        |
Tier 4 (Integration)
  [U4] Integration, Wiring, Responsive Styling, ErrorBoundary
```

### 9.2 Unit Details and Interface Contracts

#### Unit U0: Project Scaffolding

| Aspect | Detail |
|--------|--------|
| **Phase** | 1 |
| **Depends on** | Nothing |
| **Produces** | Project skeleton, manifest.xml, build pipeline |
| **Contract** | `npm start` launches Word with sideloaded add-in |

#### Unit U1: Type Definitions

| Aspect | Detail |
|--------|--------|
| **Phase** | 2 |
| **Depends on** | U0 |
| **Produces** | All files in `src/taskpane/types/` |
| **Contract** | All interfaces and types compile with `tsc --noEmit`. Exported from barrel file. |

#### Unit U2a: ConfigService

| Aspect | Detail |
|--------|--------|
| **Phase** | 3B |
| **Depends on** | U1 (types) |
| **Produces** | `src/taskpane/services/ConfigService.ts` |
| **Interface Contract** | Implements `loadConfig(url: string): Promise<AddinConfiguration>`, `reloadConfig(url: string): Promise<AddinConfiguration>`, `validateConfig(config: unknown): asserts config is AddinConfiguration`, `saveConfigUrl(url: string): void`, `getConfigUrl(): string \| null` |
| **Error Contract** | Throws `ConfigFetchError` for network/HTTP errors, `ConfigurationError` for validation failures |

#### Unit U2b: OfficeApiService

| Aspect | Detail |
|--------|--------|
| **Phase** | 3A |
| **Depends on** | U1 (types) |
| **Produces** | `src/taskpane/services/OfficeApiService.ts` |
| **Interface Contract** | Implements `getSelectedText(): Promise<string>`, `getFullDocumentText(): Promise<string>`, `getDocumentName(): Promise<string>`, `getDocumentTextInfo(): Promise<DocumentTextInfo>`, `insertAtCursor(text: string): Promise<void>`, `appendToDocument(text: string): Promise<void>` |
| **Error Contract** | Throws `OfficeApiError` for all Word API failures |
| **External Dependency** | `Office.js` (available at runtime in Word) |

#### Unit U2c: ApiExecutionService

| Aspect | Detail |
|--------|--------|
| **Phase** | 3C |
| **Depends on** | U1 (types) |
| **Produces** | `src/taskpane/services/ApiExecutionService.ts` |
| **Interface Contract** | Implements `execute(apiConfig: ApiCallConfig, payload: ApiRequestPayload): Promise<ApiResponse>`, `buildRequest(apiConfig: ApiCallConfig, payload: ApiRequestPayload): ConstructedRequest` |
| **Error Contract** | Returns `ApiErrorResponse` for all failures (does not throw for HTTP errors). Returns `ApiSuccessResponse` for successful calls. |

#### Unit U2d: HistoryService

| Aspect | Detail |
|--------|--------|
| **Phase** | 3D |
| **Depends on** | U1 (types) |
| **Produces** | `src/taskpane/services/HistoryService.ts` |
| **Interface Contract** | Implements `initialize(): Promise<void>`, `addEntry(entry: HistoryEntry): Promise<void>`, `getEntries(filter?: HistoryFilter): Promise<readonly HistoryEntry[]>`, `deleteEntry(id: string): Promise<void>`, `clearAll(): Promise<void>` |
| **Error Contract** | Throws `HistoryStorageError` for all IndexedDB failures |
| **External Dependency** | `idb` library |

#### Unit U3a: Context + Reducer

| Aspect | Detail |
|--------|--------|
| **Phase** | 4 (partial) |
| **Depends on** | U1 (types) |
| **Produces** | `src/taskpane/context/AppContext.tsx`, `src/taskpane/context/appReducer.ts` |
| **Interface Contract** | Exports `AppProvider`, `useAppState()` hook, `appReducer`, `initialAppState` |
| **Can start with** | U1 only (does not depend on services) |

#### Unit U3b: Custom Hooks

| Aspect | Detail |
|--------|--------|
| **Phase** | 4 |
| **Depends on** | U3a (context), U2a-U2d (services) |
| **Produces** | All files in `src/taskpane/hooks/` |
| **Interface Contract** | Each hook wraps a service, manages loading/error/success states, and dispatches actions to the reducer |

#### Unit U3c: Layout Components

| Aspect | Detail |
|--------|--------|
| **Phase** | 5A |
| **Depends on** | U3a (context) |
| **Produces** | `App.tsx`, `ConfigLoader.tsx`, `StatusBar.tsx`, `index.tsx`, `taskpane.html` |

#### Unit U3d: API Components

| Aspect | Detail |
|--------|--------|
| **Phase** | 5B |
| **Depends on** | U3c (layout), U3b (hooks) |
| **Produces** | `ApiTree.tsx`, `ApiGroup.tsx`, `ApiItem.tsx`, `PromptInput.tsx`, `ActionButtons.tsx`, `ResponsePanel.tsx` |

#### Unit U3e: History Components

| Aspect | Detail |
|--------|--------|
| **Phase** | 5C |
| **Depends on** | U3c (layout), U3b (hooks) |
| **Produces** | `HistoryPanel.tsx`, `HistoryEntry.tsx` |
| **Can run parallel with** | U3d |

#### Unit U4: Integration

| Aspect | Detail |
|--------|--------|
| **Phase** | 6 |
| **Depends on** | U3c, U3d, U3e |
| **Produces** | `ErrorBoundary.tsx`, `taskpane.css`, wiring updates across all components |

### 9.3 Parallel Execution Summary

| Parallel Track | Units | Can Start After |
|---------------|-------|-----------------|
| Track A | U2a (ConfigService) | U1 complete |
| Track B | U2b (OfficeApiService) | U1 complete |
| Track C | U2c (ApiExecutionService) | U1 complete |
| Track D | U2d (HistoryService) | U1 complete |
| Track E | U3a (Context/Reducer) | U1 complete |
| Track F | U3d (API Components) | U3c + U3b complete |
| Track G | U3e (History Components) | U3c + U3b complete |

**Maximum parallelism:** 5 units simultaneously (U2a, U2b, U2c, U2d, U3a after U1 completes).

---

## Appendix A: Key Design Decisions

| # | Decision | Rationale |
|---|----------|-----------|
| D1 | XML manifest over unified JSON | JSON manifest is preview-only for Word; XML is production-ready |
| D2 | React Context + useReducer over Redux/Zustand | Sufficient for the application's state complexity; no external dependency |
| D3 | IndexedDB via `idb` over raw IndexedDB | `idb` provides a Promise-based wrapper that eliminates callback complexity |
| D4 | `fetch()` over Axios | Native API, no additional dependency, sufficient for our needs |
| D5 | Discriminated union for ApiResponse | Enables exhaustive type checking for success/error handling |
| D6 | Readonly types throughout | Prevents accidental mutation of configuration and state objects |
| D7 | Service classes over plain functions | Encapsulates state (database connections, cache) and enables testing via dependency injection |
| D8 | Body template with recursive resolution | Supports arbitrarily nested JSON structures for diverse API formats |
| D9 | Separate `inputMode` over `inputType` naming | Avoids confusion with HTML input type attribute; "mode" better describes the behavioral semantics |
| D10 | Auto-pruning in HistoryService | Prevents unbounded IndexedDB growth without requiring user intervention |

## Appendix B: External Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `@fluentui/react-components` | ^9.x | Fluent UI React v9 component library |
| `@fluentui/react-icons` | ^2.x | Fluent UI icon set |
| `idb` | ^8.x | Promise-based IndexedDB wrapper |
| `office-addin-mock` | ^2.x | (dev) Office.js mocking for unit tests |
| `fake-indexeddb` | ^5.x | (dev) In-memory IndexedDB for unit tests |

## Appendix C: tsconfig.json Key Settings

```json
{
  "compilerOptions": {
    "strict": true,
    "noImplicitAny": true,
    "strictNullChecks": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noImplicitReturns": true,
    "noFallthroughCasesInSwitch": true,
    "forceConsistentCasingInFileNames": true,
    "esModuleInterop": true,
    "jsx": "react-jsx",
    "target": "ES2020",
    "module": "ESNext",
    "moduleResolution": "node",
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "outDir": "./dist",
    "rootDir": "./src",
    "baseUrl": "./src",
    "paths": {
      "@types/*": ["taskpane/types/*"],
      "@services/*": ["taskpane/services/*"],
      "@hooks/*": ["taskpane/hooks/*"],
      "@components/*": ["taskpane/components/*"],
      "@context/*": ["taskpane/context/*"]
    }
  },
  "include": ["src/**/*.ts", "src/**/*.tsx"],
  "exclude": ["node_modules", "dist", "test_scripts"]
}
```
