import { describe, it, expect } from "vitest";
import {
  appReducer,
  initialAppState,
} from "../src/taskpane/context/appReducer";
import { AppState, AppAction } from "../src/taskpane/types/state";
import { AppBaseError } from "../src/taskpane/types/errors";
import { ApiCallConfig, AddinConfiguration } from "../src/taskpane/types/config";
import { ApiSuccessResponse, ApiErrorResponse } from "../src/taskpane/types/api";
import { HistoryEntry } from "../src/taskpane/types/history";

// ---------------------------------------------------------------------------
// Mock data helpers
// ---------------------------------------------------------------------------

function createMockHistoryEntry(overrides: Partial<HistoryEntry> = {}): HistoryEntry {
  return {
    id: "entry-1",
    timestamp: Date.now(),
    apiId: "api-1",
    apiName: "Test API",
    apiUrl: "https://api.example.com/test",
    prompt: "Summarise the text",
    textSource: "selected",
    textPreview: "Lorem ipsum dolor sit amet...",
    documentName: "Document1.docx",
    wasSuccessful: true,
    responsePreview: "Summary result",
    durationMs: 1200,
    ...overrides,
  };
}

function createMockApiCallConfig(overrides: Partial<ApiCallConfig> = {}): ApiCallConfig {
  return {
    id: "api-1",
    name: "Test API",
    url: "https://api.example.com/test",
    method: "POST",
    inputMode: "selected",
    promptTemplate: "Default prompt: {{prompt}}",
    timeout: 30000,
    ...overrides,
  };
}

function createMockAddinConfiguration(
  overrides: Partial<AddinConfiguration> = {},
): AddinConfiguration {
  return {
    configVersion: "1.0.0",
    name: "Test Configuration",
    description: "A test configuration",
    groups: [
      {
        id: "group-1",
        name: "Test Group",
        apis: [createMockApiCallConfig()],
      },
    ],
    ...overrides,
  };
}

function createMockError(message = "Something went wrong"): AppBaseError {
  return new AppBaseError(message, false, message);
}

function createMockSuccessResponse(): ApiSuccessResponse {
  return {
    kind: "success",
    data: { result: "test" },
    extractedText: "Extracted result text",
    statusCode: 200,
    durationMs: 500,
  };
}

function createMockErrorResponse(): ApiErrorResponse {
  return {
    kind: "error",
    errorType: "network",
    message: "Network error",
    durationMs: 100,
  };
}

/** Deep-clone state so we can verify immutability after dispatch. */
function cloneState(state: AppState): AppState {
  return JSON.parse(JSON.stringify(state));
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("initialAppState", () => {
  it("has correct default values", () => {
    expect(initialAppState.configUrl).toBe("");
    expect(initialAppState.config).toEqual({
      data: null,
      loading: false,
      error: null,
    });
    expect(initialAppState.selectedApi).toBeNull();
    expect(initialAppState.prompt).toBe("");
    expect(initialAppState.execution).toEqual({
      data: null,
      loading: false,
      error: null,
    });
    expect(initialAppState.history).toEqual({
      data: null,
      loading: false,
      error: null,
    });
    expect(initialAppState.activeTab).toBe("apis");
  });
});

describe("appReducer", () => {
  // -----------------------------------------------------------------------
  // SET_CONFIG_URL
  // -----------------------------------------------------------------------
  describe("SET_CONFIG_URL", () => {
    it("updates configUrl", () => {
      const result = appReducer(initialAppState, {
        type: "SET_CONFIG_URL",
        url: "https://example.com/config.json",
      });
      expect(result.configUrl).toBe("https://example.com/config.json");
    });

    it("does not mutate original state", () => {
      const before = cloneState(initialAppState);
      appReducer(initialAppState, {
        type: "SET_CONFIG_URL",
        url: "https://new-url.com",
      });
      expect(initialAppState).toEqual(before);
    });
  });

  // -----------------------------------------------------------------------
  // CONFIG_LOAD_START
  // -----------------------------------------------------------------------
  describe("CONFIG_LOAD_START", () => {
    it("sets loading=true and clears error", () => {
      const stateWithError: AppState = {
        ...initialAppState,
        config: {
          data: null,
          loading: false,
          error: createMockError(),
        },
      };
      const result = appReducer(stateWithError, { type: "CONFIG_LOAD_START" });
      expect(result.config.loading).toBe(true);
      expect(result.config.error).toBeNull();
    });

    it("preserves existing config data", () => {
      const config = createMockAddinConfiguration();
      const stateWithData: AppState = {
        ...initialAppState,
        config: { data: config, loading: false, error: null },
      };
      const result = appReducer(stateWithData, { type: "CONFIG_LOAD_START" });
      expect(result.config.data).toEqual(config);
    });
  });

  // -----------------------------------------------------------------------
  // CONFIG_LOAD_SUCCESS
  // -----------------------------------------------------------------------
  describe("CONFIG_LOAD_SUCCESS", () => {
    it("sets data, loading=false, clears error", () => {
      const config = createMockAddinConfiguration();
      const loadingState: AppState = {
        ...initialAppState,
        config: { data: null, loading: true, error: null },
      };
      const result = appReducer(loadingState, {
        type: "CONFIG_LOAD_SUCCESS",
        config,
      });
      expect(result.config.data).toEqual(config);
      expect(result.config.loading).toBe(false);
      expect(result.config.error).toBeNull();
    });
  });

  // -----------------------------------------------------------------------
  // CONFIG_LOAD_ERROR
  // -----------------------------------------------------------------------
  describe("CONFIG_LOAD_ERROR", () => {
    it("sets error and loading=false", () => {
      const error = createMockError("Config load failed");
      const loadingState: AppState = {
        ...initialAppState,
        config: { data: null, loading: true, error: null },
      };
      const result = appReducer(loadingState, {
        type: "CONFIG_LOAD_ERROR",
        error,
      });
      expect(result.config.error).toBe(error);
      expect(result.config.loading).toBe(false);
    });

    it("preserves existing config data", () => {
      const config = createMockAddinConfiguration();
      const stateWithData: AppState = {
        ...initialAppState,
        config: { data: config, loading: true, error: null },
      };
      const error = createMockError("Reload failed");
      const result = appReducer(stateWithData, {
        type: "CONFIG_LOAD_ERROR",
        error,
      });
      expect(result.config.data).toEqual(config);
    });
  });

  // -----------------------------------------------------------------------
  // SELECT_API
  // -----------------------------------------------------------------------
  describe("SELECT_API", () => {
    it("sets selectedApi and populates prompt from promptTemplate", () => {
      const api = createMockApiCallConfig({
        promptTemplate: "Please {{prompt}}",
      });
      const result = appReducer(initialAppState, {
        type: "SELECT_API",
        api,
      });
      expect(result.selectedApi).toEqual(api);
      expect(result.prompt).toBe("Please {{prompt}}");
    });

    it("sets prompt to empty string when api has no promptTemplate", () => {
      const api = createMockApiCallConfig({ promptTemplate: undefined });
      const result = appReducer(initialAppState, {
        type: "SELECT_API",
        api,
      });
      expect(result.prompt).toBe("");
    });

    it("resets execution state", () => {
      const stateWithExecution: AppState = {
        ...initialAppState,
        execution: {
          data: createMockSuccessResponse(),
          loading: false,
          error: null,
        },
      };
      const api = createMockApiCallConfig();
      const result = appReducer(stateWithExecution, {
        type: "SELECT_API",
        api,
      });
      expect(result.execution).toEqual({
        data: null,
        loading: false,
        error: null,
      });
    });

    it("handles null api (deselection)", () => {
      const stateWithApi: AppState = {
        ...initialAppState,
        selectedApi: createMockApiCallConfig(),
        prompt: "existing prompt",
      };
      const result = appReducer(stateWithApi, {
        type: "SELECT_API",
        api: null,
      });
      expect(result.selectedApi).toBeNull();
      expect(result.prompt).toBe("");
    });
  });

  // -----------------------------------------------------------------------
  // SET_PROMPT
  // -----------------------------------------------------------------------
  describe("SET_PROMPT", () => {
    it("updates prompt", () => {
      const result = appReducer(initialAppState, {
        type: "SET_PROMPT",
        prompt: "New user prompt",
      });
      expect(result.prompt).toBe("New user prompt");
    });
  });

  // -----------------------------------------------------------------------
  // EXECUTION_START
  // -----------------------------------------------------------------------
  describe("EXECUTION_START", () => {
    it("sets loading=true and clears data and error", () => {
      const stateWithExecution: AppState = {
        ...initialAppState,
        execution: {
          data: createMockSuccessResponse(),
          loading: false,
          error: createMockError(),
        },
      };
      const result = appReducer(stateWithExecution, {
        type: "EXECUTION_START",
      });
      expect(result.execution.loading).toBe(true);
      expect(result.execution.data).toBeNull();
      expect(result.execution.error).toBeNull();
    });
  });

  // -----------------------------------------------------------------------
  // EXECUTION_SUCCESS
  // -----------------------------------------------------------------------
  describe("EXECUTION_SUCCESS", () => {
    it("sets data and loading=false", () => {
      const response = createMockSuccessResponse();
      const loadingState: AppState = {
        ...initialAppState,
        execution: { data: null, loading: true, error: null },
      };
      const result = appReducer(loadingState, {
        type: "EXECUTION_SUCCESS",
        response,
      });
      expect(result.execution.data).toEqual(response);
      expect(result.execution.loading).toBe(false);
      expect(result.execution.error).toBeNull();
    });

    it("works with error response variant", () => {
      const response = createMockErrorResponse();
      const result = appReducer(initialAppState, {
        type: "EXECUTION_SUCCESS",
        response,
      });
      expect(result.execution.data).toEqual(response);
    });
  });

  // -----------------------------------------------------------------------
  // EXECUTION_ERROR
  // -----------------------------------------------------------------------
  describe("EXECUTION_ERROR", () => {
    it("sets error and loading=false", () => {
      const error = createMockError("Execution failed");
      const loadingState: AppState = {
        ...initialAppState,
        execution: { data: null, loading: true, error: null },
      };
      const result = appReducer(loadingState, {
        type: "EXECUTION_ERROR",
        error,
      });
      expect(result.execution.error).toBe(error);
      expect(result.execution.loading).toBe(false);
    });

    it("preserves existing execution data", () => {
      const existingData = createMockSuccessResponse();
      const stateWithData: AppState = {
        ...initialAppState,
        execution: { data: existingData, loading: true, error: null },
      };
      const error = createMockError("Partial failure");
      const result = appReducer(stateWithData, {
        type: "EXECUTION_ERROR",
        error,
      });
      expect(result.execution.data).toEqual(existingData);
    });
  });

  // -----------------------------------------------------------------------
  // EXECUTION_RESET
  // -----------------------------------------------------------------------
  describe("EXECUTION_RESET", () => {
    it("resets execution to initial values", () => {
      const stateWithExecution: AppState = {
        ...initialAppState,
        execution: {
          data: createMockSuccessResponse(),
          loading: false,
          error: null,
        },
      };
      const result = appReducer(stateWithExecution, {
        type: "EXECUTION_RESET",
      });
      expect(result.execution).toEqual({
        data: null,
        loading: false,
        error: null,
      });
    });
  });

  // -----------------------------------------------------------------------
  // HISTORY_LOAD_START
  // -----------------------------------------------------------------------
  describe("HISTORY_LOAD_START", () => {
    it("sets loading=true and clears error", () => {
      const stateWithError: AppState = {
        ...initialAppState,
        history: { data: null, loading: false, error: createMockError() },
      };
      const result = appReducer(stateWithError, {
        type: "HISTORY_LOAD_START",
      });
      expect(result.history.loading).toBe(true);
      expect(result.history.error).toBeNull();
    });

    it("preserves existing history data", () => {
      const entries = [createMockHistoryEntry()];
      const stateWithData: AppState = {
        ...initialAppState,
        history: { data: entries, loading: false, error: null },
      };
      const result = appReducer(stateWithData, {
        type: "HISTORY_LOAD_START",
      });
      expect(result.history.data).toEqual(entries);
    });
  });

  // -----------------------------------------------------------------------
  // HISTORY_LOAD_SUCCESS
  // -----------------------------------------------------------------------
  describe("HISTORY_LOAD_SUCCESS", () => {
    it("sets data and loading=false", () => {
      const entries = [
        createMockHistoryEntry({ id: "e1" }),
        createMockHistoryEntry({ id: "e2" }),
      ];
      const loadingState: AppState = {
        ...initialAppState,
        history: { data: null, loading: true, error: null },
      };
      const result = appReducer(loadingState, {
        type: "HISTORY_LOAD_SUCCESS",
        entries,
      });
      expect(result.history.data).toEqual(entries);
      expect(result.history.loading).toBe(false);
      expect(result.history.error).toBeNull();
    });
  });

  // -----------------------------------------------------------------------
  // HISTORY_LOAD_ERROR
  // -----------------------------------------------------------------------
  describe("HISTORY_LOAD_ERROR", () => {
    it("sets error and loading=false", () => {
      const error = createMockError("History load failed");
      const loadingState: AppState = {
        ...initialAppState,
        history: { data: null, loading: true, error: null },
      };
      const result = appReducer(loadingState, {
        type: "HISTORY_LOAD_ERROR",
        error,
      });
      expect(result.history.error).toBe(error);
      expect(result.history.loading).toBe(false);
    });

    it("preserves existing history data", () => {
      const entries = [createMockHistoryEntry()];
      const stateWithData: AppState = {
        ...initialAppState,
        history: { data: entries, loading: true, error: null },
      };
      const error = createMockError("Reload failed");
      const result = appReducer(stateWithData, {
        type: "HISTORY_LOAD_ERROR",
        error,
      });
      expect(result.history.data).toEqual(entries);
    });
  });

  // -----------------------------------------------------------------------
  // HISTORY_ENTRY_ADDED
  // -----------------------------------------------------------------------
  describe("HISTORY_ENTRY_ADDED", () => {
    it("prepends entry to existing history data", () => {
      const existing = createMockHistoryEntry({ id: "old-1" });
      const newEntry = createMockHistoryEntry({ id: "new-1" });
      const stateWithHistory: AppState = {
        ...initialAppState,
        history: { data: [existing], loading: false, error: null },
      };
      const result = appReducer(stateWithHistory, {
        type: "HISTORY_ENTRY_ADDED",
        entry: newEntry,
      });
      expect(result.history.data).toHaveLength(2);
      expect(result.history.data![0].id).toBe("new-1");
      expect(result.history.data![1].id).toBe("old-1");
    });

    it("handles null data (creates array with single entry)", () => {
      const entry = createMockHistoryEntry({ id: "first-entry" });
      const result = appReducer(initialAppState, {
        type: "HISTORY_ENTRY_ADDED",
        entry,
      });
      expect(result.history.data).toHaveLength(1);
      expect(result.history.data![0].id).toBe("first-entry");
    });
  });

  // -----------------------------------------------------------------------
  // HISTORY_ENTRY_DELETED
  // -----------------------------------------------------------------------
  describe("HISTORY_ENTRY_DELETED", () => {
    it("filters out entry by id", () => {
      const entries = [
        createMockHistoryEntry({ id: "keep-1" }),
        createMockHistoryEntry({ id: "delete-me" }),
        createMockHistoryEntry({ id: "keep-2" }),
      ];
      const stateWithHistory: AppState = {
        ...initialAppState,
        history: { data: entries, loading: false, error: null },
      };
      const result = appReducer(stateWithHistory, {
        type: "HISTORY_ENTRY_DELETED",
        id: "delete-me",
      });
      expect(result.history.data).toHaveLength(2);
      expect(result.history.data!.map((e) => e.id)).toEqual([
        "keep-1",
        "keep-2",
      ]);
    });

    it("handles null data gracefully (returns empty array)", () => {
      const result = appReducer(initialAppState, {
        type: "HISTORY_ENTRY_DELETED",
        id: "non-existent",
      });
      expect(result.history.data).toEqual([]);
    });

    it("returns same entries if id not found", () => {
      const entries = [createMockHistoryEntry({ id: "e1" })];
      const stateWithHistory: AppState = {
        ...initialAppState,
        history: { data: entries, loading: false, error: null },
      };
      const result = appReducer(stateWithHistory, {
        type: "HISTORY_ENTRY_DELETED",
        id: "non-existent",
      });
      expect(result.history.data).toHaveLength(1);
    });
  });

  // -----------------------------------------------------------------------
  // HISTORY_CLEARED
  // -----------------------------------------------------------------------
  describe("HISTORY_CLEARED", () => {
    it("sets data to empty array", () => {
      const entries = [
        createMockHistoryEntry({ id: "e1" }),
        createMockHistoryEntry({ id: "e2" }),
      ];
      const stateWithHistory: AppState = {
        ...initialAppState,
        history: { data: entries, loading: false, error: null },
      };
      const result = appReducer(stateWithHistory, {
        type: "HISTORY_CLEARED",
      });
      expect(result.history.data).toEqual([]);
    });

    it("preserves loading and error state", () => {
      const error = createMockError();
      const stateWithError: AppState = {
        ...initialAppState,
        history: { data: [createMockHistoryEntry()], loading: false, error },
      };
      const result = appReducer(stateWithError, {
        type: "HISTORY_CLEARED",
      });
      expect(result.history.error).toBe(error);
      expect(result.history.loading).toBe(false);
    });
  });

  // -----------------------------------------------------------------------
  // SET_ACTIVE_TAB
  // -----------------------------------------------------------------------
  describe("SET_ACTIVE_TAB", () => {
    it("updates activeTab to history", () => {
      const result = appReducer(initialAppState, {
        type: "SET_ACTIVE_TAB",
        tab: "history",
      });
      expect(result.activeTab).toBe("history");
    });

    it("updates activeTab to apis", () => {
      const stateOnHistory: AppState = {
        ...initialAppState,
        activeTab: "history",
      };
      const result = appReducer(stateOnHistory, {
        type: "SET_ACTIVE_TAB",
        tab: "apis",
      });
      expect(result.activeTab).toBe("apis");
    });
  });

  // -----------------------------------------------------------------------
  // Unknown action type
  // -----------------------------------------------------------------------
  describe("unknown action", () => {
    it("returns current state for unknown action type", () => {
      const unknownAction = { type: "UNKNOWN_ACTION" } as unknown as AppAction;
      const result = appReducer(initialAppState, unknownAction);
      expect(result).toBe(initialAppState);
    });
  });

  // -----------------------------------------------------------------------
  // State immutability
  // -----------------------------------------------------------------------
  describe("state immutability", () => {
    it("does not mutate state on CONFIG_LOAD_SUCCESS", () => {
      const before = cloneState(initialAppState);
      appReducer(initialAppState, {
        type: "CONFIG_LOAD_SUCCESS",
        config: createMockAddinConfiguration(),
      });
      expect(initialAppState).toEqual(before);
    });

    it("does not mutate state on SELECT_API", () => {
      const state: AppState = {
        ...initialAppState,
        execution: {
          data: createMockSuccessResponse(),
          loading: false,
          error: null,
        },
      };
      const before = cloneState(state);
      appReducer(state, {
        type: "SELECT_API",
        api: createMockApiCallConfig(),
      });
      expect(state).toEqual(before);
    });

    it("does not mutate state on HISTORY_ENTRY_ADDED", () => {
      const entries = [createMockHistoryEntry({ id: "existing" })];
      const state: AppState = {
        ...initialAppState,
        history: { data: entries, loading: false, error: null },
      };
      const originalEntries = [...entries];
      appReducer(state, {
        type: "HISTORY_ENTRY_ADDED",
        entry: createMockHistoryEntry({ id: "new" }),
      });
      expect(state.history.data).toEqual(originalEntries);
    });

    it("does not mutate state on HISTORY_ENTRY_DELETED", () => {
      const entries = [
        createMockHistoryEntry({ id: "e1" }),
        createMockHistoryEntry({ id: "e2" }),
      ];
      const state: AppState = {
        ...initialAppState,
        history: { data: entries, loading: false, error: null },
      };
      const originalLength = entries.length;
      appReducer(state, { type: "HISTORY_ENTRY_DELETED", id: "e1" });
      expect(state.history.data).toHaveLength(originalLength);
    });

    it("returns a new state reference for every action", () => {
      const result = appReducer(initialAppState, {
        type: "SET_PROMPT",
        prompt: "test",
      });
      expect(result).not.toBe(initialAppState);
    });
  });
});
