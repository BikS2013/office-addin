import { AppState, AppAction } from "../types/state";

export const initialAppState: AppState = {
  configUrl: "",
  config: { data: null, loading: false, error: null },
  selectedApi: null,
  prompt: "",
  execution: { data: null, loading: false, error: null },
  history: { data: null, loading: false, error: null },
  activeTab: "apis",
};

export function appReducer(state: AppState, action: AppAction): AppState {
  switch (action.type) {
    case "SET_CONFIG_URL":
      return { ...state, configUrl: action.url };

    case "CONFIG_LOAD_START":
      return {
        ...state,
        config: { ...state.config, loading: true, error: null },
      };

    case "CONFIG_LOAD_SUCCESS":
      return {
        ...state,
        config: { data: action.config, loading: false, error: null },
      };

    case "CONFIG_LOAD_ERROR":
      return {
        ...state,
        config: { ...state.config, loading: false, error: action.error },
      };

    case "SELECT_API":
      return {
        ...state,
        selectedApi: action.api,
        prompt: action.api?.promptTemplate ?? "",
        execution: { data: null, loading: false, error: null },
      };

    case "SET_PROMPT":
      return { ...state, prompt: action.prompt };

    case "EXECUTION_START":
      return {
        ...state,
        execution: { data: null, loading: true, error: null },
      };

    case "EXECUTION_SUCCESS":
      return {
        ...state,
        execution: { data: action.response, loading: false, error: null },
      };

    case "EXECUTION_ERROR":
      return {
        ...state,
        execution: { ...state.execution, loading: false, error: action.error },
      };

    case "EXECUTION_RESET":
      return {
        ...state,
        execution: { data: null, loading: false, error: null },
      };

    case "HISTORY_LOAD_START":
      return {
        ...state,
        history: { ...state.history, loading: true, error: null },
      };

    case "HISTORY_LOAD_SUCCESS":
      return {
        ...state,
        history: { data: action.entries, loading: false, error: null },
      };

    case "HISTORY_LOAD_ERROR":
      return {
        ...state,
        history: { ...state.history, loading: false, error: action.error },
      };

    case "HISTORY_ENTRY_ADDED":
      return {
        ...state,
        history: {
          ...state.history,
          data: [action.entry, ...(state.history.data ?? [])],
        },
      };

    case "HISTORY_ENTRY_DELETED":
      return {
        ...state,
        history: {
          ...state.history,
          data: (state.history.data ?? []).filter(
            (entry) => entry.id !== action.id,
          ),
        },
      };

    case "HISTORY_CLEARED":
      return {
        ...state,
        history: { ...state.history, data: [] },
      };

    case "SET_ACTIVE_TAB":
      return { ...state, activeTab: action.tab };

    default:
      return state;
  }
}
