import { AddinConfiguration, ApiCallConfig } from "./config";
import { ApiResponse } from "./api";
import { HistoryEntry } from "./history";
import { AppBaseError } from "./errors";

export type ActiveTab = "apis" | "history";

export interface AsyncState<T> {
  readonly data: T | null;
  readonly loading: boolean;
  readonly error: AppBaseError | null;
}

export interface AppState {
  readonly configUrl: string;
  readonly config: AsyncState<AddinConfiguration>;
  readonly selectedApi: ApiCallConfig | null;
  readonly prompt: string;
  readonly execution: AsyncState<ApiResponse>;
  readonly history: AsyncState<readonly HistoryEntry[]>;
  readonly activeTab: ActiveTab;
}

export type AppAction =
  | { readonly type: "SET_CONFIG_URL"; readonly url: string }
  | { readonly type: "CONFIG_LOAD_START" }
  | { readonly type: "CONFIG_LOAD_SUCCESS"; readonly config: AddinConfiguration }
  | { readonly type: "CONFIG_LOAD_ERROR"; readonly error: AppBaseError }
  | { readonly type: "SELECT_API"; readonly api: ApiCallConfig | null }
  | { readonly type: "SET_PROMPT"; readonly prompt: string }
  | { readonly type: "EXECUTION_START" }
  | { readonly type: "EXECUTION_SUCCESS"; readonly response: ApiResponse }
  | { readonly type: "EXECUTION_ERROR"; readonly error: AppBaseError }
  | { readonly type: "EXECUTION_RESET" }
  | { readonly type: "HISTORY_LOAD_START" }
  | { readonly type: "HISTORY_LOAD_SUCCESS"; readonly entries: readonly HistoryEntry[] }
  | { readonly type: "HISTORY_LOAD_ERROR"; readonly error: AppBaseError }
  | { readonly type: "HISTORY_ENTRY_ADDED"; readonly entry: HistoryEntry }
  | { readonly type: "HISTORY_ENTRY_DELETED"; readonly id: string }
  | { readonly type: "HISTORY_CLEARED" }
  | { readonly type: "SET_ACTIVE_TAB"; readonly tab: ActiveTab };
