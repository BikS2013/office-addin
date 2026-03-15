// Configuration types
export type { InputMode, ApiCallConfig, ApiGroup, AddinConfiguration, CachedConfiguration } from "./config";

// API types
export type {
  ApiRequestPayload,
  ConstructedRequest,
  ApiSuccessResponse,
  ApiErrorResponse,
  ApiResponse,
  DocumentTextInfo,
} from "./api";

// History types
export type { HistoryEntry, HistoryFilter } from "./history";

// State types
export type { ActiveTab, AsyncState, AppState, AppAction } from "./state";

// Error classes
export {
  AppBaseError,
  ConfigurationError,
  ConfigFetchError,
  ApiExecutionError,
  OfficeApiError,
  HistoryStorageError,
} from "./errors";
