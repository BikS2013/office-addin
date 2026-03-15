import { useCallback } from "react";
import { useAppState } from "../context/AppContext";
import { ApiExecutionService } from "../services/ApiExecutionService";
import { OfficeApiService } from "../services/OfficeApiService";
import { historyServiceInstance as historyService } from "../services/historyServiceInstance";
import { AppBaseError } from "../types/errors";
import type { ApiCallConfig, ApiRequestPayload, HistoryEntry } from "../types";

const apiExecutionService = new ApiExecutionService();
const officeApiService = new OfficeApiService();

function generateId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(36).substring(2, 11)}`;
}

export function useApiExecution() {
  const { state, dispatch } = useAppState();

  const executeApi = useCallback(
    async (
      apiConfig: ApiCallConfig,
      prompt: string,
      textSource: "selected" | "full"
    ) => {
      dispatch({ type: "EXECUTION_START" });

      const startTime = performance.now();

      try {
        const text =
          textSource === "selected"
            ? await officeApiService.getSelectedText()
            : await officeApiService.getFullDocumentText();

        const documentName = await officeApiService.getDocumentName();

        const payload: ApiRequestPayload = {
          prompt,
          text,
          documentName,
          textSource,
        };

        const response = await apiExecutionService.execute(apiConfig, payload);

        if (response.kind === "success") {
          dispatch({ type: "EXECUTION_SUCCESS", response });
        } else {
          dispatch({
            type: "EXECUTION_ERROR",
            error: new AppBaseError(
              response.message,
              response.errorType === "network" || response.errorType === "timeout",
              response.message
            ),
          });
        }

        const wasSuccessful = response.kind === "success";
        const durationMs = Math.round(performance.now() - startTime);

        const textPreview = text.length > 200 ? text.substring(0, 200) + "..." : text;
        const responsePreview =
          wasSuccessful && response.kind === "success"
            ? response.extractedText.length > 200
              ? response.extractedText.substring(0, 200) + "..."
              : response.extractedText
            : response.kind === "error"
              ? response.message
              : undefined;

        const historyEntry: HistoryEntry = {
          id: generateId(),
          timestamp: Date.now(),
          apiId: apiConfig.id,
          apiName: apiConfig.name,
          apiUrl: apiConfig.url,
          prompt,
          textSource,
          textPreview,
          documentName,
          wasSuccessful,
          responsePreview,
          durationMs,
        };

        try {
          await historyService.addEntry(historyEntry);
        } catch {
          // History persistence failure is non-critical; do not disrupt execution flow
        }

        dispatch({ type: "HISTORY_ENTRY_ADDED", entry: historyEntry });
      } catch (error) {
        dispatch({
          type: "EXECUTION_ERROR",
          error: error as AppBaseError,
        });
      }
    },
    [dispatch]
  );

  const resetExecution = useCallback(() => {
    dispatch({ type: "EXECUTION_RESET" });
  }, [dispatch]);

  return {
    execution: state.execution,
    executeApi,
    resetExecution,
  } as const;
}
