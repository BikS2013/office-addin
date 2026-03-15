import { useCallback, useEffect, useRef } from "react";
import { useAppState } from "../context/AppContext";
import { historyServiceInstance as historyService } from "../services/historyServiceInstance";
import { AppBaseError } from "../types/errors";
import type { HistoryEntry, HistoryFilter } from "../types/history";

export function useHistory() {
  const { state, dispatch } = useAppState();
  const initializedRef = useRef(false);

  useEffect(() => {
    if (initializedRef.current) return;
    initializedRef.current = true;

    historyService.initialize().catch((error) => {
      dispatch({
        type: "HISTORY_LOAD_ERROR",
        error: error as AppBaseError,
      });
    });
  }, [dispatch]);

  const loadHistory = useCallback(async (filter?: HistoryFilter) => {
    dispatch({ type: "HISTORY_LOAD_START" });
    try {
      const entries = await historyService.getEntries(filter);
      dispatch({ type: "HISTORY_LOAD_SUCCESS", entries });
    } catch (error) {
      dispatch({ type: "HISTORY_LOAD_ERROR", error: error as AppBaseError });
    }
  }, [dispatch]);

  const addEntry = useCallback(async (entry: HistoryEntry) => {
    try {
      await historyService.addEntry(entry);
      dispatch({ type: "HISTORY_ENTRY_ADDED", entry });
    } catch (error) {
      dispatch({ type: "HISTORY_LOAD_ERROR", error: error as AppBaseError });
    }
  }, [dispatch]);

  const deleteEntry = useCallback(async (id: string) => {
    try {
      await historyService.deleteEntry(id);
      dispatch({ type: "HISTORY_ENTRY_DELETED", id });
    } catch (error) {
      dispatch({ type: "HISTORY_LOAD_ERROR", error: error as AppBaseError });
    }
  }, [dispatch]);

  const clearAll = useCallback(async () => {
    try {
      await historyService.clearAll();
      dispatch({ type: "HISTORY_CLEARED" });
    } catch (error) {
      dispatch({ type: "HISTORY_LOAD_ERROR", error: error as AppBaseError });
    }
  }, [dispatch]);

  return {
    history: state.history,
    loadHistory,
    addEntry,
    deleteEntry,
    clearAll,
  } as const;
}
