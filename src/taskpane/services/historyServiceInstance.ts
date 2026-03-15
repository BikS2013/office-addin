import { HistoryService } from "./HistoryService";

/**
 * Shared singleton instance of HistoryService.
 * Must be initialized via initialize() before use (handled by useHistory hook).
 */
export const historyServiceInstance = new HistoryService();
