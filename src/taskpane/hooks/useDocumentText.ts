import { useCallback } from "react";
import { OfficeApiService } from "../services/OfficeApiService";

const officeApiService = new OfficeApiService();

export function useDocumentText() {
  const getSelectedText = useCallback(async (): Promise<string> => {
    return officeApiService.getSelectedText();
  }, []);

  const getFullDocumentText = useCallback(async (): Promise<string> => {
    return officeApiService.getFullDocumentText();
  }, []);

  const getDocumentName = useCallback(async (): Promise<string> => {
    return officeApiService.getDocumentName();
  }, []);

  return {
    getSelectedText,
    getFullDocumentText,
    getDocumentName,
  } as const;
}
