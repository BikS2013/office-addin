import React, { useCallback } from "react";
import {
  Card,
  CardHeader,
  Button,
  Spinner,
  MessageBar,
  MessageBarBody,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  CopyRegular,
  InsertRegular,
} from "@fluentui/react-icons";
import { useAppState } from "../context/AppContext";
import type { ApiSuccessResponse } from "../types";
import { OfficeApiService } from "../services/OfficeApiService";

const officeApiService = new OfficeApiService();

const useStyles = makeStyles({
  container: {
    marginTop: tokens.spacingVerticalS,
  },
  responseContent: {
    maxHeight: "200px",
    overflowY: "auto",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    backgroundColor: tokens.colorNeutralBackground3,
    padding: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusMedium,
  },
  actions: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalXS,
  },
});

export const ResponsePanel: React.FC = () => {
  const styles = useStyles();
  const { state } = useAppState();
  const { execution } = state;

  const handleCopyToClipboard = useCallback(async () => {
    const response = execution.data;
    if (response && response.kind === "success") {
      try {
        await navigator.clipboard.writeText(response.extractedText);
      } catch (err) {
        console.error("Failed to copy to clipboard:", err);
      }
    }
  }, [execution.data]);

  const handleInsertAtCursor = useCallback(async () => {
    const response = execution.data;
    if (response && response.kind === "success") {
      try {
        await officeApiService.insertAtCursor(response.extractedText);
      } catch (err) {
        console.error("Failed to insert at cursor:", err);
      }
    }
  }, [execution.data]);

  if (!execution.loading && !execution.data && !execution.error) {
    return null;
  }

  return (
    <Card className={styles.container} size="small">
      <CardHeader header={<Text weight="semibold" size={200}>Response</Text>} />

      {execution.loading && (
        <Spinner size="small" label="Executing API call..." />
      )}

      {execution.error && (
        <MessageBar intent="error">
          <MessageBarBody>{execution.error.userMessage}</MessageBarBody>
        </MessageBar>
      )}

      {execution.data && execution.data.kind === "success" && (
        <>
          <div className={styles.responseContent}>
            {(execution.data as ApiSuccessResponse).extractedText}
          </div>
          <div className={styles.actions}>
            <Button
              size="small"
              icon={<CopyRegular />}
              onClick={handleCopyToClipboard}
            >
              Copy to Clipboard
            </Button>
            <Button
              size="small"
              icon={<InsertRegular />}
              onClick={handleInsertAtCursor}
            >
              Insert at Cursor
            </Button>
          </div>
        </>
      )}

      {execution.data && execution.data.kind === "error" && (
        <MessageBar intent="error">
          <MessageBarBody>{execution.data.message}</MessageBarBody>
        </MessageBar>
      )}
    </Card>
  );
};
