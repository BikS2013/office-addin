import React, { useEffect, useState } from "react";
import {
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  Button,
  Spinner,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { useAppState } from "../context/AppContext";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    padding: `${tokens.spacingVerticalXS} 0`,
  },
});

export const StatusBar: React.FC = () => {
  const styles = useStyles();
  const { state } = useAppState();
  const [showSuccess, setShowSuccess] = useState(false);

  const configLoading = state.config.loading;
  const executionLoading = state.execution.loading;
  const configError = state.config.error;
  const executionError = state.execution.error;
  const configLoaded = state.config.data !== null;

  useEffect(() => {
    if (configLoaded && !configLoading && !configError) {
      setShowSuccess(true);
      const timer = setTimeout(() => setShowSuccess(false), 3000);
      return () => clearTimeout(timer);
    }
    return undefined;
  }, [configLoaded, configLoading, configError]);

  const isLoading = configLoading || executionLoading;
  const error = configError ?? executionError;

  if (!isLoading && !error && !showSuccess) {
    return null;
  }

  return (
    <div className={styles.container}>
      {isLoading && (
        <Spinner
          size="tiny"
          label={configLoading ? "Loading configuration..." : "Executing API call..."}
        />
      )}

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error.userMessage}</MessageBarBody>
          {error.retryable && (
            <MessageBarActions>
              <Button size="small" onClick={() => window.location.reload()}>
                Retry
              </Button>
            </MessageBarActions>
          )}
        </MessageBar>
      )}

      {showSuccess && !isLoading && !error && (
        <MessageBar intent="success">
          <MessageBarBody>Configuration loaded successfully.</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
};
