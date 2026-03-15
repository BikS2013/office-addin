import React, { useCallback } from "react";
import { Button, makeStyles, tokens } from "@fluentui/react-components";
import {
  SelectAllOnRegular,
  DocumentRegular,
} from "@fluentui/react-icons";
import type { ApiCallConfig } from "../types";
import { useAppState } from "../context/AppContext";

interface ActionButtonsProps {
  readonly api: ApiCallConfig;
  readonly onExecute: (textSource: "selected" | "full") => void;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalXS,
    flexWrap: "wrap",
  },
});

export const ActionButtons: React.FC<ActionButtonsProps> = ({
  api,
  onExecute,
}) => {
  const styles = useStyles();
  const { state } = useAppState();
  const isExecuting = state.execution.loading;

  const handleSelectedClick = useCallback(() => {
    onExecute("selected");
  }, [onExecute]);

  const handleFullClick = useCallback(() => {
    onExecute("full");
  }, [onExecute]);

  const showSelectedButton =
    api.inputMode === "selected" || api.inputMode === "both";
  const showFullButton =
    api.inputMode === "full" || api.inputMode === "both";

  return (
    <div className={styles.container}>
      {showSelectedButton && (
        <Button
          appearance="primary"
          size="small"
          icon={<SelectAllOnRegular />}
          onClick={handleSelectedClick}
          disabled={isExecuting}
        >
          Send Selected Text
        </Button>
      )}
      {showFullButton && (
        <Button
          appearance="primary"
          size="small"
          icon={<DocumentRegular />}
          onClick={handleFullClick}
          disabled={isExecuting}
        >
          Send Full Document
        </Button>
      )}
    </div>
  );
};
