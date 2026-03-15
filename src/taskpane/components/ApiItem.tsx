import React, { useCallback } from "react";
import {
  Card,
  CardHeader,
  Text,
  makeStyles,
  tokens,
  mergeClasses,
} from "@fluentui/react-components";
import type { ApiCallConfig } from "../types";
import { useAppState } from "../context/AppContext";
import { useApiExecution } from "../hooks/useApiExecution";
import { PromptInput } from "./PromptInput";
import { ActionButtons } from "./ActionButtons";
import { ResponsePanel } from "./ResponsePanel";

interface ApiItemProps {
  readonly api: ApiCallConfig;
}

const useStyles = makeStyles({
  card: {
    marginBottom: tokens.spacingVerticalS,
    cursor: "pointer",
  },
  cardSelected: {
    border: `2px solid ${tokens.colorBrandStroke1}`,
  },
  description: {
    color: tokens.colorNeutralForeground3,
    marginTop: tokens.spacingVerticalXXS,
  },
  expandedContent: {
    padding: `0 ${tokens.spacingHorizontalS} ${tokens.spacingVerticalS}`,
  },
});

export const ApiItem: React.FC<ApiItemProps> = ({ api }) => {
  const styles = useStyles();
  const { state, dispatch } = useAppState();
  const { executeApi } = useApiExecution();

  const isSelected = state.selectedApi?.id === api.id;

  const handleSelect = useCallback(() => {
    dispatch({ type: "SELECT_API", api: isSelected ? null : api });
  }, [dispatch, api, isSelected]);

  const handleExecute = useCallback(
    (textSource: "selected" | "full") => {
      executeApi(api, state.prompt, textSource);
    },
    [executeApi, api, state.prompt]
  );

  return (
    <Card
      className={mergeClasses(
        styles.card,
        isSelected ? styles.cardSelected : undefined
      )}
      size="small"
      onClick={handleSelect}
    >
      <CardHeader
        header={<Text weight="semibold" size={300}>{api.name}</Text>}
        description={
          api.description ? (
            <Text className={styles.description} size={200}>
              {api.description}
            </Text>
          ) : undefined
        }
      />

      {isSelected && (
        <div className={styles.expandedContent} onClick={(e) => e.stopPropagation()}>
          <PromptInput />
          <ActionButtons api={api} onExecute={handleExecute} />
          <ResponsePanel />
        </div>
      )}
    </Card>
  );
};
