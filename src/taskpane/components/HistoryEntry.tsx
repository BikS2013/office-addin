import React, { useState, useCallback } from "react";
import {
  Card,
  CardHeader,
  Button,
  Badge,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  PlayRegular,
  DeleteRegular,
  ChevronDownRegular,
  ChevronUpRegular,
} from "@fluentui/react-icons";
import type { HistoryEntry as HistoryEntryType } from "../types";

interface HistoryEntryProps {
  readonly entry: HistoryEntryType;
  readonly onReplay: (entry: HistoryEntryType) => void;
  readonly onDelete: (id: string) => void;
}

const useStyles = makeStyles({
  card: {
    marginBottom: tokens.spacingVerticalXS,
  },
  headerRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
    width: "100%",
  },
  headerText: {
    flexGrow: 1,
    overflow: "hidden",
  },
  details: {
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
  },
  detailRow: {
    display: "flex",
    gap: tokens.spacingHorizontalXS,
  },
  detailLabel: {
    color: tokens.colorNeutralForeground3,
    minWidth: "80px",
  },
  detailValue: {
    wordBreak: "break-word",
  },
  actions: {
    display: "flex",
    gap: tokens.spacingHorizontalXS,
    marginTop: tokens.spacingVerticalXS,
    padding: `0 ${tokens.spacingHorizontalS} ${tokens.spacingVerticalXS}`,
  },
  timestamp: {
    color: tokens.colorNeutralForeground3,
  },
});

function formatRelativeTime(timestamp: number): string {
  const now = Date.now();
  const diffMs = now - timestamp;
  const diffSec = Math.floor(diffMs / 1000);
  const diffMin = Math.floor(diffSec / 60);
  const diffHour = Math.floor(diffMin / 60);
  const diffDay = Math.floor(diffHour / 24);

  if (diffSec < 60) return "just now";
  if (diffMin < 60) return `${diffMin}m ago`;
  if (diffHour < 24) return `${diffHour}h ago`;
  if (diffDay < 7) return `${diffDay}d ago`;
  return new Date(timestamp).toLocaleDateString();
}

export const HistoryEntryComponent: React.FC<HistoryEntryProps> = ({
  entry,
  onReplay,
  onDelete,
}) => {
  const styles = useStyles();
  const [isExpanded, setIsExpanded] = useState(false);

  const handleToggle = useCallback(() => {
    setIsExpanded((prev) => !prev);
  }, []);

  const handleReplay = useCallback(
    (e: React.MouseEvent) => {
      e.stopPropagation();
      onReplay(entry);
    },
    [entry, onReplay]
  );

  const handleDelete = useCallback(
    (e: React.MouseEvent) => {
      e.stopPropagation();
      onDelete(entry.id);
    },
    [entry.id, onDelete]
  );

  return (
    <Card className={styles.card} size="small">
      <CardHeader
        onClick={handleToggle}
        style={{ cursor: "pointer" }}
        header={
          <div className={styles.headerRow}>
            <div className={styles.headerText}>
              <Text weight="semibold" size={200} block>
                {entry.apiName}
              </Text>
              <Text className={styles.timestamp} size={100} block>
                {formatRelativeTime(entry.timestamp)} - {entry.documentName}
              </Text>
              {entry.prompt && (
                <Text size={100} block truncate>
                  {entry.prompt.length > 50
                    ? `${entry.prompt.substring(0, 50)}...`
                    : entry.prompt}
                </Text>
              )}
            </div>
            <Badge
              appearance="filled"
              color={entry.wasSuccessful ? "success" : "danger"}
              size="small"
            >
              {entry.wasSuccessful ? "OK" : "Fail"}
            </Badge>
            {isExpanded ? (
              <ChevronUpRegular fontSize={12} />
            ) : (
              <ChevronDownRegular fontSize={12} />
            )}
          </div>
        }
      />

      {isExpanded && (
        <>
          <div className={styles.details}>
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel} size={100} weight="semibold">
                API URL:
              </Text>
              <Text className={styles.detailValue} size={100}>
                {entry.apiUrl}
              </Text>
            </div>
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel} size={100} weight="semibold">
                Source:
              </Text>
              <Text className={styles.detailValue} size={100}>
                {entry.textSource === "selected" ? "Selected Text" : "Full Document"}
              </Text>
            </div>
            {entry.prompt && (
              <div className={styles.detailRow}>
                <Text className={styles.detailLabel} size={100} weight="semibold">
                  Prompt:
                </Text>
                <Text className={styles.detailValue} size={100}>
                  {entry.prompt}
                </Text>
              </div>
            )}
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel} size={100} weight="semibold">
                Text Preview:
              </Text>
              <Text className={styles.detailValue} size={100}>
                {entry.textPreview}
              </Text>
            </div>
            {entry.responsePreview && (
              <div className={styles.detailRow}>
                <Text className={styles.detailLabel} size={100} weight="semibold">
                  Response:
                </Text>
                <Text className={styles.detailValue} size={100}>
                  {entry.responsePreview}
                </Text>
              </div>
            )}
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel} size={100} weight="semibold">
                Duration:
              </Text>
              <Text className={styles.detailValue} size={100}>
                {entry.durationMs}ms
              </Text>
            </div>
          </div>
          <div className={styles.actions}>
            <Button
              size="small"
              icon={<PlayRegular />}
              onClick={handleReplay}
            >
              Replay
            </Button>
            <Button
              size="small"
              icon={<DeleteRegular />}
              onClick={handleDelete}
              appearance="subtle"
            >
              Delete
            </Button>
          </div>
        </>
      )}
    </Card>
  );
};
