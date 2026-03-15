import React, { useEffect, useState, useCallback, useMemo } from "react";
import {
  Button,
  Text,
  Spinner,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Dropdown,
  Option,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { DeleteRegular } from "@fluentui/react-icons";
import { useAppState } from "../context/AppContext";
import { useHistory } from "../hooks/useHistory";
import { HistoryEntryComponent } from "./HistoryEntry";
import type { HistoryEntry } from "../types";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalS,
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  filter: {
    minWidth: "150px",
  },
  list: {
    display: "flex",
    flexDirection: "column",
    overflowY: "auto",
    maxHeight: "calc(100vh - 250px)",
  },
  emptyState: {
    textAlign: "center",
    padding: tokens.spacingVerticalXXL,
    color: tokens.colorNeutralForeground3,
  },
});

export const HistoryPanel: React.FC = () => {
  const styles = useStyles();
  const { state, dispatch } = useAppState();
  const { loadHistory, deleteEntry, clearAll } = useHistory();
  const [filterDocName, setFilterDocName] = useState<string | undefined>(undefined);
  const [clearDialogOpen, setClearDialogOpen] = useState(false);

  useEffect(() => {
    loadHistory();
  }, [loadHistory]);

  const entries = state.history.data ?? [];

  const documentNames = useMemo(() => {
    const names = new Set(entries.map((e) => e.documentName));
    return Array.from(names).sort();
  }, [entries]);

  const filteredEntries = useMemo(() => {
    if (!filterDocName) return entries;
    return entries.filter((e) => e.documentName === filterDocName);
  }, [entries, filterDocName]);

  const handleReplay = useCallback(
    (entry: HistoryEntry) => {
      // Try to find the matching API config in the loaded configuration
      const findApiInGroups = (
        groups: readonly import("../types").ApiGroup[]
      ): import("../types").ApiCallConfig | undefined => {
        for (const group of groups) {
          if (group.apis) {
            const found = group.apis.find((a) => a.id === entry.apiId);
            if (found) return found;
          }
          if (group.groups) {
            const found = findApiInGroups(group.groups);
            if (found) return found;
          }
        }
        return undefined;
      };

      const loadedConfig = state.config.data;
      const matchedApi = loadedConfig
        ? findApiInGroups(loadedConfig.groups)
        : undefined;

      // Use the matched config if found, otherwise construct a minimal replay reference
      const replayApi = matchedApi ?? {
        id: entry.apiId,
        name: entry.apiName,
        url: entry.apiUrl,
        method: "POST" as const,
        inputMode: entry.textSource === "full" ? ("full" as const) : ("selected" as const),
        timeout: 30000,
      };

      dispatch({ type: "SELECT_API", api: replayApi });
      dispatch({ type: "SET_PROMPT", prompt: entry.prompt });
      dispatch({ type: "SET_ACTIVE_TAB", tab: "apis" });
    },
    [dispatch, state.config.data]
  );

  const handleDelete = useCallback(
    (id: string) => {
      deleteEntry(id);
    },
    [deleteEntry]
  );

  const handleClear = useCallback(() => {
    clearAll();
    setClearDialogOpen(false);
  }, [clearAll]);

  const handleFilterChange = useCallback(
    (_e: unknown, data: { optionValue?: string }) => {
      setFilterDocName(data.optionValue === "" ? undefined : data.optionValue);
    },
    []
  );

  if (state.history.loading) {
    return (
      <div className={styles.container}>
        <Spinner size="small" label="Loading history..." />
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Text weight="semibold" size={400}>
          History
        </Text>
        <Dialog open={clearDialogOpen} onOpenChange={(_e, data) => setClearDialogOpen(data.open)}>
          <DialogTrigger disableButtonEnhancement>
            <Button
              size="small"
              icon={<DeleteRegular />}
              appearance="subtle"
              disabled={entries.length === 0}
            >
              Clear History
            </Button>
          </DialogTrigger>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Clear History</DialogTitle>
              <DialogContent>
                Are you sure you want to clear all history entries? This action cannot be undone.
              </DialogContent>
              <DialogActions>
                <DialogTrigger disableButtonEnhancement>
                  <Button appearance="secondary">Cancel</Button>
                </DialogTrigger>
                <Button appearance="primary" onClick={handleClear}>
                  Clear All
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>

      {documentNames.length > 1 && (
        <Dropdown
          className={styles.filter}
          placeholder="Filter by document..."
          size="small"
          onOptionSelect={handleFilterChange}
          value={filterDocName ?? ""}
        >
          <Option value="">All Documents</Option>
          {documentNames.map((name) => (
            <Option key={name} value={name}>
              {name}
            </Option>
          ))}
        </Dropdown>
      )}

      <div className={styles.list}>
        {filteredEntries.length === 0 ? (
          <div className={styles.emptyState}>
            <Text size={200}>No history entries yet.</Text>
          </div>
        ) : (
          filteredEntries.map((entry) => (
            <HistoryEntryComponent
              key={entry.id}
              entry={entry}
              onReplay={handleReplay}
              onDelete={handleDelete}
            />
          ))
        )}
      </div>
    </div>
  );
};
