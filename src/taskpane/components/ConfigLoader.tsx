import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  Input,
  Button,
  makeStyles,
  tokens,
  Text,
  Divider,
} from "@fluentui/react-components";
import {
  ArrowSyncRegular,
  FolderOpenRegular,
} from "@fluentui/react-icons";
import { useAppState } from "../context/AppContext";
import { useConfig } from "../hooks/useConfig";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalS,
  },
  row: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    alignItems: "center",
  },
  input: {
    flexGrow: 1,
  },
  dividerRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  dividerLabel: {
    whiteSpace: "nowrap",
    color: tokens.colorNeutralForeground3,
  },
  fileInput: {
    display: "none",
  },
  sourceLabel: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
});

export const ConfigLoader: React.FC = () => {
  const styles = useStyles();
  const { state } = useAppState();
  const { loadConfig, loadConfigFromFile, reloadConfig, getSavedUrl } = useConfig();
  const [inputUrl, setInputUrl] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const savedUrl = getSavedUrl();
    if (savedUrl) {
      setInputUrl(savedUrl);
    }
  }, [getSavedUrl]);

  const handleLoad = useCallback(() => {
    const trimmedUrl = inputUrl.trim();
    if (trimmedUrl) {
      loadConfig(trimmedUrl);
    }
  }, [inputUrl, loadConfig]);

  const handleReload = useCallback(() => {
    reloadConfig();
  }, [reloadConfig]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === "Enter") {
        handleLoad();
      }
    },
    [handleLoad]
  );

  const handleFileButtonClick = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const handleFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) {
        loadConfigFromFile(file);
      }
      // Reset so the same file can be re-selected
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    },
    [loadConfigFromFile]
  );

  const isLoading = state.config.loading;
  const configLoaded = state.config.data !== null;
  const isFileSource = state.configUrl.startsWith("file://");

  return (
    <div className={styles.container}>
      <div className={styles.row}>
        <Input
          className={styles.input}
          placeholder="Enter configuration URL..."
          value={inputUrl}
          onChange={(_e, data) => setInputUrl(data.value)}
          onKeyDown={handleKeyDown}
          disabled={isLoading}
          size="small"
        />
        <Button
          appearance="primary"
          size="small"
          onClick={handleLoad}
          disabled={isLoading || !inputUrl.trim()}
        >
          Load
        </Button>
        {configLoaded && !isFileSource && (
          <Button
            appearance="subtle"
            size="small"
            icon={<ArrowSyncRegular />}
            onClick={handleReload}
            disabled={isLoading}
            title="Reload configuration"
          />
        )}
      </div>

      <div className={styles.dividerRow}>
        <Divider style={{ flexGrow: 1 }} />
        <Text className={styles.dividerLabel} size={200}>or</Text>
        <Divider style={{ flexGrow: 1 }} />
      </div>

      <Button
        appearance="secondary"
        size="small"
        icon={<FolderOpenRegular />}
        onClick={handleFileButtonClick}
        disabled={isLoading}
      >
        Load from File
      </Button>
      <input
        ref={fileInputRef}
        type="file"
        accept=".json"
        className={styles.fileInput}
        onChange={handleFileChange}
      />

      {configLoaded && (
        <Text className={styles.sourceLabel}>
          Loaded: {isFileSource ? state.configUrl.replace("file://", "") : state.configUrl}
        </Text>
      )}
    </div>
  );
};
