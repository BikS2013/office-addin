import React from "react";
import {
  FluentProvider,
  webLightTheme,
  TabList,
  Tab,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { SelectTabData, SelectTabEvent } from "@fluentui/react-components";
import { ErrorBoundary } from "./components/ErrorBoundary";
import { ConfigLoader } from "./components/ConfigLoader";
import { StatusBar } from "./components/StatusBar";
import { ApiTree } from "./components/ApiTree";
import { HistoryPanel } from "./components/HistoryPanel";
import { useAppState } from "./context/AppContext";
import type { ActiveTab } from "./types";
import "./taskpane.css";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    width: "100%",
    overflow: "hidden",
  },
  tabContent: {
    flexGrow: 1,
    overflowY: "auto",
    padding: `0 ${tokens.spacingHorizontalS}`,
  },
  tabList: {
    padding: `0 ${tokens.spacingHorizontalS}`,
  },
});

const AppContent: React.FC = () => {
  const styles = useStyles();
  const { state, dispatch } = useAppState();

  const handleTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
    dispatch({ type: "SET_ACTIVE_TAB", tab: data.value as ActiveTab });
  };

  return (
    <div className={styles.root}>
      <ConfigLoader />
      <StatusBar />

      <TabList
        className={styles.tabList}
        selectedValue={state.activeTab}
        onTabSelect={handleTabSelect}
        size="small"
      >
        <Tab value="apis">APIs</Tab>
        <Tab value="history">History</Tab>
      </TabList>

      <div className={styles.tabContent}>
        {state.activeTab === "apis" && state.config.data && (
          <ApiTree groups={state.config.data.groups} />
        )}
        {state.activeTab === "history" && <HistoryPanel />}
      </div>
    </div>
  );
};

export const App: React.FC = () => {
  return (
    <FluentProvider theme={webLightTheme}>
      <ErrorBoundary>
        <AppContent />
      </ErrorBoundary>
    </FluentProvider>
  );
};
