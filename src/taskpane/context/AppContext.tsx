import React, { createContext, useContext, useReducer, Dispatch } from "react";
import { AppState, AppAction } from "../types/state";
import { appReducer, initialAppState } from "./appReducer";

interface AppContextValue {
  readonly state: AppState;
  readonly dispatch: Dispatch<AppAction>;
}

const AppContext = createContext<AppContextValue | null>(null);

export function AppProvider({
  children,
}: {
  readonly children: React.ReactNode;
}): React.ReactElement {
  const [state, dispatch] = useReducer(appReducer, initialAppState);

  return (
    <AppContext.Provider value={{ state, dispatch }}>
      {children}
    </AppContext.Provider>
  );
}

export function useAppState(): AppContextValue {
  const context = useContext(AppContext);
  if (!context) {
    throw new Error("useAppState must be used within an AppProvider");
  }
  return context;
}
