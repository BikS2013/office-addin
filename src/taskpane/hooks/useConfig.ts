import { useCallback } from "react";
import { useAppState } from "../context/AppContext";
import { ConfigService } from "../services/ConfigService";
import { AppBaseError } from "../types/errors";

const configService = new ConfigService();

export function useConfig() {
  const { state, dispatch } = useAppState();

  const loadConfig = useCallback(async (url: string) => {
    dispatch({ type: "SET_CONFIG_URL", url });
    dispatch({ type: "CONFIG_LOAD_START" });
    try {
      const config = await configService.loadConfig(url);
      configService.saveConfigUrl(url);
      dispatch({ type: "CONFIG_LOAD_SUCCESS", config });
    } catch (error) {
      dispatch({ type: "CONFIG_LOAD_ERROR", error: error as AppBaseError });
    }
  }, [dispatch]);

  const loadConfigFromFile = useCallback(async (file: File) => {
    dispatch({ type: "SET_CONFIG_URL", url: `file://${file.name}` });
    dispatch({ type: "CONFIG_LOAD_START" });
    try {
      const config = await configService.loadConfigFromFile(file);
      dispatch({ type: "CONFIG_LOAD_SUCCESS", config });
    } catch (error) {
      dispatch({ type: "CONFIG_LOAD_ERROR", error: error as AppBaseError });
    }
  }, [dispatch]);

  const reloadConfig = useCallback(async () => {
    if (!state.configUrl) return;
    dispatch({ type: "CONFIG_LOAD_START" });
    try {
      const config = await configService.reloadConfig(state.configUrl);
      dispatch({ type: "CONFIG_LOAD_SUCCESS", config });
    } catch (error) {
      dispatch({ type: "CONFIG_LOAD_ERROR", error: error as AppBaseError });
    }
  }, [dispatch, state.configUrl]);

  const getSavedUrl = useCallback((): string | null => {
    return configService.getConfigUrl();
  }, []);

  return {
    config: state.config,
    configUrl: state.configUrl,
    loadConfig,
    loadConfigFromFile,
    reloadConfig,
    getSavedUrl,
  } as const;
}
