import type {
  AddinConfiguration,
  CachedConfiguration,
  InputMode,
} from "../types";
import { ConfigurationError, ConfigFetchError } from "../types/errors";

const VALID_METHODS: readonly string[] = ["GET", "POST"];
const VALID_INPUT_MODES: readonly InputMode[] = ["selected", "full", "both"];

export class ConfigService {
  private static readonly CACHE_TTL_MS = 3600000; // 1 hour
  private static readonly STORAGE_KEY_PREFIX = "word-addin-sidebar";

  /**
   * Load configuration from the given URL.
   * Returns cached configuration if the cache is valid and not expired.
   * @throws ConfigFetchError if the URL is unreachable or returns non-200.
   * @throws ConfigurationError if the configuration fails validation.
   */
  async loadConfig(url: string): Promise<AddinConfiguration> {
    this.assertHttpsUrl(url);

    const cached = this.getCachedConfig(url);
    if (cached !== null) {
      return cached;
    }

    return this.fetchAndCache(url);
  }

  /**
   * Bypass cache and fetch fresh configuration from the URL.
   * @throws ConfigFetchError if the URL is unreachable or returns non-200.
   * @throws ConfigurationError if the configuration fails validation.
   */
  async reloadConfig(url: string): Promise<AddinConfiguration> {
    this.assertHttpsUrl(url);
    return this.fetchAndCache(url);
  }

  /**
   * Validate the configuration object against the expected schema.
   * Performs a recursive, depth-first traversal.
   * @throws ConfigurationError with the field path for any validation failure.
   * No fallback values are ever substituted.
   */
  validateConfig(data: unknown): asserts data is AddinConfiguration {
    if (data === null || data === undefined || typeof data !== "object") {
      throw new ConfigurationError("root", "Configuration must be a non-null object");
    }

    const obj = data as Record<string, unknown>;

    this.assertRequiredString(obj, "configVersion", "configVersion");
    this.assertRequiredString(obj, "name", "name");

    if (!Array.isArray(obj.groups)) {
      throw new ConfigurationError("groups", "Required field 'groups' must be an array");
    }

    if (obj.groups.length === 0) {
      throw new ConfigurationError("groups", "Configuration must contain at least one group");
    }

    for (let i = 0; i < obj.groups.length; i++) {
      this.validateGroup(obj.groups[i], `groups[${i}]`);
    }
  }

  /**
   * Load configuration from a local JSON file (via File API).
   * @throws ConfigurationError if the file content fails validation.
   */
  async loadConfigFromFile(file: File): Promise<AddinConfiguration> {
    let text: string;
    try {
      text = await file.text();
    } catch (error) {
      throw new ConfigFetchError(
        `Failed to read file '${file.name}': ${(error as Error).message}`,
        undefined,
        error as Error
      );
    }

    let data: unknown;
    try {
      data = JSON.parse(text);
    } catch (error) {
      throw new ConfigFetchError(
        `File '${file.name}' contains invalid JSON: ${(error as Error).message}`,
        undefined,
        error as Error
      );
    }

    this.validateConfig(data);
    return data;
  }

  /**
   * Persist the last-used configuration URL to localStorage.
   */
  saveConfigUrl(url: string): void {
    localStorage.setItem(this.getStorageKey("config-url"), url);
  }

  /**
   * Retrieve the last-used configuration URL from localStorage.
   * Returns null if no URL was previously saved.
   */
  getConfigUrl(): string | null {
    return localStorage.getItem(this.getStorageKey("config-url"));
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  /**
   * Fetch configuration from the URL, validate it, cache it, and return it.
   */
  private async fetchAndCache(url: string): Promise<AddinConfiguration> {
    let response: Response;

    try {
      response = await fetch(url, {
        method: "GET",
        headers: { Accept: "application/json" },
      });
    } catch (error) {
      throw new ConfigFetchError(
        `Failed to fetch configuration from '${url}': ${(error as Error).message}`,
        undefined,
        error as Error
      );
    }

    if (!response.ok) {
      throw new ConfigFetchError(
        `Configuration URL '${url}' returned HTTP ${response.status}`,
        response.status
      );
    }

    let data: unknown;
    try {
      data = await response.json();
    } catch (error) {
      throw new ConfigFetchError(
        `Failed to parse JSON response from '${url}': ${(error as Error).message}`,
        undefined,
        error as Error
      );
    }

    this.validateConfig(data);
    this.cacheConfig(url, data);

    return data;
  }

  /**
   * Assert that the URL uses HTTPS (localhost is exempt for development).
   * @throws ConfigurationError if the URL does not use HTTPS.
   */
  private assertHttpsUrl(url: string): void {
    let parsed: URL;
    try {
      parsed = new URL(url);
    } catch {
      throw new ConfigurationError("url", `Invalid URL: '${url}'`);
    }

    const isLocalhost =
      parsed.hostname === "localhost" || parsed.hostname === "127.0.0.1";

    if (parsed.protocol !== "https:" && !isLocalhost) {
      throw new ConfigurationError(
        "url",
        `Configuration URL must use HTTPS. Received: '${url}'`
      );
    }
  }

  /**
   * Cache configuration to localStorage with TTL metadata.
   */
  private cacheConfig(url: string, config: AddinConfiguration): void {
    const cached: CachedConfiguration = {
      config,
      url,
      cachedAt: Date.now(),
      ttlMs: ConfigService.CACHE_TTL_MS,
    };
    try {
      localStorage.setItem(
        this.getStorageKey("addin-config-cache"),
        JSON.stringify(cached)
      );
    } catch {
      // localStorage may be full or unavailable; silently ignore cache write failures
    }
  }

  /**
   * Retrieve cached configuration if TTL has not expired and URL matches.
   * Returns null if no cache exists, the cache is expired, or the URL differs.
   */
  private getCachedConfig(url: string): AddinConfiguration | null {
    try {
      const raw = localStorage.getItem(this.getStorageKey("addin-config-cache"));
      if (raw === null) {
        return null;
      }

      const cached: CachedConfiguration = JSON.parse(raw);

      if (cached.url !== url) {
        return null;
      }

      const elapsed = Date.now() - cached.cachedAt;
      if (elapsed >= cached.ttlMs) {
        return null;
      }

      return cached.config;
    } catch {
      return null;
    }
  }

  /**
   * Generate a localStorage key with partition key prefix for Office web isolation.
   */
  private getStorageKey(suffix: string): string {
    return `${ConfigService.STORAGE_KEY_PREFIX}:${suffix}`;
  }

  // ---------------------------------------------------------------------------
  // Validation helpers
  // ---------------------------------------------------------------------------

  /**
   * Assert that a required string field exists and is non-empty.
   */
  private assertRequiredString(
    obj: Record<string, unknown>,
    field: string,
    path: string
  ): void {
    if (obj[field] === undefined || obj[field] === null) {
      throw new ConfigurationError(path, `Required field '${field}' is missing`);
    }
    if (typeof obj[field] !== "string") {
      throw new ConfigurationError(path, `Field '${field}' must be a string`);
    }
    if ((obj[field] as string).trim() === "") {
      throw new ConfigurationError(path, `Field '${field}' must not be empty`);
    }
  }

  /**
   * Validate a single ApiGroup recursively, including nested groups and APIs.
   */
  private validateGroup(data: unknown, path: string): void {
    if (data === null || data === undefined || typeof data !== "object") {
      throw new ConfigurationError(path, "Group must be a non-null object");
    }

    const group = data as Record<string, unknown>;

    this.assertRequiredString(group, "id", `${path}.id`);
    this.assertRequiredString(group, "name", `${path}.name`);

    const hasGroups = Array.isArray(group.groups) && group.groups.length > 0;
    const hasApis = Array.isArray(group.apis) && group.apis.length > 0;

    if (!hasGroups && !hasApis) {
      throw new ConfigurationError(
        path,
        "Group must contain at least one of 'groups' or 'apis'"
      );
    }

    if (Array.isArray(group.groups)) {
      for (let i = 0; i < group.groups.length; i++) {
        this.validateGroup(group.groups[i], `${path}.groups[${i}]`);
      }
    }

    if (Array.isArray(group.apis)) {
      for (let i = 0; i < group.apis.length; i++) {
        this.validateApiCall(group.apis[i], `${path}.apis[${i}]`);
      }
    }
  }

  /**
   * Validate a single ApiCallConfig object.
   */
  private validateApiCall(data: unknown, path: string): void {
    if (data === null || data === undefined || typeof data !== "object") {
      throw new ConfigurationError(path, "API call must be a non-null object");
    }

    const api = data as Record<string, unknown>;

    this.assertRequiredString(api, "id", `${path}.id`);
    this.assertRequiredString(api, "name", `${path}.name`);
    this.assertRequiredString(api, "url", `${path}.url`);
    this.assertRequiredString(api, "method", `${path}.method`);
    this.assertRequiredString(api, "inputMode", `${path}.inputMode`);

    // Validate URL is HTTPS
    const apiUrl = api.url as string;
    try {
      const parsed = new URL(apiUrl);
      const isLocalhost =
        parsed.hostname === "localhost" || parsed.hostname === "127.0.0.1";
      if (parsed.protocol !== "https:" && !isLocalhost) {
        throw new ConfigurationError(
          `${path}.url`,
          `API URL must use HTTPS. Received: '${apiUrl}'`
        );
      }
    } catch (error) {
      if (error instanceof ConfigurationError) {
        throw error;
      }
      throw new ConfigurationError(`${path}.url`, `Invalid URL: '${apiUrl}'`);
    }

    // Validate method
    const method = api.method as string;
    if (!VALID_METHODS.includes(method)) {
      throw new ConfigurationError(
        `${path}.method`,
        `Method must be one of ${VALID_METHODS.join(", ")}. Received: '${method}'`
      );
    }

    // Validate inputMode
    const inputMode = api.inputMode as string;
    if (!VALID_INPUT_MODES.includes(inputMode as InputMode)) {
      throw new ConfigurationError(
        `${path}.inputMode`,
        `inputMode must be one of ${VALID_INPUT_MODES.join(", ")}. Received: '${inputMode}'`
      );
    }

    // Validate required timeout
    if (api.timeout === undefined || api.timeout === null) {
      throw new ConfigurationError(
        `${path}.timeout`,
        "Required field 'timeout' is missing. Each API call must have an explicit timeout configured."
      );
    }
    if (typeof api.timeout !== "number" || api.timeout <= 0) {
      throw new ConfigurationError(
        `${path}.timeout`,
        "timeout must be a positive number (milliseconds)"
      );
    }
  }
}
