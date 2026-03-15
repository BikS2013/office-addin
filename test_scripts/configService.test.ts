import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { ConfigService } from "../src/taskpane/services/ConfigService";
import { ConfigurationError, ConfigFetchError } from "../src/taskpane/types/errors";
import type { AddinConfiguration } from "../src/taskpane/types/config";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeValidConfig(overrides?: Partial<AddinConfiguration>): AddinConfiguration {
  return {
    configVersion: "1.0",
    name: "Test Config",
    groups: [
      {
        id: "group-1",
        name: "Test Group",
        apis: [
          {
            id: "api-1",
            name: "Test API",
            url: "https://api.example.com/v1/test",
            method: "POST",
            inputMode: "selected",
            timeout: 30000,
            bodyTemplate: { prompt: "{{prompt}}", text: "{{text}}" },
            responseField: "result",
          },
        ],
      },
    ],
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// localStorage mock
// ---------------------------------------------------------------------------

const localStorageMock = (() => {
  let store: Record<string, string> = {};
  return {
    getItem: vi.fn((key: string) => store[key] ?? null),
    setItem: vi.fn((key: string, value: string) => {
      store[key] = value;
    }),
    removeItem: vi.fn((key: string) => {
      delete store[key];
    }),
    clear: vi.fn(() => {
      store = {};
    }),
    get length() {
      return Object.keys(store).length;
    },
    key: vi.fn((_index: number) => null),
  } as unknown as Storage;
})();

Object.defineProperty(globalThis, "localStorage", { value: localStorageMock, writable: true });

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("ConfigService", () => {
  let service: ConfigService;

  beforeEach(() => {
    service = new ConfigService();
    localStorageMock.clear();
    vi.restoreAllMocks();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // =========================================================================
  // validateConfig
  // =========================================================================

  describe("validateConfig()", () => {
    it("succeeds with valid configuration", () => {
      const config = makeValidConfig();
      expect(() => service.validateConfig(config)).not.toThrow();
    });

    it("throws ConfigurationError for missing configVersion", () => {
      const config = { ...makeValidConfig(), configVersion: undefined } as unknown;
      expect(() => service.validateConfig(config)).toThrow(ConfigurationError);
    });

    it("throws ConfigurationError for missing group id", () => {
      const config = {
        configVersion: "1.0",
        name: "Test",
        groups: [
          {
            name: "Group without ID",
            apis: [
              {
                id: "api-1",
                name: "API 1",
                url: "https://api.example.com/test",
                method: "POST",
                inputMode: "selected",
                timeout: 5000,
              },
            ],
          },
        ],
      };
      expect(() => service.validateConfig(config)).toThrow(ConfigurationError);
    });

    it("throws ConfigurationError for invalid inputMode", () => {
      const config = {
        configVersion: "1.0",
        name: "Test",
        groups: [
          {
            id: "g1",
            name: "G1",
            apis: [
              {
                id: "api-1",
                name: "API",
                url: "https://api.example.com/test",
                method: "POST",
                inputMode: "invalid-mode",
                timeout: 5000,
              },
            ],
          },
        ],
      };
      expect(() => service.validateConfig(config)).toThrow(ConfigurationError);
    });

    it("throws ConfigurationError for non-HTTPS URL", () => {
      const config = {
        configVersion: "1.0",
        name: "Test",
        groups: [
          {
            id: "g1",
            name: "G1",
            apis: [
              {
                id: "api-1",
                name: "API",
                url: "http://api.example.com/test",
                method: "POST",
                inputMode: "selected",
                timeout: 5000,
              },
            ],
          },
        ],
      };
      expect(() => service.validateConfig(config)).toThrow(ConfigurationError);
    });

    it("throws ConfigurationError for invalid method", () => {
      const config = {
        configVersion: "1.0",
        name: "Test",
        groups: [
          {
            id: "g1",
            name: "G1",
            apis: [
              {
                id: "api-1",
                name: "API",
                url: "https://api.example.com/test",
                method: "DELETE",
                inputMode: "selected",
                timeout: 5000,
              },
            ],
          },
        ],
      };
      expect(() => service.validateConfig(config)).toThrow(ConfigurationError);
    });

    it("throws ConfigurationError for missing timeout", () => {
      const config = {
        configVersion: "1.0",
        name: "Test",
        groups: [
          {
            id: "g1",
            name: "G1",
            apis: [
              {
                id: "api-1",
                name: "API",
                url: "https://api.example.com/test",
                method: "POST",
                inputMode: "selected",
              },
            ],
          },
        ],
      };
      expect(() => service.validateConfig(config)).toThrow(ConfigurationError);
    });

    it("throws ConfigurationError for negative timeout", () => {
      const config = {
        configVersion: "1.0",
        name: "Test",
        groups: [
          {
            id: "g1",
            name: "G1",
            apis: [
              {
                id: "api-1",
                name: "API",
                url: "https://api.example.com/test",
                method: "POST",
                inputMode: "selected",
                timeout: -100,
              },
            ],
          },
        ],
      };
      expect(() => service.validateConfig(config)).toThrow(ConfigurationError);
    });

    it("validates nested groups recursively", () => {
      const config = {
        configVersion: "1.0",
        name: "Nested Config",
        groups: [
          {
            id: "outer",
            name: "Outer Group",
            groups: [
              {
                id: "inner",
                name: "Inner Group",
                apis: [
                  {
                    id: "api-nested",
                    name: "Nested API",
                    url: "https://api.example.com/nested",
                    method: "GET",
                    inputMode: "full",
                    timeout: 10000,
                  },
                ],
              },
            ],
          },
        ],
      };
      expect(() => service.validateConfig(config)).not.toThrow();
    });

    it("throws ConfigurationError for invalid nested group (missing id in inner group)", () => {
      const config = {
        configVersion: "1.0",
        name: "Nested Config",
        groups: [
          {
            id: "outer",
            name: "Outer Group",
            groups: [
              {
                name: "Inner Group Missing ID",
                apis: [
                  {
                    id: "api-nested",
                    name: "Nested API",
                    url: "https://api.example.com/nested",
                    method: "GET",
                    inputMode: "full",
                    timeout: 10000,
                  },
                ],
              },
            ],
          },
        ],
      };
      expect(() => service.validateConfig(config)).toThrow(ConfigurationError);
    });
  });

  // =========================================================================
  // loadConfig
  // =========================================================================

  describe("loadConfig()", () => {
    it("fetches and returns valid config", async () => {
      const validConfig = makeValidConfig();

      vi.stubGlobal(
        "fetch",
        vi.fn().mockResolvedValue({
          ok: true,
          status: 200,
          json: () => Promise.resolve(validConfig),
        })
      );

      const result = await service.loadConfig("https://config.example.com/config.json");

      expect(result).toEqual(validConfig);
      expect(fetch).toHaveBeenCalledOnce();
    });

    it("throws ConfigFetchError on network error", async () => {
      vi.stubGlobal(
        "fetch",
        vi.fn().mockRejectedValue(new Error("Network failure"))
      );

      await expect(
        service.loadConfig("https://config.example.com/config.json")
      ).rejects.toThrow(ConfigFetchError);
    });

    it("uses cache when within TTL", async () => {
      const validConfig = makeValidConfig();

      vi.stubGlobal(
        "fetch",
        vi.fn().mockResolvedValue({
          ok: true,
          status: 200,
          json: () => Promise.resolve(validConfig),
        })
      );

      // First load populates cache
      await service.loadConfig("https://config.example.com/config.json");
      expect(fetch).toHaveBeenCalledTimes(1);

      // Second load should use cache -- no additional fetch
      const cached = await service.loadConfig("https://config.example.com/config.json");
      expect(cached).toEqual(validConfig);
      expect(fetch).toHaveBeenCalledTimes(1);
    });

    it("throws ConfigurationError for non-HTTPS URL", async () => {
      await expect(
        service.loadConfig("http://insecure.example.com/config.json")
      ).rejects.toThrow(ConfigurationError);
    });
  });

  // =========================================================================
  // reloadConfig
  // =========================================================================

  describe("reloadConfig()", () => {
    it("bypasses cache and fetches fresh configuration", async () => {
      const configV1 = makeValidConfig({ name: "Version 1" });
      const configV2 = makeValidConfig({ name: "Version 2" });

      let callCount = 0;
      vi.stubGlobal(
        "fetch",
        vi.fn().mockImplementation(() => {
          callCount++;
          const config = callCount === 1 ? configV1 : configV2;
          return Promise.resolve({
            ok: true,
            status: 200,
            json: () => Promise.resolve(config),
          });
        })
      );

      // First load
      await service.loadConfig("https://config.example.com/config.json");
      expect(fetch).toHaveBeenCalledTimes(1);

      // Reload bypasses cache
      const result = await service.reloadConfig("https://config.example.com/config.json");
      expect(fetch).toHaveBeenCalledTimes(2);
      expect(result.name).toBe("Version 2");
    });
  });

  // =========================================================================
  // saveConfigUrl / getConfigUrl
  // =========================================================================

  describe("saveConfigUrl() / getConfigUrl()", () => {
    it("round-trips correctly", () => {
      const url = "https://config.example.com/my-config.json";
      service.saveConfigUrl(url);
      expect(service.getConfigUrl()).toBe(url);
    });

    it("returns null when no URL was saved", () => {
      expect(service.getConfigUrl()).toBeNull();
    });
  });
});
