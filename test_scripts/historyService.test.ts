import { describe, it, expect, beforeEach, afterEach } from "vitest";
import "fake-indexeddb/auto";
import { HistoryService } from "../src/taskpane/services/HistoryService";
import { HistoryStorageError } from "../src/taskpane/types/errors";
import type { HistoryEntry } from "../src/taskpane/types/history";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeEntry(overrides?: Partial<HistoryEntry>): HistoryEntry {
  return {
    id: `entry-${Math.random().toString(36).slice(2, 10)}`,
    timestamp: Date.now(),
    apiId: "api-1",
    apiName: "Test API",
    apiUrl: "https://api.example.com/v1/test",
    prompt: "Summarize",
    textSource: "selected",
    textPreview: "Some text...",
    documentName: "report.docx",
    wasSuccessful: true,
    responsePreview: "Summary result",
    durationMs: 450,
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("HistoryService", () => {
  let service: HistoryService;

  beforeEach(async () => {
    service = new HistoryService();
  });

  afterEach(() => {
    // Reset indexedDB databases between tests
    indexedDB = new IDBFactory();
  });

  // =========================================================================
  // initialize
  // =========================================================================

  describe("initialize()", () => {
    it("creates IndexedDB database", async () => {
      await expect(service.initialize()).resolves.toBeUndefined();
    });
  });

  // =========================================================================
  // Methods throw when not initialized
  // =========================================================================

  describe("not initialized", () => {
    it("addEntry throws HistoryStorageError when not initialized", async () => {
      await expect(service.addEntry(makeEntry())).rejects.toThrow(HistoryStorageError);
    });

    it("getEntries throws HistoryStorageError when not initialized", async () => {
      await expect(service.getEntries()).rejects.toThrow(HistoryStorageError);
    });

    it("deleteEntry throws HistoryStorageError when not initialized", async () => {
      await expect(service.deleteEntry("some-id")).rejects.toThrow(HistoryStorageError);
    });

    it("clearAll throws HistoryStorageError when not initialized", async () => {
      await expect(service.clearAll()).rejects.toThrow(HistoryStorageError);
    });
  });

  // =========================================================================
  // addEntry / getEntries
  // =========================================================================

  describe("addEntry() / getEntries()", () => {
    beforeEach(async () => {
      await service.initialize();
    });

    it("stores entry and can be retrieved", async () => {
      const entry = makeEntry({ id: "e1" });
      await service.addEntry(entry);

      const entries = await service.getEntries();
      expect(entries).toHaveLength(1);
      expect(entries[0].id).toBe("e1");
    });

    it("returns entries sorted by timestamp descending", async () => {
      const now = Date.now();
      const entry1 = makeEntry({ id: "e1", timestamp: now - 2000 });
      const entry2 = makeEntry({ id: "e2", timestamp: now });
      const entry3 = makeEntry({ id: "e3", timestamp: now - 1000 });

      await service.addEntry(entry1);
      await service.addEntry(entry2);
      await service.addEntry(entry3);

      const entries = await service.getEntries();
      expect(entries.map((e) => e.id)).toEqual(["e2", "e3", "e1"]);
    });

    it("filters by apiId", async () => {
      await service.addEntry(makeEntry({ id: "e1", apiId: "api-alpha" }));
      await service.addEntry(makeEntry({ id: "e2", apiId: "api-beta" }));
      await service.addEntry(makeEntry({ id: "e3", apiId: "api-alpha" }));

      const entries = await service.getEntries({ apiId: "api-alpha" });
      expect(entries).toHaveLength(2);
      expect(entries.every((e) => e.apiId === "api-alpha")).toBe(true);
    });

    it("filters by documentName", async () => {
      await service.addEntry(makeEntry({ id: "e1", documentName: "doc-a.docx" }));
      await service.addEntry(makeEntry({ id: "e2", documentName: "doc-b.docx" }));

      const entries = await service.getEntries({ documentName: "doc-a.docx" });
      expect(entries).toHaveLength(1);
      expect(entries[0].documentName).toBe("doc-a.docx");
    });

    it("filters by wasSuccessful", async () => {
      await service.addEntry(makeEntry({ id: "e1", wasSuccessful: true }));
      await service.addEntry(makeEntry({ id: "e2", wasSuccessful: false }));
      await service.addEntry(makeEntry({ id: "e3", wasSuccessful: true }));

      const failures = await service.getEntries({ wasSuccessful: false });
      expect(failures).toHaveLength(1);
      expect(failures[0].id).toBe("e2");
    });

    it("respects limit", async () => {
      const now = Date.now();
      await service.addEntry(makeEntry({ id: "e1", timestamp: now - 2000 }));
      await service.addEntry(makeEntry({ id: "e2", timestamp: now - 1000 }));
      await service.addEntry(makeEntry({ id: "e3", timestamp: now }));

      const entries = await service.getEntries({ limit: 2 });
      expect(entries).toHaveLength(2);
      // Should be the 2 newest
      expect(entries.map((e) => e.id)).toEqual(["e3", "e2"]);
    });
  });

  // =========================================================================
  // deleteEntry
  // =========================================================================

  describe("deleteEntry()", () => {
    beforeEach(async () => {
      await service.initialize();
    });

    it("removes specific entry", async () => {
      await service.addEntry(makeEntry({ id: "e1" }));
      await service.addEntry(makeEntry({ id: "e2" }));

      await service.deleteEntry("e1");

      const entries = await service.getEntries();
      expect(entries).toHaveLength(1);
      expect(entries[0].id).toBe("e2");
    });
  });

  // =========================================================================
  // clearAll
  // =========================================================================

  describe("clearAll()", () => {
    beforeEach(async () => {
      await service.initialize();
    });

    it("removes all entries", async () => {
      await service.addEntry(makeEntry({ id: "e1" }));
      await service.addEntry(makeEntry({ id: "e2" }));
      await service.addEntry(makeEntry({ id: "e3" }));

      await service.clearAll();

      const entries = await service.getEntries();
      expect(entries).toHaveLength(0);
    });
  });

  // =========================================================================
  // auto-prune
  // =========================================================================

  describe("auto-prune", () => {
    beforeEach(async () => {
      await service.initialize();
    });

    it("addEntry triggers auto-prune (entries are not lost for small counts)", async () => {
      // With fewer entries than MAX_ENTRIES, prune should not remove anything
      for (let i = 0; i < 5; i++) {
        await service.addEntry(
          makeEntry({ id: `e${i}`, timestamp: Date.now() + i })
        );
      }

      const entries = await service.getEntries();
      expect(entries).toHaveLength(5);
    });
  });
});
