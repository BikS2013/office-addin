import { openDB, IDBPDatabase } from "idb";
import { HistoryEntry, HistoryFilter, HistoryStorageError } from "../types";

export class HistoryService {
  private static readonly DB_NAME = "word-addin-history";
  private static readonly DB_VERSION = 1;
  private static readonly STORE_NAME = "entries";
  private static readonly MAX_ENTRIES = 1000;
  private static readonly MAX_AGE_DAYS = 90;

  private db: IDBPDatabase | null = null;

  /**
   * Initialize the IndexedDB database connection.
   * Creates the object store and indexes if they do not exist.
   * @throws HistoryStorageError if IndexedDB is unavailable or initialization fails.
   */
  async initialize(): Promise<void> {
    try {
      this.db = await openDB(HistoryService.DB_NAME, HistoryService.DB_VERSION, {
        upgrade(db) {
          const store = db.createObjectStore(HistoryService.STORE_NAME, {
            keyPath: "id",
          });
          store.createIndex("timestamp", "timestamp");
          store.createIndex("apiId", "apiId");
          store.createIndex("documentName", "documentName");
        },
      });
    } catch (error) {
      throw new HistoryStorageError(
        `Failed to initialize IndexedDB: ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Add a new history entry.
   * Triggers auto-pruning after the entry is added.
   * @throws HistoryStorageError if not initialized or the write operation fails.
   */
  async addEntry(entry: HistoryEntry): Promise<void> {
    const db = this.ensureDb();
    try {
      await db.add(HistoryService.STORE_NAME, entry);
    } catch (error) {
      throw new HistoryStorageError(
        `Failed to add history entry: ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
    await this.prune();
  }

  /**
   * Retrieve history entries ordered by timestamp descending (newest first).
   * Applies optional filters.
   * @throws HistoryStorageError if not initialized or the read operation fails.
   */
  async getEntries(filter?: HistoryFilter): Promise<readonly HistoryEntry[]> {
    const db = this.ensureDb();
    try {
      const allEntries: HistoryEntry[] = await db.getAll(HistoryService.STORE_NAME);

      allEntries.sort((a, b) => b.timestamp - a.timestamp);

      let filtered = allEntries;

      if (filter) {
        if (filter.apiId !== undefined) {
          filtered = filtered.filter((e) => e.apiId === filter.apiId);
        }
        if (filter.documentName !== undefined) {
          filtered = filtered.filter((e) => e.documentName === filter.documentName);
        }
        if (filter.fromTimestamp !== undefined) {
          const from = filter.fromTimestamp;
          filtered = filtered.filter((e) => e.timestamp >= from);
        }
        if (filter.toTimestamp !== undefined) {
          const to = filter.toTimestamp;
          filtered = filtered.filter((e) => e.timestamp <= to);
        }
        if (filter.wasSuccessful !== undefined) {
          const success = filter.wasSuccessful;
          filtered = filtered.filter((e) => e.wasSuccessful === success);
        }
        if (filter.limit !== undefined) {
          filtered = filtered.slice(0, filter.limit);
        }
      }

      return filtered;
    } catch (error) {
      if (error instanceof HistoryStorageError) {
        throw error;
      }
      throw new HistoryStorageError(
        `Failed to retrieve history entries: ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Delete a single history entry by ID.
   * @throws HistoryStorageError if not initialized or the delete operation fails.
   */
  async deleteEntry(id: string): Promise<void> {
    const db = this.ensureDb();
    try {
      await db.delete(HistoryService.STORE_NAME, id);
    } catch (error) {
      throw new HistoryStorageError(
        `Failed to delete history entry '${id}': ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Delete all history entries.
   * @throws HistoryStorageError if not initialized or the clear operation fails.
   */
  async clearAll(): Promise<void> {
    const db = this.ensureDb();
    try {
      await db.clear(HistoryService.STORE_NAME);
    } catch (error) {
      throw new HistoryStorageError(
        `Failed to clear history: ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Remove entries older than MAX_AGE_DAYS and entries exceeding MAX_ENTRIES.
   * Called automatically after addEntry.
   */
  private async prune(): Promise<void> {
    const db = this.ensureDb();
    try {
      const cutoffTimestamp = Date.now() - HistoryService.MAX_AGE_DAYS * 24 * 60 * 60 * 1000;

      // Delete entries older than MAX_AGE_DAYS using the timestamp index
      const tx = db.transaction(HistoryService.STORE_NAME, "readwrite");
      const index = tx.store.index("timestamp");
      let cursor = await index.openCursor(IDBKeyRange.upperBound(cutoffTimestamp, true));

      while (cursor) {
        await cursor.delete();
        cursor = await cursor.continue();
      }

      await tx.done;

      // If still more than MAX_ENTRIES, delete oldest to get down to MAX_ENTRIES
      const allEntries: HistoryEntry[] = await db.getAll(HistoryService.STORE_NAME);

      if (allEntries.length > HistoryService.MAX_ENTRIES) {
        allEntries.sort((a, b) => b.timestamp - a.timestamp);

        const entriesToDelete = allEntries.slice(HistoryService.MAX_ENTRIES);
        const deleteTx = db.transaction(HistoryService.STORE_NAME, "readwrite");

        for (const entry of entriesToDelete) {
          await deleteTx.store.delete(entry.id);
        }

        await deleteTx.done;
      }
    } catch (error) {
      // Pruning failures are non-critical; log but do not throw
      // to avoid disrupting the addEntry caller.
      console.warn(
        `History auto-prune failed: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Ensure the database connection is initialized.
   * @throws HistoryStorageError if not initialized.
   */
  private ensureDb(): IDBPDatabase {
    if (!this.db) {
      throw new HistoryStorageError(
        "HistoryService is not initialized. Call initialize() before using the service."
      );
    }
    return this.db;
  }
}
