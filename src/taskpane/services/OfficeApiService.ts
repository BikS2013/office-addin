import { DocumentTextInfo, OfficeApiError } from "../types";

/**
 * OfficeApiService
 *
 * Interacts with the Word document through Office.js Word API.
 * All Office.js calls are wrapped in try/catch and throw OfficeApiError on failure.
 *
 * The Word and Office namespaces are globally available at runtime from @types/office-js.
 */
export class OfficeApiService {
  /**
   * Extract the currently selected text from the active Word document.
   * Returns an empty string if the cursor has no selection.
   * @throws OfficeApiError if the Word API call fails.
   */
  async getSelectedText(): Promise<string> {
    try {
      return await Word.run(async (context: Word.RequestContext) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        return selection.text;
      });
    } catch (error) {
      throw new OfficeApiError(
        `Failed to extract selected text: ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Extract the full text content of the active Word document body.
   * @throws OfficeApiError if the Word API call fails.
   */
  async getFullDocumentText(): Promise<string> {
    try {
      return await Word.run(async (context: Word.RequestContext) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        return body.text;
      });
    } catch (error) {
      throw new OfficeApiError(
        `Failed to extract full document text: ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Get the active document filename.
   * Uses Office.context.document.url which is synchronous and does not require Word.run.
   * Returns "Untitled" if no path is available.
   * @throws OfficeApiError if the document properties cannot be read.
   */
  async getDocumentName(): Promise<string> {
    try {
      const documentUrl: string | undefined = Office.context.document.url;

      if (!documentUrl) {
        return "Untitled";
      }

      // The URL may be a file path or a URL; extract the filename from either format
      const separatorIndex = Math.max(
        documentUrl.lastIndexOf("/"),
        documentUrl.lastIndexOf("\\")
      );

      const filename =
        separatorIndex >= 0
          ? documentUrl.substring(separatorIndex + 1)
          : documentUrl;

      return filename || "Untitled";
    } catch (error) {
      throw new OfficeApiError(
        `Failed to get document name: ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Get complete document text information including selection status.
   * Combines getSelectedText and getDocumentName results.
   * @throws OfficeApiError if any Word API call fails.
   */
  async getDocumentTextInfo(): Promise<DocumentTextInfo> {
    const [selectedText, documentName] = await Promise.all([
      this.getSelectedText(),
      this.getDocumentName(),
    ]);

    const hasSelection = selectedText.trim().length > 0;

    return {
      selectedText,
      hasSelection,
      documentName,
    };
  }

  /**
   * Insert text at the current cursor position, replacing any selection.
   * @throws OfficeApiError if the insert operation fails.
   */
  async insertAtCursor(text: string): Promise<void> {
    try {
      await Word.run(async (context: Word.RequestContext) => {
        const selection = context.document.getSelection();
        selection.insertText(text, Word.InsertLocation.replace);
        await context.sync();
      });
    } catch (error) {
      throw new OfficeApiError(
        `Failed to insert text at cursor: ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Append text at the end of the document body.
   * @throws OfficeApiError if the append operation fails.
   */
  async appendToDocument(text: string): Promise<void> {
    try {
      await Word.run(async (context: Word.RequestContext) => {
        const body = context.document.body;
        body.insertText(text, Word.InsertLocation.end);
        await context.sync();
      });
    } catch (error) {
      throw new OfficeApiError(
        `Failed to append text to document: ${error instanceof Error ? error.message : String(error)}`,
        error instanceof Error ? error : undefined
      );
    }
  }
}
