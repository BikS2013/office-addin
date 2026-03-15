# Investigation: Microsoft Word Add-in with Sidebar (Task Pane) Functionality

**Date:** 2026-02-25
**Status:** Complete
**Objective:** Research and document how to build a Microsoft Word add-in implemented as a sidebar (task pane), capable of dynamic API configuration loading, text extraction, and prompt history management.

---

## Table of Contents

1. [Executive Summary](#executive-summary)
2. [Office Add-in Architecture](#1-office-add-in-architecture)
3. [Scaffolding Tools Comparison](#2-scaffolding-tools-comparison)
4. [Office.js Word API -- Text Extraction](#3-officejs-word-api----text-extraction)
5. [UI Framework -- Fluent UI React](#4-ui-framework----fluent-ui-react)
6. [Sidebar / Task Pane Configuration](#5-sidebar--task-pane-configuration)
7. [Configuration Loading from Remote URL](#6-configuration-loading-from-remote-url)
8. [Local Storage and Persistence](#7-local-storage-and-persistence)
9. [Sideloading and Debugging](#8-sideloading-and-debugging)
10. [Deployment Options](#9-deployment-options)
11. [Known Limitations](#10-known-limitations)
12. [Comparison Matrix -- Scaffolding Tools](#comparison-matrix----scaffolding-tools)
13. [Recommended Approach](#recommended-approach)
14. [Key Code Patterns](#key-code-patterns)
15. [References](#references)

---

## Executive Summary

A Microsoft Word Office Add-in is a web application (HTML/CSS/JavaScript or TypeScript) hosted inside a **task pane** -- a sidebar panel that appears on the right side of the Word window. The add-in communicates with the Word document through the **Office.js JavaScript API** (`Word.run()` context).

**Key findings:**

- **Architecture:** Office Add-ins are web apps running inside a sandboxed webview (Edge Chromium on Windows, WebKit on Mac). They use a manifest file to declare capabilities, ribbon buttons, and task pane entry points.
- **Manifest format:** Two options exist -- the legacy **XML manifest** (production-ready for all hosts) and the newer **unified JSON manifest** (production for Outlook, preview for Word/Excel/PowerPoint). The XML manifest is the safer choice for Word add-ins today.
- **Scaffolding:** The **Yeoman Office Generator** (`yo office`) is the most mature tool for creating a React + TypeScript Word task pane add-in. Teams Toolkit is an alternative but primarily targets Outlook.
- **Text extraction:** Office.js provides `context.document.getSelection()` for selected text and `context.document.body` for full document text, both using the proxy object + `load()` + `sync()` pattern.
- **UI framework:** **Fluent UI React v9** is the official Microsoft UI framework for Office Add-ins, providing components that match the Office visual style.
- **Storage:** `localStorage` (5MB limit), `IndexedDB` (larger capacity), and `Office.Settings` (document-bound) are all available. For cross-session prompt history, `localStorage` or `IndexedDB` are the recommended approaches.
- **Development:** `npm start` handles sideloading and local HTTPS dev server setup automatically when using the Yeoman-generated project.
- **Deployment:** Centralized Deployment via Microsoft 365 Admin Center is recommended for organizational distribution. AppSource (Microsoft Marketplace) for public distribution.

---

## 1. Office Add-in Architecture

### Runtime Model

Office Add-ins are web applications that run inside a **sandboxed webview control** embedded within the Office application:

| Platform | Webview Engine |
|----------|---------------|
| Windows (Microsoft 365) | Edge Chromium (WebView2) |
| macOS | WebKit (WKWebView) |
| Office on the web | Browser iframe (sandboxed) |
| Older Windows (Office 2019, perpetual) | Edge Legacy or Internet Explorer 11 (Trident) |

The add-in runs in an **isolated process** separate from the Office application. Communication with the document happens exclusively through the Office.js API bridge.

### Core Components

```
+---------------------------+
|    Office Application     |
|   (Word, Excel, etc.)     |
|                           |
|  +---------------------+  |
|  |    Task Pane         |  |
|  |  (Webview Control)   |  |
|  |                      |  |
|  |  HTML + CSS + JS/TS  |  |
|  |  + Office.js         |  |
|  |  + Fluent UI React   |  |
|  +---------------------+  |
|                           |
+---------------------------+
         |
         | Office.js API Bridge
         |
+---------------------------+
|   Word Document Model     |
|   (Body, Paragraphs,      |
|    Ranges, Selections)    |
+---------------------------+
```

### Manifest File

The manifest declares:
- Add-in identity (ID, version, name, description)
- Permissions required (e.g., `ReadWriteDocument`)
- Ribbon button configuration (icons, labels, actions)
- Task pane source URL (the HTML entry point)
- Supported Office hosts and form factors

### Manifest Format Options

| Feature | XML Manifest (Add-in Only) | Unified JSON Manifest |
|---------|---------------------------|----------------------|
| **Format** | XML | JSON |
| **Word support** | Production | Preview |
| **Outlook support** | Production | Production |
| **Teams integration** | No | Yes (single package) |
| **Feature parity** | Full | Gaps remain |
| **Tooling support** | Yeoman, VS, VS Code | Yeoman, Teams Toolkit |
| **Recommended for Word** | **Yes** | Not yet for production |

**Decision: Use XML manifest** for the Word add-in until the unified JSON manifest reaches GA for Word.

---

## 2. Scaffolding Tools Comparison

### Option A: Yeoman Office Generator (`yo office`)

The official CLI scaffolding tool for Office Add-ins.

**Setup:**
```bash
npm install -g yo generator-office
yo office
```

**Interactive prompts for our use case:**
```
? Choose a project type: Office Add-in Task Pane project
? Choose a script type: TypeScript
? What do you want to name your add-in? Word API Sidebar
? Which Office client application would you like to support? Word
```

**Or non-interactive:**
```bash
yo office --projectType react --name "Word API Sidebar" --host word --ts true
```

**Pros:**
- Mature and well-documented
- Full support for React + TypeScript
- Generates XML or unified JSON manifest
- Built-in webpack configuration
- Automatic HTTPS certificate generation
- `npm start` handles sideloading

**Cons:**
- CLI-only (no GUI)
- Generated project uses webpack (not Vite)

### Option B: Teams Toolkit (VS Code Extension)

**Pros:**
- GUI-based workflow in VS Code
- Integrated debugging experience

**Cons:**
- Primarily targets Outlook add-ins with unified manifest
- Word support is through the "Office Add-ins Development Kit" (preview)
- Less mature for Word-specific development

### Option C: Manual Setup

**Pros:**
- Full control over toolchain (Vite, esbuild, etc.)
- Can use manifest-only Yeoman option: `yo office --projectType manifest`

**Cons:**
- Significant manual configuration required
- Must handle HTTPS, webpack/Vite config, and Office.js integration manually

---

## 3. Office.js Word API -- Text Extraction

### Proxy Object Pattern

Office.js uses a **proxy object pattern** with batch operations. You must:
1. Request a context via `Word.run()`
2. Create proxy objects for document elements
3. Call `.load("property")` to queue property reads
4. Call `context.sync()` to execute the batch
5. Read values from proxy objects after sync

### Get Selected Text

```typescript
async function getSelectedText(): Promise<string> {
  return await Word.run(async (context: Word.RequestContext) => {
    const selection: Word.Range = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text;
  });
}
```

### Get Full Document Text

```typescript
async function getFullDocumentText(): Promise<string> {
  return await Word.run(async (context: Word.RequestContext) => {
    const body: Word.Body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text;
  });
}
```

### Get Selected Text (Common API -- Alternative)

```typescript
function getSelectedTextCommonAPI(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      (result: Office.AsyncResult<string>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}
```

### Insert/Replace Text at Selection

```typescript
async function replaceSelection(newText: string): Promise<void> {
  await Word.run(async (context: Word.RequestContext) => {
    const selection: Word.Range = context.document.getSelection();
    selection.insertText(newText, Word.InsertLocation.replace);
    await context.sync();
  });
}
```

### Insert Text at End of Document

```typescript
async function appendToDocument(text: string): Promise<void> {
  await Word.run(async (context: Word.RequestContext) => {
    const body: Word.Body = context.document.body;
    body.insertText(text, Word.InsertLocation.end);
    await context.sync();
  });
}
```

### Check if Text is Selected

```typescript
async function hasSelection(): Promise<boolean> {
  return await Word.run(async (context: Word.RequestContext) => {
    const selection: Word.Range = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text.length > 0;
  });
}
```

### Known Issues with Text Extraction

| Issue | Description | Workaround |
|-------|-------------|------------|
| Full document selection (Ctrl+A) | `range.insertText()` with Replace may throw exception | Use `body.insertText()` instead |
| Table row selection | `getSelection()` may return corrupted results for multi-row table selections | Limit to single row or full table |
| Mac `insertText` bug | `insertText` was reported not working on Mac desktop in some versions | Ensure Office is updated |
| Formatting loss on Replace | Replacing a hyperlink range may cause formatting to bleed from preceding content | Insert after clearing, then reformat |

---

## 4. UI Framework -- Fluent UI React

### Fluent UI React v9 (Recommended)

Fluent UI React v9 is the official open-source UI framework for Office Add-ins. It provides components that visually integrate with the Office experience.

**Installation:**
```bash
npm install @fluentui/react-components
```

**Basic Setup in Task Pane:**
```tsx
import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./App";

// Ensure Office.js is ready before rendering
Office.onReady(() => {
  const root = createRoot(document.getElementById("root")!);
  root.render(
    <FluentProvider theme={webLightTheme}>
      <App />
    </FluentProvider>
  );
});
```

**Key Components for Our Add-in:**

```tsx
import {
  Button,
  Input,
  Textarea,
  Dropdown,
  Option,
  Tree,
  TreeItem,
  TreeItemLayout,
  Card,
  Spinner,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Toolbar,
  ToolbarButton,
  Tab,
  TabList,
} from "@fluentui/react-components";
```

### Version Compatibility

| Fluent UI Version | Trident (IE11) Support | Recommended |
|---|---|---|
| v9 | No | Yes (for modern Office) |
| v8 | Yes | Only if legacy support needed |

**Decision: Use Fluent UI React v9.** Trident/IE11 support is not needed for new development targeting Microsoft 365.

---

## 5. Sidebar / Task Pane Configuration

### Task Pane Dimensions

At 1366x768 screen resolution with the Office ribbon visible:

| Application | Task Pane Size (pixels) |
|-------------|------------------------|
| Word | 329 x 445 |
| Excel | 350 x 378 |
| PowerPoint | 348 x 391 |
| Outlook (web) | 320 x 570 |

### Resizing Behavior

- Users can manually drag the task pane border to resize.
- Office remembers the last size and position between sessions.
- **Programmatic resizing is NOT supported** in OfficeJS web add-ins (only available in VSTO).
- Default width is approximately 320 pixels.
- The personality menu (close button) occupies 12x32px (Windows) or 34x32px (Mac) in the top-right corner.

### Manifest Configuration for Task Pane (XML)

The task pane is configured in the manifest via `<VersionOverrides>` and the `ShowTaskpane` action:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <!-- Basic identity -->
  <Id>YOUR-GUID-HERE</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Your Company</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Word API Sidebar" />
  <Description DefaultValue="A configurable API sidebar for Word" />

  <!-- Task pane entry point -->
  <Hosts>
    <Host Name="Document" /> <!-- Word -->
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <!-- Version Overrides for ribbon customization -->
  <VersionOverrides
    xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
    xsi:type="VersionOverridesV1_0">

    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>

          <!-- Commands file (for ExecuteFunction actions) -->
          <FunctionFile resid="Commands.Url" />

          <!-- Ribbon extension point -->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16" />
                  <bt:Image size="32" resid="Icon.32x32" />
                  <bt:Image size="80" resid="Icon.80x80" />
                </Icon>

                <!-- Button that opens the task pane -->
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>

              </Group>
            </OfficeTab>
          </ExtensionPoint>

        </DesktopFormFactor>
      </Host>
    </Hosts>

    <!-- Resource strings and URLs -->
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16"
          DefaultValue="https://localhost:3000/assets/icon-16.png" />
        <bt:Image id="Icon.32x32"
          DefaultValue="https://localhost:3000/assets/icon-32.png" />
        <bt:Image id="Icon.80x80"
          DefaultValue="https://localhost:3000/assets/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url"
          DefaultValue="https://localhost:3000/commands.html" />
        <bt:Url id="Taskpane.Url"
          DefaultValue="https://localhost:3000/taskpane.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="CommandsGroup.Label"
          DefaultValue="API Sidebar" />
        <bt:String id="TaskpaneButton.Label"
          DefaultValue="Open Sidebar" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip"
          DefaultValue="Open the API configuration sidebar" />
      </bt:LongStrings>
    </Resources>

  </VersionOverrides>
</OfficeApp>
```

### Responsive Design Guidelines

Since task panes are narrow (~320px), the UI must be designed for a mobile-like viewport:

- Use single-column layouts
- Avoid horizontal scrolling
- Use collapsible sections (Accordion) for API groups
- Keep buttons full-width
- Use compact Fluent UI component variants

---

## 6. Configuration Loading from Remote URL

### Approach: Fetch JSON Configuration at Runtime

Since the add-in runs as a web app, standard `fetch()` API is available for loading remote configuration:

```typescript
interface ApiCallConfig {
  name: string;
  url: string;
  method: "GET" | "POST";
  inputType: "selected" | "full" | "both";
  promptTemplate?: string;
  headers?: Record<string, string>;
}

interface ApiGroup {
  name: string;
  icon?: string;
  children?: ApiGroup[];
  apis?: ApiCallConfig[];
}

interface AddinConfiguration {
  version: string;
  title: string;
  configUrl?: string;  // For nested config loading
  groups: ApiGroup[];
}

async function loadConfiguration(configUrl: string): Promise<AddinConfiguration> {
  const response = await fetch(configUrl, {
    method: "GET",
    headers: {
      "Accept": "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(
      `Failed to load configuration from ${configUrl}: ${response.status} ${response.statusText}`
    );
  }

  return await response.json() as AddinConfiguration;
}
```

### Example Configuration JSON

```json
{
  "version": "1.0",
  "title": "AI Writing Tools",
  "groups": [
    {
      "name": "Text Analysis",
      "apis": [
        {
          "name": "Summarize",
          "url": "https://api.example.com/v1/summarize",
          "method": "POST",
          "inputType": "both",
          "promptTemplate": "Summarize the following text:\n\n{{text}}"
        },
        {
          "name": "Sentiment Analysis",
          "url": "https://api.example.com/v1/sentiment",
          "method": "POST",
          "inputType": "selected",
          "promptTemplate": "Analyze the sentiment of:\n\n{{text}}"
        }
      ]
    },
    {
      "name": "Content Generation",
      "children": [
        {
          "name": "Creative",
          "apis": [
            {
              "name": "Expand Text",
              "url": "https://api.example.com/v1/expand",
              "method": "POST",
              "inputType": "selected",
              "promptTemplate": "Expand and elaborate on:\n\n{{text}}"
            }
          ]
        }
      ]
    }
  ]
}
```

### CORS Considerations for Config Loading

The remote server hosting the configuration JSON **must** include appropriate CORS headers:

```
Access-Control-Allow-Origin: *
Access-Control-Allow-Methods: GET
Access-Control-Allow-Headers: Accept, Content-Type
```

Alternatively, host the configuration on the same domain as the add-in's web assets.

### Configuration Caching Strategy

```typescript
const CONFIG_CACHE_KEY = "addin_config";
const CONFIG_CACHE_TTL = 3600000; // 1 hour in ms

async function getConfiguration(configUrl: string): Promise<AddinConfiguration> {
  const cached = localStorage.getItem(CONFIG_CACHE_KEY);
  if (cached) {
    const { data, timestamp } = JSON.parse(cached);
    if (Date.now() - timestamp < CONFIG_CACHE_TTL) {
      return data as AddinConfiguration;
    }
  }

  const config = await loadConfiguration(configUrl);
  localStorage.setItem(
    CONFIG_CACHE_KEY,
    JSON.stringify({ data: config, timestamp: Date.now() })
  );
  return config;
}
```

---

## 7. Local Storage and Persistence

### Available Storage Mechanisms

| Mechanism | Capacity | Scope | Persistence | Best For |
|-----------|----------|-------|-------------|----------|
| `localStorage` | ~5 MB | Per origin + partition | Across sessions | Config cache, settings, small history |
| `sessionStorage` | ~5 MB | Per tab/session | Current session only | Temporary state |
| `IndexedDB` | 50%+ of disk | Per origin | Across sessions | Large history, structured data |
| `Office.Settings` | ~2 MB | Per document + add-in | Saved in document | Document-specific settings |
| Custom XML Parts | Varies | Per document | Saved in document | Structured data in document |
| `Office.context.document.settings` (Word API) | ~2 MB | Per document | Saved in document | Word-specific settings |

### Prompt History with IndexedDB

For storing prompt history (which can grow large), IndexedDB is the recommended approach:

```typescript
interface HistoryEntry {
  id?: number;
  timestamp: number;
  apiName: string;
  apiUrl: string;
  prompt: string;
  inputType: "selected" | "full";
  inputText: string;
  response?: string;
}

class HistoryStore {
  private dbName = "WordAddinHistory";
  private storeName = "prompts";
  private db: IDBDatabase | null = null;

  async open(): Promise<void> {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, 1);

      request.onupgradeneeded = (event) => {
        const db = (event.target as IDBOpenDBRequest).result;
        if (!db.objectStoreNames.contains(this.storeName)) {
          const store = db.createObjectStore(this.storeName, {
            keyPath: "id",
            autoIncrement: true,
          });
          store.createIndex("timestamp", "timestamp", { unique: false });
          store.createIndex("apiName", "apiName", { unique: false });
        }
      };

      request.onsuccess = (event) => {
        this.db = (event.target as IDBOpenDBRequest).result;
        resolve();
      };

      request.onerror = () => reject(request.error);
    });
  }

  async addEntry(entry: HistoryEntry): Promise<number> {
    return new Promise((resolve, reject) => {
      if (!this.db) {
        reject(new Error("Database not initialized"));
        return;
      }
      const tx = this.db.transaction(this.storeName, "readwrite");
      const store = tx.objectStore(this.storeName);
      const request = store.add(entry);
      request.onsuccess = () => resolve(request.result as number);
      request.onerror = () => reject(request.error);
    });
  }

  async getRecentEntries(limit: number = 50): Promise<HistoryEntry[]> {
    return new Promise((resolve, reject) => {
      if (!this.db) {
        reject(new Error("Database not initialized"));
        return;
      }
      const tx = this.db.transaction(this.storeName, "readonly");
      const store = tx.objectStore(this.storeName);
      const index = store.index("timestamp");
      const request = index.openCursor(null, "prev");
      const results: HistoryEntry[] = [];

      request.onsuccess = (event) => {
        const cursor = (event.target as IDBRequest<IDBCursorWithValue>).result;
        if (cursor && results.length < limit) {
          results.push(cursor.value);
          cursor.continue();
        } else {
          resolve(results);
        }
      };

      request.onerror = () => reject(request.error);
    });
  }
}
```

### Storage Partitioning

When running in Office on the web, use `Office.context.partitionKey` for localStorage isolation:

```typescript
function setPartitionedItem(key: string, value: string): void {
  const partitionKey = Office.context.partitionKey;
  const fullKey = partitionKey ? `${partitionKey}_${key}` : key;
  localStorage.setItem(fullKey, value);
}

function getPartitionedItem(key: string): string | null {
  const partitionKey = Office.context.partitionKey;
  const fullKey = partitionKey ? `${partitionKey}_${key}` : key;
  return localStorage.getItem(fullKey);
}
```

### Known Storage Issues

1. **localStorage not shared between task pane and dialog** -- If the add-in opens a dialog via `Office.context.ui.displayDialogAsync()`, the dialog may have a separate localStorage instance.
2. **Browser cache clearing** -- If the user clears browser data, localStorage and IndexedDB data is lost.
3. **Sideloaded add-ins on web** -- Manifest is stored in browser's local storage; clearing cache removes the sideloaded add-in.
4. **Storage partitioning** -- On Office for the web, each host+add-in combination gets its own partition.

---

## 8. Sideloading and Debugging

### Development Workflow

#### Using Yeoman-Generated Project (Recommended)

```bash
# 1. Create the project
yo office --projectType react --name "Word API Sidebar" --host word --ts true

# 2. Navigate to project
cd "Word API Sidebar"

# 3. Start dev server + sideload into Word desktop
npm start

# 4. Stop and uninstall
npm stop
```

`npm start` automatically:
- Starts a local webpack dev server on HTTPS (port 3000)
- Generates and trusts a self-signed SSL certificate (first run)
- Sideloads the manifest into Word desktop
- Opens Word with the add-in loaded

#### Sideloading on macOS (Manual)

```bash
# Copy manifest to Word's add-in directory
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```

#### Sideloading on Office for the Web (Manual)

1. Open Word on the web (word.cloud.microsoft)
2. Open a document
3. Select **Home** > **Add-ins** > **More Add-ins**
4. Select **Upload My Add-in**
5. Browse to and upload the manifest file

#### Debugging in VS Code

1. Install the **Office Add-in Debugger** extension
2. Use the launch configuration:

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "type": "office-addin",
      "request": "attach",
      "name": "Attach to Office Add-in",
      "port": 9229,
      "trace": "verbose",
      "url": "https://localhost:3000/taskpane.html",
      "webRoot": "${workspaceFolder}/src"
    }
  ]
}
```

3. Use F12 Developer Tools on Windows (Edge DevTools) to inspect the webview directly.

#### Hot Reload

The Yeoman-generated webpack config supports hot module replacement (HMR). Changes to source files trigger automatic reload of the task pane content during development.

---

## 9. Deployment Options

### Option 1: Centralized Deployment (Recommended for Organizations)

Deploy via the **Microsoft 365 Admin Center** Integrated Apps portal.

| Feature | Details |
|---------|---------|
| **Who deploys** | Microsoft 365 admin |
| **Target audience** | Specific users, groups, or entire organization |
| **Platforms** | Windows, Mac, Office on the web |
| **Auto-update** | Yes, when published via Marketplace |
| **Deployment modes** | Fixed, Available, or Optional |
| **Timing** | Up to 24 hours for new deployment; up to 72 hours for updates |
| **Requirements** | Microsoft 365 Business/Enterprise license, Exchange Online |

**Steps:**
1. Admin signs in to `admin.microsoft.com`
2. Navigate to **Settings** > **Integrated Apps**
3. Select **Deploy Add-in**
4. Upload the manifest file or point to AppSource listing
5. Assign to users/groups

### Option 2: Microsoft Marketplace (AppSource)

For public distribution. Requires:
- Compliance with [Commercial Marketplace Certification Policies](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies)
- Partner Center account
- Manifest validation
- Review process (can take days/weeks)

### Option 3: SharePoint App Catalog

For organizations using SharePoint. Limited features -- does not support `VersionOverrides` (no ribbon commands).

### Option 4: Network Share (Windows Only)

For internal testing/distribution on Windows:
1. Share a network folder containing the manifest
2. Users add the shared folder URL as a Trusted Catalog in Office Trust Center
3. Add-in appears in the add-in catalog

### Option 5: Direct Sideloading

For development and testing only. Not suitable for production distribution.

### Deployment Recommendation

For the described use case (organizational tool):
- **Development/Testing:** Sideloading via `npm start`
- **Staging/UAT:** Network share or manual sideloading
- **Production:** Centralized Deployment via Microsoft 365 Admin Center

---

## 10. Known Limitations

### CORS

- The task pane webview supports **full CORS** for `fetch()` and `XMLHttpRequest` calls.
- The target API server must include appropriate CORS headers (`Access-Control-Allow-Origin`).
- **Custom functions** (not relevant for task pane) have limited CORS unless using a shared runtime.
- All communication must use **HTTPS** -- HTTP is blocked in production and generates warnings in development.

### Sandbox Restrictions

- No access to the local file system (except through File Picker dialogs)
- No ActiveX components
- Cannot navigate the main Office window
- Limited to standard web APIs available in the webview engine
- No Node.js APIs (it is a browser environment, not Electron)

### API Limits

- Office.js batch operations (`context.sync()`) have practical limits on the number of queued operations
- Very large documents may cause performance issues when loading full body text
- `getSelection()` with multi-row table selections may return corrupted data
- `insertText()` with `Replace` on full document selection (Ctrl+A) may throw exceptions

### Task Pane Limitations

- Cannot programmatically set task pane width via Office.js
- Default width ~320px; users can resize manually
- The personality menu (12x32px Windows, 34x32px Mac) occupies the top-right corner
- Task pane is destroyed and recreated when the add-in is closed and reopened (state must be persisted externally)

### Platform Differences

| Feature | Windows | macOS | Web |
|---------|---------|-------|-----|
| WebView | Edge Chromium | WebKit | Browser |
| localStorage | Available | Available | Partitioned |
| IndexedDB | Available | Available | Available |
| Dev Tools | F12 (Edge) | Safari Web Inspector | Browser DevTools |
| Sideloading | npm start / Network Share | Manual copy | Upload via UI |

---

## Comparison Matrix -- Scaffolding Tools

| Criteria | Yeoman (`yo office`) | Teams Toolkit | Manual Setup |
|----------|---------------------|---------------|-------------|
| **React + TypeScript** | Full support | Supported | Full control |
| **Word add-in support** | Production | Preview (Dev Kit) | Production |
| **XML manifest** | Yes | No (JSON only) | Yes |
| **JSON manifest** | Yes (preview) | Yes | Yes |
| **Built-in HTTPS** | Yes (auto-cert) | Yes | Manual |
| **Sideloading** | Automatic | Automatic | Manual |
| **Hot reload** | Yes (webpack HMR) | Yes | Depends on setup |
| **Fluent UI integration** | Template available | Template available | Manual |
| **Learning curve** | Low | Low-Medium | High |
| **Flexibility** | Medium | Low | High |
| **Community/Docs** | Extensive | Growing | N/A |
| **Maintenance** | Microsoft-maintained | Microsoft-maintained | Self-maintained |
| **Recommendation** | **Best choice** | Alternative | Advanced users only |

---

## Recommended Approach

### Technology Stack

| Layer | Technology | Rationale |
|-------|-----------|-----------|
| **Scaffolding** | Yeoman (`yo office`) with React + TypeScript | Most mature, best docs, automatic setup |
| **Manifest** | XML (add-in only) | Production-ready for Word; JSON is still preview |
| **UI Framework** | Fluent UI React v9 | Official Microsoft UI framework; Office-native look |
| **State Management** | React Context + hooks | Lightweight; sufficient for task pane scope |
| **Storage (history)** | IndexedDB via `idb` wrapper library | Large capacity; structured data; async API |
| **Storage (settings)** | localStorage (with partition key) | Simple key-value for config cache and preferences |
| **HTTP Client** | Native `fetch()` API | Available in all supported webviews; no extra deps |
| **Build Tool** | Webpack (Yeoman default) | Pre-configured; proven with Office.js |

### Project Structure

```
word-api-sidebar/
  manifest.xml                    # Office Add-in manifest
  package.json
  tsconfig.json
  webpack.config.js
  src/
    taskpane/
      taskpane.html               # Task pane entry point
      taskpane.css                 # Global styles
      index.tsx                   # React entry point + Office.onReady
      App.tsx                     # Main app component
      components/
        ApiGroupTree.tsx          # Tree view of API groups
        ApiButton.tsx             # Configurable API action button
        PromptEditor.tsx          # Prompt text area with template support
        HistoryPanel.tsx          # Prompt/response history view
        ConfigLoader.tsx          # Configuration URL input/loader
        StatusBar.tsx             # Status/error messages
      services/
        officeApi.ts              # Word document interaction (get text, insert text)
        configService.ts          # Remote config fetching and caching
        apiService.ts             # Execute configured API calls
        historyService.ts         # IndexedDB history management
      types/
        config.ts                 # Configuration type definitions
        history.ts                # History entry types
      hooks/
        useConfig.ts              # Configuration state hook
        useHistory.ts             # History state hook
        useDocumentText.ts        # Document text extraction hook
    commands/
      commands.html               # Commands entry point
      commands.ts                 # Ribbon button command handlers
  assets/
    icon-16.png                   # Add-in icons
    icon-32.png
    icon-80.png
```

### Implementation Steps

1. **Scaffold project** with `yo office` (React + TypeScript + Word)
2. **Install dependencies:** `@fluentui/react-components`, `idb` (IndexedDB wrapper)
3. **Define TypeScript interfaces** for configuration schema and history entries
4. **Implement `officeApi.ts`** with text extraction and insertion functions
5. **Implement `configService.ts`** for remote JSON config loading with caching
6. **Implement `historyService.ts`** for IndexedDB-based prompt history
7. **Build UI components** using Fluent UI React v9
8. **Configure manifest** with ribbon button and task pane source
9. **Test locally** via `npm start` (sideloading)
10. **Deploy** via Centralized Deployment or AppSource

---

## Key Code Patterns

### Pattern 1: Complete Office.js Document Service

```typescript
// src/taskpane/services/officeApi.ts

export class OfficeApiService {
  /**
   * Get the currently selected text in the Word document.
   * Returns empty string if no text is selected (cursor only).
   */
  async getSelectedText(): Promise<string> {
    return await Word.run(async (context: Word.RequestContext) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text;
    });
  }

  /**
   * Get the full text content of the document body.
   */
  async getFullDocumentText(): Promise<string> {
    return await Word.run(async (context: Word.RequestContext) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      return body.text;
    });
  }

  /**
   * Replace the current selection with the given text.
   */
  async replaceSelection(text: string): Promise<void> {
    await Word.run(async (context: Word.RequestContext) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  }

  /**
   * Insert text at the end of the document.
   */
  async appendText(text: string): Promise<void> {
    await Word.run(async (context: Word.RequestContext) => {
      context.document.body.insertText(text, Word.InsertLocation.end);
      await context.sync();
    });
  }

  /**
   * Determine what type of input is available.
   */
  async getAvailableInput(): Promise<{
    hasSelection: boolean;
    selectedText: string;
    fullText: string;
  }> {
    return await Word.run(async (context: Word.RequestContext) => {
      const selection = context.document.getSelection();
      const body = context.document.body;
      selection.load("text");
      body.load("text");
      await context.sync();
      return {
        hasSelection: selection.text.trim().length > 0,
        selectedText: selection.text,
        fullText: body.text,
      };
    });
  }
}
```

### Pattern 2: Configuration Service with Caching

```typescript
// src/taskpane/services/configService.ts

import { AddinConfiguration } from "../types/config";

const CONFIG_CACHE_KEY = "addin_config_cache";
const CONFIG_URL_KEY = "addin_config_url";
const CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour

export class ConfigService {
  /**
   * Save the config URL to localStorage for persistence.
   */
  saveConfigUrl(url: string): void {
    const partitionKey = Office.context.partitionKey || "";
    localStorage.setItem(`${partitionKey}${CONFIG_URL_KEY}`, url);
  }

  /**
   * Retrieve the previously saved config URL.
   */
  getConfigUrl(): string | null {
    const partitionKey = Office.context.partitionKey || "";
    return localStorage.getItem(`${partitionKey}${CONFIG_URL_KEY}`);
  }

  /**
   * Load configuration from URL with caching.
   */
  async loadConfig(url: string): Promise<AddinConfiguration> {
    const partitionKey = Office.context.partitionKey || "";
    const cacheKey = `${partitionKey}${CONFIG_CACHE_KEY}`;

    // Check cache
    const cached = localStorage.getItem(cacheKey);
    if (cached) {
      try {
        const { data, timestamp, sourceUrl } = JSON.parse(cached);
        if (sourceUrl === url && Date.now() - timestamp < CACHE_TTL_MS) {
          return data as AddinConfiguration;
        }
      } catch {
        // Invalid cache, proceed to fetch
      }
    }

    // Fetch fresh config
    const response = await fetch(url, {
      method: "GET",
      headers: { Accept: "application/json" },
    });

    if (!response.ok) {
      throw new Error(
        `Configuration load failed: ${response.status} ${response.statusText}`
      );
    }

    const config = (await response.json()) as AddinConfiguration;

    // Cache the result
    localStorage.setItem(
      cacheKey,
      JSON.stringify({
        data: config,
        timestamp: Date.now(),
        sourceUrl: url,
      })
    );

    this.saveConfigUrl(url);
    return config;
  }

  /**
   * Force reload configuration, bypassing cache.
   */
  async reloadConfig(url: string): Promise<AddinConfiguration> {
    const partitionKey = Office.context.partitionKey || "";
    localStorage.removeItem(`${partitionKey}${CONFIG_CACHE_KEY}`);
    return this.loadConfig(url);
  }
}
```

### Pattern 3: Dynamic Button Rendering Based on API Config

```tsx
// src/taskpane/components/ApiButton.tsx

import React, { useState } from "react";
import { Button, Spinner, Textarea, tokens } from "@fluentui/react-components";
import { ApiCallConfig } from "../types/config";
import { OfficeApiService } from "../services/officeApi";

interface ApiButtonProps {
  config: ApiCallConfig;
  officeApi: OfficeApiService;
  onHistoryAdd: (entry: any) => void;
}

export const ApiButton: React.FC<ApiButtonProps> = ({
  config,
  officeApi,
  onHistoryAdd,
}) => {
  const [loading, setLoading] = useState(false);
  const [prompt, setPrompt] = useState(config.promptTemplate || "");
  const [result, setResult] = useState<string | null>(null);

  const executeApi = async (useSelection: boolean) => {
    setLoading(true);
    try {
      const input = await officeApi.getAvailableInput();
      const text = useSelection ? input.selectedText : input.fullText;

      // Replace template placeholder with actual text
      const finalPrompt = prompt.replace("{{text}}", text);

      const response = await fetch(config.url, {
        method: config.method,
        headers: {
          "Content-Type": "application/json",
          ...(config.headers || {}),
        },
        body: JSON.stringify({ prompt: finalPrompt, text }),
      });

      if (!response.ok) {
        throw new Error(`API call failed: ${response.status}`);
      }

      const data = await response.json();
      setResult(data.result || JSON.stringify(data));

      onHistoryAdd({
        timestamp: Date.now(),
        apiName: config.name,
        apiUrl: config.url,
        prompt: finalPrompt,
        inputType: useSelection ? "selected" : "full",
        inputText: text.substring(0, 500),
        response: data.result || JSON.stringify(data),
      });
    } catch (error) {
      setResult(`Error: ${(error as Error).message}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ marginBottom: tokens.spacingVerticalM }}>
      {config.promptTemplate && (
        <Textarea
          value={prompt}
          onChange={(_, data) => setPrompt(data.value)}
          resize="vertical"
          style={{ width: "100%", marginBottom: tokens.spacingVerticalS }}
        />
      )}

      <div style={{ display: "flex", gap: tokens.spacingHorizontalS }}>
        {(config.inputType === "selected" || config.inputType === "both") && (
          <Button
            appearance="primary"
            onClick={() => executeApi(true)}
            disabled={loading}
            style={{ flex: 1 }}
          >
            {loading ? <Spinner size="tiny" /> : `${config.name} (Selection)`}
          </Button>
        )}

        {(config.inputType === "full" || config.inputType === "both") && (
          <Button
            appearance="secondary"
            onClick={() => executeApi(false)}
            disabled={loading}
            style={{ flex: 1 }}
          >
            {loading ? <Spinner size="tiny" /> : `${config.name} (Full Doc)`}
          </Button>
        )}
      </div>

      {result && (
        <div
          style={{
            marginTop: tokens.spacingVerticalS,
            padding: tokens.spacingVerticalS,
            backgroundColor: tokens.colorNeutralBackground3,
            borderRadius: tokens.borderRadiusMedium,
            fontSize: tokens.fontSizeBase200,
            whiteSpace: "pre-wrap",
            maxHeight: "200px",
            overflow: "auto",
          }}
        >
          {result}
        </div>
      )}
    </div>
  );
};
```

### Pattern 4: React Entry Point with Office.onReady

```tsx
// src/taskpane/index.tsx

import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./App";

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const container = document.getElementById("root");
    if (!container) {
      throw new Error("Root element not found");
    }
    const root = createRoot(container);
    root.render(
      <FluentProvider theme={webLightTheme}>
        <App />
      </FluentProvider>
    );
  }
});
```

### Pattern 5: Main App Component with Tab Navigation

```tsx
// src/taskpane/App.tsx

import React, { useState, useEffect } from "react";
import {
  Tab,
  TabList,
  SelectTabEvent,
  SelectTabData,
  Input,
  Button,
  Spinner,
  Title3,
  tokens,
} from "@fluentui/react-components";
import { ConfigService } from "./services/configService";
import { OfficeApiService } from "./services/officeApi";
import { HistoryStore } from "./services/historyService";
import { ApiGroupTree } from "./components/ApiGroupTree";
import { HistoryPanel } from "./components/HistoryPanel";
import { AddinConfiguration } from "./types/config";

const configService = new ConfigService();
const officeApi = new OfficeApiService();
const historyStore = new HistoryStore();

const App: React.FC = () => {
  const [selectedTab, setSelectedTab] = useState<string>("apis");
  const [configUrl, setConfigUrl] = useState<string>(
    configService.getConfigUrl() || ""
  );
  const [config, setConfig] = useState<AddinConfiguration | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    historyStore.open();
    // Auto-load saved config URL
    const savedUrl = configService.getConfigUrl();
    if (savedUrl) {
      loadConfig(savedUrl);
    }
  }, []);

  const loadConfig = async (url: string) => {
    setLoading(true);
    setError(null);
    try {
      const cfg = await configService.loadConfig(url);
      setConfig(cfg);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const onTabSelect = (_: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value as string);
  };

  return (
    <div style={{ padding: tokens.spacingVerticalM, height: "100vh", display: "flex", flexDirection: "column" }}>
      <Title3 style={{ marginBottom: tokens.spacingVerticalM }}>
        {config?.title || "API Sidebar"}
      </Title3>

      {/* Config URL input */}
      <div style={{ display: "flex", gap: tokens.spacingHorizontalXS, marginBottom: tokens.spacingVerticalM }}>
        <Input
          value={configUrl}
          onChange={(_, data) => setConfigUrl(data.value)}
          placeholder="Configuration URL..."
          style={{ flex: 1 }}
        />
        <Button
          appearance="primary"
          onClick={() => loadConfig(configUrl)}
          disabled={loading || !configUrl}
        >
          {loading ? <Spinner size="tiny" /> : "Load"}
        </Button>
      </div>

      {error && (
        <div style={{ color: tokens.colorPaletteRedForeground1, marginBottom: tokens.spacingVerticalS }}>
          {error}
        </div>
      )}

      {/* Tab navigation */}
      <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
        <Tab value="apis">APIs</Tab>
        <Tab value="history">History</Tab>
      </TabList>

      {/* Tab content */}
      <div style={{ flex: 1, overflow: "auto", marginTop: tokens.spacingVerticalS }}>
        {selectedTab === "apis" && config && (
          <ApiGroupTree
            groups={config.groups}
            officeApi={officeApi}
            historyStore={historyStore}
          />
        )}
        {selectedTab === "history" && (
          <HistoryPanel historyStore={historyStore} />
        )}
      </div>
    </div>
  );
};

export default App;
```

---

## References

### Official Microsoft Documentation

1. [Office Add-ins Platform Overview](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins) -- Core architecture and concepts
2. [Task Panes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/task-pane-add-ins) -- Task pane design guidelines and dimensions
3. [Build Your First Word Task Pane Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart-yo) -- Yeoman quickstart for Word
4. [Word Add-in Tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial) -- Step-by-step Word API tutorial
5. [Office Add-ins Manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests) -- Manifest format reference
6. [Compare XML vs Unified JSON Manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/json-manifest-overview) -- Manifest format comparison
7. [XML Manifest Overview](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview) -- XML manifest reference
8. [Unified Manifest Overview](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview) -- JSON manifest reference
9. [Create Add-in Commands (XML)](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/create-addin-commands) -- Ribbon button configuration
10. [Create Add-in Commands (Unified Manifest)](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/create-addin-commands-unified-manifest) -- JSON ribbon configuration
11. [Persist Add-in State and Settings](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings) -- Storage mechanisms
12. [Fluent UI React in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/fluent-react-quickstart) -- Fluent UI v9 quickstart
13. [Design the UI of Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-design) -- UI design best practices
14. [UX Design Pattern Templates](https://learn.microsoft.com/en-us/office/dev/add-ins/design/ux-design-pattern-templates) -- UX patterns
15. [Sideload Office Add-ins for Testing](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) -- Sideloading guide
16. [Deploy and Publish Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish) -- Deployment options
17. [Privacy and Security for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/privacy-and-security) -- Security model
18. [Addressing Same-Origin Policy Limitations](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/addressing-same-origin-policy-limitations) -- CORS workarounds
19. [Read and Write Data to Active Selection](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet) -- Selection API
20. [Runtimes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes) -- Runtime model details
21. [Set Up Development Environment](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/set-up-your-dev-environment) -- Dev environment setup
22. [Yeoman Generator Overview](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview) -- Yeoman generator docs

### API References

23. [Word.Document class](https://learn.microsoft.com/en-us/javascript/api/word/word.document?view=word-js-preview) -- Document API
24. [Word.Body class](https://learn.microsoft.com/en-us/javascript/api/word/word.body?view=word-js-preview) -- Body text API
25. [Word.GetTextOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.gettextoptions?view=word-js-preview) -- getText options
26. [Office.Document interface](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=common-js-preview) -- Common API document

### GitHub Repositories

27. [OfficeDev/generator-office](https://github.com/OfficeDev/generator-office) -- Yeoman generator source
28. [OfficeDev/Office-Addin-TaskPane-JS](https://github.com/OfficeDev/Office-Addin-TaskPane-JS) -- Task pane template (manifest example)
29. [microsoft/fluentui](https://github.com/microsoft/fluentui) -- Fluent UI React source
30. [OfficeDev/Office-Add-ins-Fluent-React-version-8](https://github.com/OfficeDev/Office-Add-ins-Fluent-React-version-8) -- Fluent v8 samples (legacy)
31. [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) -- Official samples collection

### Deployment

32. [Centralized Deployment FAQ](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-faq) -- Centralized deployment FAQ
33. [Deploy Office Add-ins in Admin Center](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins) -- Admin center deployment
34. [Requirements for Centralized Deployment](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins) -- Requirements
35. [Best Practices for Designing Word, Excel, and PowerPoint Add-ins](https://devblogs.microsoft.com/microsoft365dev/best-practices-for-designing-word-excel-and-powerpoint-add-ins/) -- Microsoft 365 Developer Blog

### Community / Issues

36. [Office.js Developer Concerns (Issue #6513)](https://github.com/OfficeDev/office-js/issues/6513) -- Developer open letter on platform stability
37. [localStorage not syncing between add-in windows (Issue #600)](https://github.com/OfficeDev/office-js/issues/600) -- Storage sharing issue
38. [Task Pane Width Ignored by Word (Issue #1629)](https://github.com/OfficeDev/office-js/issues/1629) -- Width configuration limitation
