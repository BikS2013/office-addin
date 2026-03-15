# Configuration Guide: Word Add-in API Sidebar

**Date:** 2026-03-15
**Project:** Word Add-in API Sidebar

---

## Table of Contents

1. [Configuration Sources and Priority](#1-configuration-sources-and-priority)
2. [Remote API Configuration (JSON)](#2-remote-api-configuration-json)
3. [Local Storage Settings](#3-local-storage-settings)
4. [IndexedDB Storage](#4-indexeddb-storage)
5. [Build and Development Configuration](#5-build-and-development-configuration)
6. [XML Manifest Configuration](#6-xml-manifest-configuration)

---

## 1. Configuration Sources and Priority

The add-in uses multiple configuration sources. There is **no single config file or environment variable system**. Each source serves a distinct purpose:

| Priority | Source | Purpose | Persistence |
|----------|--------|---------|-------------|
| 1 (highest) | **Remote JSON URL** | Defines API groups, endpoints, and call configurations | Cached in localStorage (1h TTL) |
| 2 | **localStorage** | Caches the remote config + stores the last-used config URL | Browser storage, survives task pane close |
| 3 | **IndexedDB** | Stores execution history entries | Browser storage, persistent |
| 4 | **manifest.xml** | Declares add-in identity, permissions, ribbon button | Static, bundled with the add-in |
| 5 | **webpack.config.js** | Dev server port, HTTPS certificates, build output path | Static, development-time only |

**Important:** Per project policy, no fallback or default values are substituted for missing configuration. Missing required fields raise `ConfigurationError` exceptions.

---

## 2. Remote API Configuration (JSON)

This is the primary runtime configuration. The user provides a URL, and the add-in fetches, validates, and renders the configuration.

### How to Provide

The user enters an HTTPS URL in the **ConfigLoader** input field in the sidebar UI. The URL must:
- Use the `https://` protocol (exception: `http://localhost` and `http://127.0.0.1` are allowed for development)
- Return a valid JSON response with `Content-Type: application/json`
- Support CORS headers if hosted on a different origin than the add-in

Alternatively, the user can load a configuration from a **local JSON file** using the file picker in the ConfigLoader component.

### How to Obtain

The configuration JSON file must be authored and hosted by the organization or API provider. A sample configuration is available at `docs/reference/sample-config.json`.

### Recommended Management

Host the configuration JSON on a static file server, CDN, or API endpoint under your organization's control. Version the file alongside your API definitions.

### Schema Reference

#### Top-Level Fields

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `configVersion` | `string` | Yes | Version identifier for the configuration schema (e.g., `"1.0"`) |
| `name` | `string` | Yes | Display name for the configuration shown in the sidebar header |
| `description` | `string` | No | Optional description of the configuration suite |
| `groups` | `ApiGroup[]` | Yes | Array of API group definitions. Must contain at least one group |

#### ApiGroup Fields

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `id` | `string` | Yes | Unique identifier for the group |
| `name` | `string` | Yes | Display name rendered in the sidebar tree |
| `description` | `string` | No | Optional description shown as tooltip or subtitle |
| `groups` | `ApiGroup[]` | Conditional | Nested sub-groups. At least one of `groups` or `apis` must be present |
| `apis` | `ApiCallConfig[]` | Conditional | API call definitions. At least one of `groups` or `apis` must be present |

#### ApiCallConfig Fields

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `id` | `string` | Yes | Unique identifier for the API call |
| `name` | `string` | Yes | Display name shown on the action button(s) |
| `description` | `string` | No | Optional description shown as tooltip |
| `url` | `string` (HTTPS URL) | Yes | Target API endpoint URL. Must use HTTPS (localhost exempt) |
| `method` | `"GET"` or `"POST"` | Yes | HTTP method. Only GET and POST are supported |
| `inputMode` | `"selected"`, `"full"`, or `"both"` | Yes | Determines which action buttons are rendered and what document text is sent |
| `timeout` | `number` (ms) | Yes | Request timeout in milliseconds. **No default — must be explicitly set** |
| `headers` | `Record<string, string>` | No | Custom HTTP headers included in the request (e.g., `Authorization`) |
| `bodyTemplate` | `object` | No | JSON body template for POST requests. Supports `{{prompt}}`, `{{text}}`, `{{documentName}}` placeholders |
| `responseField` | `string` | No | Dot-notation path to extract text from the JSON response (e.g., `"choices.0.message.content"`). If absent, the raw JSON is displayed |
| `promptTemplate` | `string` | No | Default prompt pre-populated in the prompt editor. Supports `{{text}}` placeholder resolved at execution time |

#### inputMode Options

| Value | Buttons Rendered | Behavior |
|-------|-----------------|----------|
| `"selected"` | 1 button: "{name} (Selection)" | Sends only the currently selected text from the document |
| `"full"` | 1 button: "{name} (Full Doc)" | Sends the entire document text |
| `"both"` | 2 buttons: "(Selection)" + "(Full Doc)" | User chooses which text source to use |

#### Placeholder Tokens in bodyTemplate

| Placeholder | Replaced With |
|-------------|---------------|
| `{{prompt}}` | The user-edited prompt text from the prompt editor |
| `{{text}}` | The extracted document text (selected or full, depending on button clicked) |
| `{{documentName}}` | The active Word document's filename |

Placeholders are resolved recursively through all string values in the `bodyTemplate` object at execution time.

### Configuration Caching

- Successfully loaded configurations are cached in **localStorage** with a **1-hour TTL**
- The cache key is `word-addin-sidebar:addin-config-cache`
- The cache stores: the config object, the source URL, the cache timestamp, and the TTL
- If the cache is valid (same URL, TTL not expired), the cached config is served without a network request
- The user can force a cache bypass using the **Reload** button in the UI
- Cache write failures (e.g., localStorage full) are silently ignored — the add-in continues without caching

### CORS Requirements for API Providers

API endpoints called by the add-in must include the following CORS headers:

```
Access-Control-Allow-Origin: https://localhost:3000  (or * for development)
Access-Control-Allow-Methods: GET, POST, OPTIONS
Access-Control-Allow-Headers: Content-Type, Authorization, Accept
```

---

## 3. Local Storage Settings

localStorage is used for two purposes. Both use the key prefix `word-addin-sidebar:`.

| Key | Purpose | Value Format | Expiration |
|-----|---------|-------------|------------|
| `word-addin-sidebar:config-url` | Last-used configuration URL, auto-restored on task pane reopen | Plain string (URL) | Never expires |
| `word-addin-sidebar:addin-config-cache` | Cached configuration data | JSON: `{ config, url, cachedAt, ttlMs }` | TTL-based: 1 hour (3,600,000 ms) |

### How to Manage

- These values are managed automatically by the `ConfigService`
- To clear cached state, use the browser DevTools > Application > Local Storage, or clear the Office web cache
- No manual configuration of localStorage is required

### Known Limitation

The design specifies using `Office.context.partitionKey` as a key prefix for isolation in Office for the web. The current implementation uses a hardcoded prefix `"word-addin-sidebar"` instead. This is tracked as **P6** in the Issues file.

---

## 4. IndexedDB Storage

IndexedDB stores the execution history. This is managed entirely by the `HistoryService`.

| Setting | Value | Description |
|---------|-------|-------------|
| Database name | `word-addin-history` | The IndexedDB database name |
| Database version | `1` | Schema version |
| Object store | `entries` | Store name for history entries |
| Max entries | `1000` | Auto-prune oldest entries when exceeded |
| Max age | `90 days` | Auto-prune entries older than this |

### Indexes

| Index | Field | Purpose |
|-------|-------|---------|
| `timestamp` | `timestamp` | Chronological ordering and age-based pruning |
| `apiId` | `apiId` | Filtering history by API call |
| `documentName` | `documentName` | Filtering history by document |

### How to Manage

- History is managed automatically. No user configuration is needed
- The **Clear History** action in the History panel wipes all entries
- Individual entries can be deleted from the History panel
- To fully reset, delete the `word-addin-history` database from browser DevTools > Application > IndexedDB

---

## 5. Build and Development Configuration

### webpack.config.js

| Setting | Value | Description |
|---------|-------|-------------|
| Dev server port | `3000` | The HTTPS dev server listens on `https://localhost:3000` |
| HTTPS certificates | Auto-generated via `office-addin-dev-certs` | Self-signed certs for local HTTPS development |
| Output directory | `dist/` | Production build output |
| Content hashing | Enabled | Output filenames include content hashes for cache busting |
| HMR | Enabled | Hot Module Replacement for development |
| CORS header | `Access-Control-Allow-Origin: *` | Dev server allows all origins |

### How to Obtain HTTPS Certificates

HTTPS certificates are automatically generated by the `office-addin-dev-certs` package on first run of `npm start`. No manual action is required. If certificate issues occur, run:

```bash
npx office-addin-dev-certs install
```

### tsconfig.json

TypeScript is configured in **strict mode** with all recommended checks enabled. The `npm run lint` command runs `tsc --noEmit` to verify type correctness without emitting files.

---

## 6. XML Manifest Configuration

The `manifest.xml` file declares the add-in's identity to Microsoft Word. Key settings:

| Setting | Description | How to Change |
|---------|-------------|---------------|
| Add-in ID | Unique GUID identifying the add-in | Edit `<Id>` in manifest.xml. Generate a new GUID for each deployment |
| Display name | Name shown in the ribbon and store | Edit `<DisplayName>` |
| Task pane URL | URL of the sidebar HTML page | Edit `<SourceLocation>` (defaults to `https://localhost:3000/taskpane.html`) |
| Permissions | Document access level | `<Permissions>ReadWriteDocument</Permissions>` — required for text extraction and insertion |
| Ribbon button | Button in the Home tab that opens the task pane | Configured in the `<Action>` element |
| Icon sizes | 16x16, 32x32, 80x80 | Referenced in `<IconUrl>` and `<HighResolutionIconUrl>` elements |

### How to Obtain

The manifest is generated during project scaffolding and must be customized for each deployment environment. For production:
1. Generate a new GUID for `<Id>`
2. Update `<SourceLocation>` to point to the production hosting URL
3. Update `<SupportUrl>` with a valid support page

### Recommended Management

- Keep `manifest.xml` in version control
- Maintain separate manifests for development and production if the hosting URLs differ
- Validate the manifest using `npx office-addin-manifest validate -m manifest.xml` (note: requires `office-addin-manifest` package — see **P7** in Issues)
