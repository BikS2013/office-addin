# Word Add-in API Sidebar

A generic Microsoft Word add-in implemented as a sidebar (task pane) that dynamically loads API configurations from a remote URL, renders API groups as a hierarchical tree, and allows users to execute API calls with optional prompts combined with selected text or full document text.

## Quick Start

```bash
npm install          # Install dependencies
npm start            # Launch dev server + sideload in Word
npm test             # Run test suite (83 tests)
npm run build        # Production build to dist/
npm run lint         # TypeScript type checking
```

## Architecture

- **Runtime**: Office.js Word API (WordApi 1.3+)
- **Framework**: React 18 + TypeScript (strict mode)
- **UI Library**: Fluent UI React v9
- **State**: React Context + useReducer
- **Storage**: IndexedDB (history) + localStorage (config cache)
- **Build**: Webpack 5 with HTTPS dev server
- **Tests**: Vitest

## Project Structure

```
src/taskpane/
  types/        - TypeScript type definitions
  services/     - ConfigService, OfficeApiService, ApiExecutionService, HistoryService
  context/      - React Context provider + app reducer
  hooks/        - useConfig, useHistory, useDocumentText, useApiExecution
  components/   - All React UI components
  App.tsx       - Root component
  index.tsx     - Entry point (Office.onReady)
```

## Configuration

No fallback values for configuration settings. Missing config raises exceptions.

## Tools

<devServer>
    <objective>
        Launch the development server with HTTPS and sideload the add-in in Word
    </objective>
    <command>
        npm start
    </command>
    <info>
        Starts webpack-dev-server on https://localhost:3000 with HTTPS certificates.
        The add-in is served from taskpane.html.
        Hot module replacement is enabled for development.
    </info>
</devServer>

<buildProject>
    <objective>
        Build the project for production deployment
    </objective>
    <command>
        npm run build
    </command>
    <info>
        Compiles TypeScript and bundles with webpack in production mode.
        Output goes to the dist/ directory with content hashing.
        Assets are copied from the assets/ directory.
    </info>
</buildProject>

<runTests>
    <objective>
        Run the full test suite using Vitest
    </objective>
    <command>
        npm test
    </command>
    <info>
        Runs all tests in the test_scripts/ directory using Vitest.
        Test files:
        - configService.test.ts (17 tests) - Config loading, validation, caching
        - apiExecutionService.test.ts (13 tests) - Request building, execution, error handling
        - historyService.test.ts (14 tests) - IndexedDB CRUD, filtering, auto-prune
        - appReducer.test.ts (39 tests) - All 17 action types, immutability checks
        Total: 83 tests
    </info>
</runTests>

<typeCheck>
    <objective>
        Run TypeScript type checking without emitting files
    </objective>
    <command>
        npm run lint
    </command>
    <info>
        Runs tsc --noEmit to verify all TypeScript types resolve correctly.
        Uses strict mode with all recommended checks enabled.
    </info>
</typeCheck>
