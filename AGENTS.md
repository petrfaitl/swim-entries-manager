# AGENTS.md: Swim Entries Manager

## Overview

**Swim Entries Manager** is a Google Workspace Add-on for Google Sheets designed to streamline swim meet organisation. It provides structured entry templates, school/cluster registration management, automated validation, and export capabilities to SDIF v3 (.sd3 for Hy-Tek Meet Manager) and CSV format.

The functional specification lives in `docs/SPEC.md`. Read it, and the relevant section for the area you are changing, before making significant changes. Update it when behaviour changes.

---

## Tech Stack

### Core Platform & Runtime

- **Runtime Environment**: Google Apps Script (V8 runtime)
- **Target Application**: Google Sheets (Google Workspace Add-on)
- **Client UI**: Google Apps Script HTML Service / Card Service / Modal Dialogs / Custom Sidebars (`SpreadsheetApp.getUi()`)

### APIs & Advanced Services

- **Google Sheets API v4** (Advanced Service identifier: `Sheets`)
- **Google Drive API v3** (Advanced Service identifier: `Drive`)
- **Google Apps Script Built-in Services**: `SpreadsheetApp`, `DriveApp`, `HtmlService`, `Logger`, `Session`

### External Libraries & Integrations

- **TableApp**: Apps Script library for managing Google Sheets structured tables and data formatting.
- **SDIF v3.0 Specification**: Standard Data Interchange Format v3 with FINA organisation coding.
- **Hy-Tek Meet Manager**: Target meet management software consuming SDIF exports.

### Development Tooling

- **CLI / Deployment**: `@google/clasp` (Google Cloud Apps Script CLI)
- **Version Control**: Git / GitHub

---

## Architecture & Codebase Layout

```
SwimEntriesManager/
├── AGENTS.md                    # Agent instructions and project context (this document)
├── appsscript.json              # Project manifest (V8 runtime, scopes, triggers, services)
├── ui.js                        # UI orchestration (Add-on menus, sidebars, modal launchers)
├── tableConfig.js               # Table schemas, validation rules, column definitions, defaults
├── tableBuilder.js              # Table generation, named range bindings, TableApp management
├── duplicateTemplate.js         # Sheet duplication for schools/teams with collision detection
├── EntryDataProcessor.js        # Shared data parsing, validation, capitalisation, normalisation
├── SDIFCreator.js               # SDIF v3 (.sd3) record formatting and exception reporting
├── SDIFCreator_Dialog.html      # SDIF export modal dialog UI and client-side logic
├── exportToCSV.js               # CSV export generation and drive file persistence
├── export-csv-ui.html           # CSV export modal dialog UI
├── resizeSwim.js                # Pool distance & course conversion algorithms (18m, 33m, 50m LC to 25m SC)
├── utils.js                     # Shared utility functions (time parsing, formatting, helpers)
├── tests.js                     # Test fixtures, unit verification, and debug logging functions
└── docs/                        # Functional spec, user guides, privacy policy, and terms of service
    └── SPEC.md                  # Core functional specification
```

---

## Coding Standards & Conventions

### Language & Syntax

- **JavaScript Standard**: Modern ES6+ syntax supported by the Apps Script V8 engine (`const`, `let`, arrow functions, template literals, destructuring, spread syntax).
- **Prohibited**: Avoid legacy ES5 patterns (e.g., `var`, global scope pollution without namespaces).

### Naming Conventions

- **Public Functions & Methods**: `camelCase` (e.g., `createConfiguredTable`, `processSwimmerEntries`)
- **Private / Internal Helper Functions**: Suffix with a trailing underscore `_` (e.g., `createNamedRangeForColumn_`, `ensureStructuredTable_`, `parseTimeString_`)
- **Constants & Configuration Dictionaries**: `UPPER_SNAKE_CASE` (e.g., `DEFAULT_CLUSTERS`, `SDIF_SPEC`, `TABLE_DEFINITIONS`)
- **Variables & Function Parameters**: `camelCase` (e.g., `sheetName`, `tableConfig`)

### Function Design & Modularity

- Functions must remain small, modular, and single-purpose.
- Prefer pure helper functions for calculations and string/time parsing (e.g., in `utils.js` and `resizeSwim.js`).
- Encapsulate spreadsheet mutations and API calls within dedicated service functions.
- Keep UI presentation markup separated in HTML template files (`*.html`), communicating with server-side Apps Script via `google.script.run`.

### Table schemas, validation rules, column definitions, defaults

- Keep all schemas, validation rules, definitions and defaults in `tableConfig.js`

### Error Handling & Logging

- Wrap API calls, user inputs, and batch spreadsheet manipulations in `try...catch` blocks.
- Log operational errors and diagnostics to Stackdriver / Cloud Logging via `console.error` / `Logger.log`.
- Present actionable, user-friendly error messages via `SpreadsheetApp.getUi().alert()` or inline modal dialog error summaries; never expose raw stack traces to end users.
- Provide pre-export validation reports (exceptions/warnings) before generating downstream export files.

### Documentation & JSDoc

- Document all exported and core functions using standard JSDoc blocks (`@param`, `@returns`, `@throws`).
- Include inline explanations for non-trivial logic, particularly:
    - SDIF record byte positioning and column slicing
    - Pool course conversion math (e.g., Individual Medley distance formulas)
    - Named range dynamic coordinate resolutions
- Sync changes to our documentation in `README.md`, `USER_GUIDE.md` and `docs/SWIM_ENTRIES_MANAGER_SPEC.md`

---

## Testing & Quality Assurance

### Testing Strategy

- **Manual & Integration Testing in Google Sheets**:
    - Test deployments executed against live test spreadsheets using `clasp push` and `clasp open`.
    - Verification of table generation modes (`rebuild` vs `append`).
    - Validation of formula references and named range bindings across duplicated sheets.
- **Scripted Diagnostics**:
    - Use test functions located in `tests.js` for verifying parsing routines (e.g., `testFlexibleTimeParsing`).
    - Auto-detection testing for diverse entry formats (Method 1: event in cells vs Method 2: event in headers).

---

## Development Workflow

1. **Local Setup & Syncing**:
    - Use `@google/clasp` for pulling and pushing changes to Google Apps Script.
    - Keep `appsscript.json` synchronised with required OAuth scopes and advanced services.
2. **Specification**:
    - `docs/SWIM_ENTRIES_MANAGER_SPEC.md` describes what the system must do. Check it before implementing significant features or refactorings, and keep it aligned with the code.
3. **Version Control**:
    - Branch off `main` for feature work and bug fixes.
    - Commit clean, descriptive commits adhering to project conventions.
    - Push to a remote
