---
sessionId: session-261117-091644-2mkw
---

# Requirements

### Overview & Goals
Update `README.md` so it accurately reflects all current features, workflow steps, project files, dependencies, OAuth permissions, and recent updates of **Swim Entries Manager**, bringing it into parity with `docs/USER_GUIDE.md`, `appsscript.json`, and the current repository structure.

### Scope
- **In Scope:**
  - Updating **Features** and **Quick Start → First Use** in `README.md` to match `docs/USER_GUIDE.md` (Workspace & Export Setup, Meet Setup with competition types, Google Form Entry Generator & Sync, Relay Tables, and smart export destination / exception reporting).
  - Updating the **Project Structure** tree in `README.md` to include all current `.js`, `.html`, and `docs/` files.
  - Updating **Dependencies** (bundled `TableApp_Local.js` instead of external library ID, plus Drive API v3 & Google Picker API) and **Permissions** (matching `appsscript.json`: removing the sensitive `spreadsheets` scope, adding `userinfo.profile`, and explaining `GOOGLE_PICKER_API_KEY` / `GOOGLE_CLOUD_PROJECT_NUMBER` for **Workspace & Export Setup**).
  - Updating **Testing** (mentioning `runAllDiagnosticTests()` in `tests.js`) and **Recent Updates** to describe the latest releases above `v2.1`.
- **Out of Scope:**
  - Modifying any `.js`, `.html`, `appsscript.json`, or `docs/` files.
  - Rewriting `CONTRIBUTING.md`, `LICENSE`, or support/donate links in `README.md`.

### Functional Requirements
- **FR-1 (Features & Quick Start Parity):** `README.md` shall document:
  - **Workspace & Export Setup**: One-click Google Picker authorization for the active spreadsheet under `drive.file`, per-spreadsheet My Drive / Shared Drive export destination selection, and in-place sidepanel **Refresh Status**.
  - **Meet Setup (STEP 1)**: Competition-type-aware setup (`School`, `College`, `Age Group`, `Open`, `Masters`) generating `Meet` (with `AgeBrackets` for School/College), `Events`, `Schools` or `Teams`, `Individual Events Template`, and optional `Team Officials` / `How to fill entries` sample sheet.
  - **Google Form Entry Generator (`Create / Sync Entry Form`)**: Public-facing form creation linked to `Form Entries`, configurable event slots (1–9), optional email collection & response editing, optional Date of Birth for School/College, optional write-in `"Other"` team for Age Group/Open/Masters, and in-place option syncing.
  - **Sheet Duplication & Relay Tables**: `Create Team Sheets` (STEP 2) with **View sheet** preview and `Create Relay Tables` (STEP 3, optional).
  - **SDIF & CSV Export**: Auto-detection of `Form Entries`, `Schools` vs `Teams` table awareness, duplicate-event warnings, and saving `.sd3`, `.csv`, and `_ExceptionReport.txt` to the configured Drive folder (with My Drive root fallback).
- **FR-2 (Accurate Project Structure Diagram):** The directory tree in `README.md` shall list every current root `.js` and `.html` file (`TableApp_Local.js`, `core-tables-ui.html`, `export-csv-ui.html`, `form-setup-ui.html`, `formBuilder.js`, `meet-setup-ui.html`, `relay-ui-full.html`, `resizeSwim.js`, `shared-dialog-styles.html`, `shared-drive-picker.html`, `sheets-from-template-ui.html`, `workspace-setup-ui.html`) and `docs/SWIM_ENTRIES_MANAGER_SPEC.md`.
- **FR-3 (Accurate Dependencies & Permissions):**
  - **Dependencies** shall reflect **Google Sheets API v4**, **Google Drive API v3**, **Google Picker API**, and the bundled **TableApp (`TableApp_Local.js`)** rather than the removed external library ID.
  - **Permissions** shall match `appsscript.json` exactly (`spreadsheets.currentonly`, `drive.file`, `userinfo.email`, `userinfo.profile`, `script.container.ui`, `script.locale`, `forms`), explicitly omitting the removed sensitive `spreadsheets` scope, and explain `GOOGLE_PICKER_API_KEY` and `GOOGLE_CLOUD_PROJECT_NUMBER` Script Properties for **Workspace & Export Setup**.
- **FR-4 (Recent Updates):** **Recent Updates** shall include entries above `v2.1` covering Workspace & Export Setup / non-sensitive `drive.file` authorization, Google Form Entry Generator & Sync, competition-type-aware Meet Setup (`School`/`College`/`Age Group`/`Open`/`Masters`), `Teams` table & duplicate-event exception reporting, and the updated in-app Help.

# Technical Design

### Current Implementation
- `README.md` (284 lines) still describes the legacy **Create Core Tables** workflow, omits **Workspace & Export Setup**, **Meet Setup**, **Google Form Entry Generator**, and **Relay Tables**, lists an outdated **Project Structure** missing 12 root files and `docs/SWIM_ENTRIES_MANAGER_SPEC.md`, references the old external TableApp Library ID in **Dependencies**, still lists the removed sensitive `spreadsheets` scope (and omits `userinfo.profile`) in **Permissions**, and stops **Recent Updates** at `v2.1`.

### Proposed Changes
All edits are confined to `README.md`:
1. **Features & Quick Start (`README.md` lines 11–83):**
   - Reorganize **Features** into clear subsections: **Workspace & Export Setup**, **Meet & Template Management** (`Meet Setup` + competition types + `How to fill entries`), **Sheet Duplication & Relay Tables**, **Google Form Self-Entry**, and **Data Export & Validation** (`.sd3`, `.csv`, `Form Entries` auto-detection, `Schools`/`Teams` lookup, duplicate-event warnings, destination folder persistence).
   - Update **Quick Start → First Use** to follow the current `docs/USER_GUIDE.md` workflow: Step 0 (**Workspace & Export Setup**), Step 1 (**Meet Setup**), Populate Tables (`Events`, `Schools`/`Teams`, `Team Officials`), Step 2 (**Create Team Entry Sheets** or **Entry Form**), Step 3 (**Create Relay Tables**, optional), Fill Swimmer Entries (noting `Date of Birth` vs `School Year` by competition type), and Step 4 (**Export for Meet Manager** / **Export to CSV**).
2. **Project Structure (`README.md` lines 84–104):**
   - Replace the outdated ASCII tree with the complete repository layout matching the root directory and `docs/` (`SWIM_ENTRIES_MANAGER_SPEC.md`, `USER_GUIDE.md`, `PRIVACY_POLICY.md`, `TERMS_OF_SERVICE.md`).
3. **Configuration, Dependencies, Permissions & Testing (`README.md` lines 106–208):**
   - Update **Dependencies** to list Google Sheets API v4, Google Drive API v3, Google Picker API, and bundled `TableApp_Local.js` (removing the stale external Library ID).
   - Update **Permissions** to match `appsscript.json` (`spreadsheets.currentonly`, `drive.file`, `userinfo.email`, `userinfo.profile`, `script.container.ui`, `script.locale`, `forms` — no sensitive `spreadsheets` scope) and document `GOOGLE_PICKER_API_KEY` and `GOOGLE_CLOUD_PROJECT_NUMBER` Script Properties for **Workspace & Export Setup**.
   - Update **Testing** to reference `runAllDiagnosticTests()` in `tests.js`.
4. **Recent Updates (`README.md` lines 221–245):**
   - Add **v2.3 — Workspace & Export Setup & Non-Sensitive Scopes** (Google Picker spreadsheet authorization under `drive.file`, removal of sensitive `spreadsheets` scope, per-spreadsheet Drive/Shared Drive export folder selection, sidepanel `Refresh Status`).
   - Add **v2.2 — Meet Setup, Competition Types & Google Form Entry Generator** (5 competition types with `School Year` / `AgeBrackets` and `Teams` vs `Schools`, `formBuilder.js` / `form-setup-ui.html` with `Form Entries` auto-detection & duplicate-column cleanup, optional `How to fill entries` sample sheet, duplicate-event warnings, and updated in-app Help).

# Testing

### Validation Approach
- Verify every file path listed in the **Project Structure** section of `README.md` exists on disk.
- Verify every OAuth scope listed in **Permissions** matches `appsscript.json` `oauthScopes` 1:1.
- Verify all Markdown links in `README.md` resolve to valid files/anchors in the repository.

# Delivery Steps

### ✓ Step 1: Update Features and Quick Start workflow in README.md
The **Features** and **Quick Start → First Use** sections of `README.md` match `docs/USER_GUIDE.md` and the current sidepanel/menu workflow.

- Update **Features** in `README.md` to cover **Workspace & Export Setup**, **Meet Setup** (5 competition types + `How to fill entries`), **Sheet Duplication & Relay Tables**, **Google Form Self-Entry** (`Create / Sync Entry Form`), and **Data Export & Validation** (`Form Entries` auto-detection, `Schools`/`Teams` lookup, duplicate-event warnings, Drive folder destination).
- Update **Quick Start → First Use** in `README.md` to walk through Steps 0–4 (`Workspace & Export Setup`, `Meet Setup`, populating `Events`/`Schools`/`Teams`, `Create Team Entry Sheets` or `Entry Form`, optional `Create Relay Tables`, swimmer entry requirements by competition type, and SDIF/CSV export).

### ✓ Step 2: Update Project Structure, Dependencies, Permissions, and Testing in README.md
The **Project Structure** diagram, **Dependencies**, **Permissions**, and **Testing** sections in `README.md` match the actual codebase and `appsscript.json`.

- Update the **Project Structure** tree in `README.md` to include all current root `.js`/`.html` files and `docs/SWIM_ENTRIES_MANAGER_SPEC.md`.
- Update **Dependencies** to reference bundled `TableApp_Local.js`, Google Sheets API v4, Google Drive API v3, and Google Picker API.
- Update **Permissions** to match `appsscript.json` (`spreadsheets.currentonly`, `drive.file`, `userinfo.email`, `userinfo.profile`, `script.container.ui`, `script.locale`, `forms`) and document `GOOGLE_PICKER_API_KEY` and `GOOGLE_CLOUD_PROJECT_NUMBER` for **Workspace & Export Setup**.
- Update **Testing** to include `runAllDiagnosticTests()` in `tests.js`.

### ✓ Step 3: Add Recent Updates and verify README.md consistency
**Recent Updates** in `README.md` documents recent releases above `v2.1`, and all paths, scopes, and Markdown links are verified.

- Add `v2.3` and `v2.2` entries to **Recent Updates** in `README.md` summarizing Workspace & Export Setup, non-sensitive `drive.file` authorization, Google Form Entry Generator & Sync, competition-type-aware Meet Setup, `Teams`/`Schools` exception reporting, and updated in-app Help.
- Verify that every file in **Project Structure**, every scope in **Permissions**, and every relative link in `README.md` matches the repository.