---
sessionId: session-261117-091644-2mkw
---

# Requirements

### Overview & Goals
Update `showHelpCard()` in `ui.js` so the **Help & Quick Start** sidepanel card matches the current **Workspace & Export Setup** authentication/authorization workflow, removes all `STEP 1`–`STEP 4` labels, and consistently uses `Alternative` for both the Google Form entry option and CSV export.

### Scope
- **In Scope:**
  - Updating `quickStartSection` (and `tipsSection` if needed for authorization/folder reminders) inside `showHelpCard()` (`ui.js`).
  - Aligning all workflow headings and terminology in `showHelpCard()` with `buildAdminSidebarCard()`, `buildExportCard()`, and `onOpen()` in `ui.js`:
    - Adding **Workspace & Export Setup** (spreadsheet authorization via Google Picker and optional export destination folder selection).
    - Removing all `STEP 1`, `STEP 2`, `STEP 3`, and `STEP 4` labels.
    - Using `Alternative` for online forms (`Alternative: Entry Form`) and CSV export (`Alternative: Export to CSV`).
- **Out of Scope:**
  - Modifying `buildAdminSidebarCard()`, `buildExportCard()`, `showAboutCard()`, `workspace-setup-ui.html`, or `README.md`.

### Functional Requirements
- **FR-1 (Workspace & Export Setup / Authentication):** `quickStartSection` in `showHelpCard()` shall start with **Workspace & Export Setup** (or step `1. Workspace & Export Setup`), explaining that users must authorize the active spreadsheet (`Authorize & Setup`) so the add-on can create/update tables under `drive.file`, optionally choose an export destination folder in Google Drive (defaults to `My Drive (root)`), and use **Refresh Status** if needed.
- **FR-2 (Remove Step Labels):** Remove all `(STEP 1)`, `(STEP 2)`, `(STEP 2 / Optional)`, `(STEP 3, Optional)`, and `(STEP 4)` labels from `quickStartSection` in `showHelpCard()` to match the updated sidebar and menu labels (`Meet Setup`, `Create Team Entry Sheets`, `Optional: Create Relay Tables`, `Export for Meet Manager`).
- **FR-3 ("Alternative" Terminology):** Label both the Google Form entry workflow (**Alternative: Entry Form** via **Create / Sync Entry Form**) and CSV export (**Alternative: Export to CSV**) using `Alternative` instead of `Optional` or ad-hoc phrasing.
- **FR-4 (Preserve Validated Content):** Keep the existing `criticalFieldsSection`, `exceptionReportsSection`, `tipsSection` (competition-type-aware `Date of Birth` vs `School Year`, `Schools`/`Teams` lookup, duplicate-event warnings, and Form re-sync reminder), and `resourcesSection` intact and accurate.

# Technical Design

### Current Implementation
In `ui.js`:
- `buildAdminSidebarCard()` (lines 138–334) and `buildExportCard()` (lines 409–455) have already been updated to:
  - Lead with **Workspace & Export Setup** (`Authorize & Setup` / `Change Export Folder` / `Refresh Status`).
  - Remove `STEP 1`–`STEP 4` from **Meet Setup**, **Create Team Entry Sheets**, **Optional: Create Relay Tables**, and **Export for Meet Manager (Recommended)**.
  - Use **Alternative: Entry Form** and **Alternative: Export to CSV**.
- `showHelpCard()` (`ui.js`, lines 606–726) is out of sync in `quickStartSection` (lines 613–639):
  - Omits **Workspace & Export Setup** (spreadsheet authorization and export folder setup).
  - Still includes `(STEP 1)`, `(STEP 2 / Optional)`, `(STEP 2)`, `(STEP 3, Optional)`, and `(STEP 4)`.
  - Calls the form option `Entry Form (Optional)` instead of `Alternative: Entry Form` and uses `Alternative: Export to CSV` only casually in a sentence.

### Proposed Changes
Update `quickStartSection` (and `tipsSection` if helpful) inside `showHelpCard()` in `ui.js` using `CardService.newTextParagraph().setText(...)` with `CardService`-supported HTML (`<b>`, `<i>`, `<br>`):

1. **`quickStartSection` ("Quick Start Workflow"):**
   - **1. Workspace & Export Setup:** Authorize the active spreadsheet via **Authorize & Setup** (grants `drive.file` access for table creation and updates) and optionally choose a Drive folder for exports (defaults to `My Drive (root)`). Click **Refresh Status** in the sidepanel after closing the setup modal if the status does not update immediately.
   - **2. Meet Setup:** Select your Competition Type (<i>School</i>, <i>College</i>, <i>Age Group</i>, <i>Open</i>, or <i>Masters</i>) and enter meet details to auto-generate <i>Meet</i> (with <i>AgeBrackets</i> for School/College), <i>Events</i>, <i>Schools</i> or <i>Teams</i>, and <i>Individual Events Template</i> (plus optional <i>Team Officials</i> and <i>How to fill entries</i>).
   - **3. Review &amp; Populate Tables:**
     - <i>Events:</i> Enter Distance and Discipline (Event name is auto-generated)
     - <i>Schools / Teams:</i> Add participating organisations with Code and Team Name
     - <i>Team Officials (if enabled):</i> Enter team/school/cluster officials
   - **4. Collect Entries:**
     - <i>Create Team Entry Sheets:</i> Duplicate <i>Individual Events Template</i> per school, cluster, or team via **Create Team Sheets**.
     - <i>Alternative: Entry Form:</i> Generate a public Google Form via **Create / Sync Entry Form** (responses land in <i>Form Entries</i>).
     - <i>Optional: Create Relay Tables:</i> Generate relay sheets if your meet includes relay events.
   - **5. Fill Swimmer Entries:**
     - **Required:** First Name, Last Name, Gender, valid School/Team/Club, plus Date of Birth (<i>Age Group</i>, <i>Open</i>, <i>Masters</i>) or School Year (<i>School</i>, <i>College</i>)
     - **Recommended:** Entry times
   - **6. Export for Meet Manager:**
     - Generate an SDIF (`.sd3`) file and review any Exception Report warnings/errors (saved to your configured export folder or `My Drive (root)`).
     - <i>Alternative: Export to CSV:</i> Export entries as a CSV file for manual processing or external tools.

2. **`tipsSection` ("Best Practices"):**
   - Keep existing tips and ensure they align with the authorization/folder workflow (no `STEP` references).

# Testing

### Validation Approach
1. **Node.js `CardService` Mock Check:**
   - Load `ui.js` in a Node.js `vm` context with a mocked `CardService` and invoke `showHelpCard()`.
   - Verify all text paragraphs render without runtime errors and all `CardService` HTML tags (`<b>`, `<i>`, `<br>`) are balanced.
2. **Content Parity Assertions:**
   - Assert rendered Help text contains `Workspace & Export Setup`, `Authorize & Setup`, `Alternative: Entry Form`, and `Alternative: Export to CSV`.
   - Assert rendered Help text contains **no** occurrences of `STEP 1`, `STEP 2`, `STEP 3`, or `STEP 4` (case-insensitive).

# Delivery Steps

### ✓ Step 1: Update `showHelpCard()` in `ui.js`
`showHelpCard()` in `ui.js` accurately reflects Workspace & Export Setup authentication, removes all `STEP` labels, and uses `Alternative` for both Entry Form and CSV export.

- Add **Workspace & Export Setup** instructions (spreadsheet authorization via `Authorize & Setup`, optional export folder selection, and `Refresh Status`) to `quickStartSection` in `showHelpCard()` (`ui.js`).
- Remove `(STEP 1)`, `(STEP 2)`, `(STEP 2 / Optional)`, `(STEP 3, Optional)`, and `(STEP 4)` from `quickStartSection` in `ui.js`.
- Update the online form and CSV export entries in `quickStartSection` to use **Alternative: Entry Form** and **Alternative: Export to CSV**, and note that exported `.sd3`, `.csv`, and Exception Report files save to the configured export folder.

### ✓ Step 2: Verify `showHelpCard()` rendering and copy parity
The updated Help card builds cleanly with valid `CardService` HTML and passes all parity assertions.

- Run a Node.js `vm` check invoking `showHelpCard()` with a mocked `CardService` to confirm zero runtime errors and balanced `<b>`, `<i>`, and `<br>` tags.
- Assert presence of `Workspace & Export Setup`, `Authorize & Setup`, `Alternative: Entry Form`, and `Alternative: Export to CSV`, and absence of any `STEP 1`–`STEP 4` strings.