---
sessionId: session-261117-091644-2mkw
---

# Requirements

### Overview & Goals
Update the **Help & Quick Start** card (`showHelpCard()` in `ui.js`) in the Google Sheets Add-on sidepanel so every workflow step, field requirement, exception report item, and best-practice tip matches the current functionality of **Swim Entries Manager**.

### Scope
- **In Scope:**
  - Updating `quickStartSection`, `criticalFieldsSection`, `exceptionReportsSection`, and `tipsSection` inside `showHelpCard()` (`ui.js`).
  - Aligning terminology with the current sidebar (`buildAdminSidebarCard` and `buildExportCard`) and Add-on menu (`onOpen`): **Meet Setup**, **Create Team Entry Sheets**, **Entry Form (Optional)**, **Create Relay Tables (Optional)**, **Export for Meet Manager**, and **Export to CSV**.
- **Out of Scope:**
  - Modifying `buildAdminSidebarCard()`, `buildExportCard()`, or `showAboutCard()` in `ui.js`.
  - Rewriting `docs/USER_GUIDE.md` or `docs/SWIM_ENTRIES_MANAGER_SPEC.md`.

### Functional Requirements
- **FR-1 (Quick Start Workflow):** Replace references to "Create Core Tables" with **Meet Setup (STEP 1)**, describing competition type selection (`School`, `College`, `Age Group`, `Open`, `Masters`) and auto-generated sheets (`Meet`, `AgeBrackets` for School/College, `Events`, `Schools`/`Teams`, `Individual Events Template`, and optional `Team Officials` / `How to fill entries`).
- **FR-2 (Entry Collection & Relays):** Document both **Create Team Entry Sheets (STEP 2)** and **Entry Form (Optional)** (`Create / Sync Entry Form`), as well as **Create Relay Tables (STEP 3, Optional)**.
- **FR-3 (Critical Data Requirements):** Reflect competition-type-aware requirements: `Date of Birth` is required for `Age Group`, `Open`, and `Masters`, whereas `School` and `College` meets accept `School Year` (mapped via `AgeBrackets` on `Meet`) when `Date of Birth` is omitted; team codes/names are validated against `Schools` (`School`/`College`) or `Teams` (`Age Group`/`Open`/`Masters`).
- **FR-4 (Exception Reports & Best Practices):** Reference both `Schools` and `Teams` tables, include duplicate-event warnings, and advise re-syncing the Entry Form when `Events` or `Schools`/`Teams` change.

# Technical Design

### Current Implementation
`showHelpCard()` in `ui.js` (lines 467–581) builds a `CardService` card with five sections:
1. `quickStartSection` — still describes the removed "Create Core Tables (STEP 1)" dialog, omits `AgeBrackets`, omits the `Teams` table alternative, omits **Entry Form (Optional)** (`form-setup-ui.html` / `formBuilder.js`), and omits **Create Relay Tables (STEP 3)**.
2. `criticalFieldsSection` — states `Date of Birth (DD/MM/YYYY)` is always required and only references the `Schools` table.
3. `exceptionReportsSection` — only mentions missing teams in `Schools` table and omits duplicate-event warnings added in `SDIFCreator.js` / `EntryDataProcessor.js`.
4. `tipsSection` — only mentions adding teams to the `Schools` table and does not mention re-syncing the Entry Form.
5. `resourcesSection` — links to `docs/USER_GUIDE.md`, GitHub issues, and `recorder@lvwasc.co.nz` (already up to date).

### Proposed Changes
All changes are localized to `showHelpCard()` in `ui.js` using `CardService.newTextParagraph().setText(...)` with `CardService`-supported HTML (`<b>`, `<i>`, `<br>`):

- **`quickStartSection` ("Quick Start Workflow"):**
  - **1. Meet Setup (STEP 1):** Select Competition Type (`School`, `College`, `Age Group`, `Open`, `Masters`) and enter meet details to auto-generate `Meet` (with `AgeBrackets` for School/College), `Events`, `Schools` or `Teams`, and `Individual Events Template` (plus optional `Team Officials` and `How to fill entries` sample).
  - **2. Review & Populate Tables:** Verify `Events` (`Distance` and `Discipline`; `Event` is auto-generated), add participating organisations to `Schools` or `Teams` (with `Code` and `Team Name`), and populate `Team Officials` if enabled.
  - **3. Collect Entries (STEP 2 / Optional):** Duplicate `Individual Events Template` per school, cluster, or team via **Create Team Sheets**, or generate a public Google Form via **Create / Sync Entry Form** (responses land in `Form Entries`). Optionally run **Create Relay Tables (STEP 3)** for relay meets.
  - **4. Fill Swimmer Entries:** **Required:** `First Name`, `Last Name`, `Gender`, valid `School`/`Team`/`Club`, plus `Date of Birth` (Age Group/Open/Masters) or `School Year` (School/College). **Recommended:** Entry times.
  - **5. Export for Meet Manager (STEP 4):** Generate an SDIF (`.sd3`) file and review any Exception Report warnings/errors. Alternative: **Export to CSV**.
- **`criticalFieldsSection` ("Critical Data Requirements"):**
  - Clarify that `Date of Birth (DD/MM/YYYY)` is required for `Age Group`, `Open`, and `Masters`, while `School` and `College` meets can use `School Year` (mapped in `AgeBrackets`) when DOB is omitted.
  - State that `Team Code` and `Team Name` must exist in the `Schools` or `Teams` table.
  - Retain the existing `[Distance] [Stroke]` event format examples.
- **`exceptionReportsSection` ("Understanding Exception Reports"):**
  - Update `"Missing teams in Schools table"` to `"Missing teams in Schools / Teams table"`.
  - Add `"Duplicate event entries (duplicates skipped, warning only)"`.
- **`tipsSection` ("Best Practices"):**
  - Update `"Add all teams to Schools table first"` to `"Add all teams/schools to the Schools or Teams table first"`.
  - Add `"Re-sync your Entry Form after updating Events, Schools, or Teams"`.

# Testing

### Validation Approach
1. **Static & Runtime Check in Node.js:**
   - Load `ui.js` in a Node.js `vm` context with a lightweight `CardService` mock (`newCardBuilder`, `newCardHeader`, `newCardSection`, `newTextParagraph`, `newButtonSet`, `newTextButton`, `newOpenLink`, `newNavigation`, `Icon`).
   - Invoke `showHelpCard()` and assert that all text paragraphs render without runtime errors and that all `CardService` HTML tags (`<b>`, `<i>`, `<br>`) are balanced.
2. **Parity & Regression Assertions:**
   - Verify the rendered Help text no longer contains `"Create Core Tables"`.
   - Verify the rendered Help text mentions `Meet Setup`, `Entry Form`, `Create Relay Tables`, `Schools` and `Teams`, `School Year` / `AgeBrackets`, and duplicate-event warnings.

# Delivery Steps

### ✓ Step 1: Update Quick Start and Critical Data sections in showHelpCard
The Quick Start Workflow and Critical Data Requirements sections of `showHelpCard()` in `ui.js` accurately reflect Meet Setup, Entry Forms, Relay Tables, and competition-type-aware swimmer fields.

- Replace the legacy "Create Core Tables (STEP 1)" instructions in `quickStartSection` with **Meet Setup (STEP 1)**, covering competition types (`School`, `College`, `Age Group`, `Open`, `Masters`) and auto-generated sheets (`Meet` with `AgeBrackets`, `Events`, `Schools`/`Teams`, `Individual Events Template`, plus optional `Team Officials` and `How to fill entries`).
- Update table population guidance in `quickStartSection` for `Events`, `Schools` / `Teams`, and optional `Team Officials`.
- Update entry collection steps to cover both **Create Team Entry Sheets (STEP 2)** and **Entry Form (Optional)** (`Create / Sync Entry Form` -> `Form Entries`), and add **Create Relay Tables (STEP 3, Optional)**.
- Update swimmer entry and `criticalFieldsSection` requirements to distinguish `Date of Birth` (`Age Group`, `Open`, `Masters`) vs `School Year` + `AgeBrackets` (`School`, `College`), and `Schools` vs `Teams` table lookups.

### ✓ Step 2: Update Exception Reports, Best Practices, and verify showHelpCard
The Exception Reports and Best Practices sections of `showHelpCard()` in `ui.js` match current validation behavior, and the card renders cleanly without errors.

- Update `exceptionReportsSection` in `ui.js` to reference `Schools / Teams table` and include duplicate-event warnings alongside missing critical data, invalid event formats, and swimmers with no events.
- Update `tipsSection` in `ui.js` to reference `Schools or Teams` tables and remind users to re-sync the Entry Form (`Create / Sync Entry Form`) after updating `Events`, `Schools`, or `Teams`.
- Verify `showHelpCard()` in a Node.js check with a mocked `CardService` to confirm the card builds cleanly, all HTML tags (`<b>`, `<i>`, `<br>`) are valid, and no references to `Create Core Tables` remain.