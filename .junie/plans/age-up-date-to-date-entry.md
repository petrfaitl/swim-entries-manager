---
sessionId: session-261109-222811-xolk
---

# Requirements

### Overview & Goals
The goal of this change is to transition the **Age-up Date** field in Swim Entries Manager from a fixed dropdown selection (`Start Date`, `1st January`, `31st December`) into a concrete calendar date entry (`<input type="date">` / `DATE_IS_VALID`), matching the behavior and format of `Start Date` and `End Date`.

This provides meet organizers with maximum operational flexibility and precision, eliminating ambiguity in age calculation and ensuring consistency across all date fields in both Google Sheets tables and downstream SDIF exports.

### Scope
- **In Scope**:
  - Update `Meet` table definition in `tableConfig.js` so that `Age-up Date` uses `type: 'date'` with `DATE_IS_VALID` validation rule.
  - Update `meet-setup-ui.html` to render a native date picker (`<input type="date" id="ageUpDate">`) instead of a `<select>` dropdown.
  - Update client-side validation, date formatting (`DD/MM/YYYY`), and payload creation in `meet-setup-ui.html`.
  - Update diagnostic tests in `tests.js` to verify the new column schema.
  - Update documentation references if applicable.
- **Out of Scope**:
  - Altering the business logic determining which competition types require an Age-up Date (`College`, `Age Group`, `Masters`).
  - Changing SDIF export file formatting rules beyond accepting the standard date string.

### User Stories
- **As a meet organizer**, I want to pick an explicit calendar date for the Age-up Date so that swimmer ages are computed according to our exact meet regulations.
- **As a spreadsheet administrator**, I want the `Age-up Date` column in the `Meet` table to enforce valid date formatting in Google Sheets just like `Start Date` and `End Date`, preventing invalid or malformed text entries.

### Functional Requirements
- **FR-1**: The `Meet` table column `Age-up Date` shall use data type `date` with Google Sheets data validation rule `DATE_IS_VALID`.
- **FR-2**: In the **Meet Setup** modal dialog (`meet-setup-ui.html`), when a competition type requiring an Age-up Date (`College`, `Age Group`, `Masters`) is selected, an HTML5 date input (`<input type="date" id="ageUpDate">`) shall be displayed and marked required.
- **FR-3**: When switching to a competition type that does not use an Age-up Date (`School`, `Open`), the Age-up Date input shall be hidden, non-required, and cleared.
- **FR-4**: When the meet setup is saved, `ageUpDate` shall be formatted into standard `DD/MM/YYYY` display format and populated into row 2 of the `Meet` table.

# Technical Design

### Current Implementation
- In `tableConfig.js`, the `Meet` table defines `'Age-up Date'` as:
  ```javascript
  'Age-up Date': {
    type: 'text',
    validation: {
      type: 'list',
      args: { values: ['Start Date', '1st January', '31st December'] }
    }
  }
  ```
- In `meet-setup-ui.html`, the Age-up Date is rendered as a `<select>` dropdown with three fixed text options:
  ```html
  <div class="field" id="ageUpContainer" style="display: none;">
    <label for="ageUpDate">Age-up Date <span class="required">*</span></label>
    <select id="ageUpDate">
      <option value="Start Date" selected>Start Date</option>
      <option value="1st January">1st January</option>
      <option value="31st December">31st December</option>
    </select>
  </div>
  ```
- In `meet-setup-ui.html` JS, `submitMeetSetup()` reads the string directly from `document.getElementById('ageUpDate').value` without passing it through `formatDateToDisplay()`.

### Key Decisions
- **Decision 1: Use `type: 'date'` with `DATE_IS_VALID` in `tableConfig.js`**
  - *Rationale*: Aligns `Age-up Date` with `Start Date` and `End Date` in the `Meet` table schema, ensuring uniform spreadsheet validation and cell formatting behavior.
- **Decision 2: Use native `<input type="date">` in `meet-setup-ui.html`**
  - *Rationale*: Reuses existing CSS styles for date inputs, offers standard browser calendar pickers, and integrates smoothly with the existing `formatDateToDisplay()` helper function.

### Proposed Changes

#### 1. Schema Configuration (`tableConfig.js`)
Update `Meet` table column definition:
```javascript
'Age-up Date': {
  type: 'date',
  validation: {
    type: 'date',
    args: { condition: 'DATE_IS_VALID' }
  }
}
```

#### 2. Dialog Form & Client Script (`meet-setup-ui.html`)
- Replace `<select id="ageUpDate">` with `<input type="date" id="ageUpDate" />`.
- In `handleCompTypeChange()`, reset `ageUpDate.value = ''` when hiding.
- In `submitMeetSetup()`, format the date value via `formatDateToDisplay(ageUpDateRaw)` before sending to `saveAndGenerateMeetSetup(payload)`.
- Enforce validation error if a required competition type has an empty `ageUpDateRaw`.

#### 3. Test Suite (`tests.js`)
- Update `testTableConfigGeneration()` to verify that `schoolCfg.tables.Meet.columns['Age-up Date'].type === 'date'` and validation condition is `'DATE_IS_VALID'`.

### File Structure
- `tableConfig.js` — Core table definitions and column validation rules.
- `meet-setup-ui.html` — Meet Setup modal dialog template and client-side logic.
- `tests.js` — Script diagnostic assertions.
- `docs/SWIM_ENTRIES_MANAGER_SPEC.md` — Core functional specification.

# Testing

### Validation Approach
Automated and scripted verifications will be conducted using the built-in diagnostic test runner in `tests.js`, as well as static code inspection of schema objects and HTML elements.

### Key Scenarios
- **Scenario 1: Meet Table Generation with Date Validation**
  - Generate the `Meet` table configuration using `getTableConfig({ competitionType: 'college' })`.
  - Validate that `Age-up Date` column has `type: 'date'` and validation `condition: 'DATE_IS_VALID'`.
- **Scenario 2: UI Dialog Field Toggle & Submission**
  - Select `College`, `Age Group`, or `Masters` in Meet Setup: Age-up Date field appears as a date picker and is required.
  - Select `School` or `Open`: Age-up Date field is hidden and optional.
  - Submit the form with an Age-up Date selected: Verify payload contains date in `DD/MM/YYYY` format.
- **Scenario 3: Meet Sheet Table Population**
  - Execute `saveAndGenerateMeetSetup(formData)` with a specific date (e.g. `2026-10-31` → `31/10/2026`).
  - Verify that cell E2 in `Meet` sheet contains the expected date value with date validation applied.

### Test Changes
- Update `testTableConfigGeneration()` in `tests.js` to assert the updated date type and validation condition on the `Age-up Date` column.

# Delivery Steps

### ✓ Step 1: Update Meet table schema and validation rules in tableConfig.js
Update the schema and validation configuration in `tableConfig.js` and verify table generation diagnostics in `tests.js`.

- Change the `Age-up Date` column definition under `Meet` table in `tableConfig.js` from a dropdown list (`type: 'text'`, values: `['Start Date', '1st January', '31st December']`) to a valid date cell (`type: 'date'`, validation condition `DATE_IS_VALID`), identical to `Start Date` and `End Date`.
- Update `tests.js` in `testTableConfigGeneration()` to assert that the `Age-up Date` column has date type and `DATE_IS_VALID` validation.
- Verify that default competition type presets in `tableConfig.js` (`COLLEGE`, `AGE_GROUP`, `MASTERS`) retain their boolean flag `hasAgeUpDate: true` for downstream UI gating.

### ✓ Step 2: Update Meet Setup UI dialog and date payload handling
Revamp the Meet Setup dialog in `meet-setup-ui.html` to capture an exact calendar date for Age-up Date and pass formatted date strings to server generation.

- Replace the `<select id="ageUpDate">` dropdown element in `meet-setup-ui.html` with an `<input type="date" id="ageUpDate" />` field matching the styling of `startDate` and `endDate`.
- Update `handleCompTypeChange()` to properly toggle required validation and clear state on the date input when switching competition types.
- Update `submitMeetSetup()` to format `ageUpDate` via `formatDateToDisplay` (`DD/MM/YYYY`) before passing it in the generation payload.
- Update `docs/SWIM_ENTRIES_MANAGER_SPEC.md` if necessary to document the date entry schema for Age-up Date.