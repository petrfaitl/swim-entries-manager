---
sessionId: session-261116-161111-1fwi
---

# Requirements

### Overview & Goals
In the **Meet Setup** dialog (`meet-setup-ui.html`), the **Setup Options** section currently provides a single checkbox to optionally generate the `Team Officials` sheet (`includeOfficials`). However, `tableConfig.js` already defines `SampleIndividualEntries`, which generates a sample `"How to fill entries"` sheet with pre-populated example swimmer data and dropdowns.

This change adds a second checkbox in the **Setup Options** section of the Meet Setup dialog—placed right next to / below the **Include Team Officials sheet** option—allowing meet organizers to optionally generate the sample `"How to fill entries"` sheet during Meet Setup, while ensuring that the generated sample sheet is excluded from SDIF/CSV candidate entry sheet lists.

### Scope
#### In Scope
- **Meet Setup UI (`meet-setup-ui.html`)**: Add a checkbox (`id="includeSampleHowToFill"`) in the **Setup Options** section near `includeOfficials`, and include `includeSampleHowToFill` in the form submission payload.
- **Meet Setup Backend (`tableBuilder.js`)**: Update `saveAndGenerateMeetSetup(formData)` to call `createConfiguredTable('SampleIndividualEntries', configOptions)` when `formData.includeSampleHowToFill` is truthy.
- **Export Sheet Filtering (`SDIFCreator.js`)**: Add `'how to fill entries'` to `reservedNames` in `getEntriesSheetNames_(ss)` so the sample sheet is not listed as a candidate swimmer entry sheet in SDIF/CSV export dialogs.
- **Automated Unit Testing (`tests.js`)**: Add test coverage verifying `saveAndGenerateMeetSetup` generates `SampleIndividualEntries` when `includeSampleHowToFill` is `true` (and skips it when `false`), and that `getEntriesSheetNames_` excludes `"How to fill entries"`.
- **Documentation (`docs/SWIM_ENTRIES_MANAGER_SPEC.md`, `docs/USER_GUIDE.md`)**: Document the optional `"How to fill entries"` sample sheet in Meet Setup and export exclusion rules.

#### Out of Scope
- Modifying the `SampleIndividualEntries` schema, sample data defaults, or target sheet name (`"How to fill entries"`) in `tableConfig.js`.
- Changing `createConfiguredTable` default-value population logic in `tableBuilder.js`.

### User Stories
- **As a Meet Administrator**, I want a checkbox in Meet Setup's **Setup Options** to generate the `"How to fill entries"` sample sheet, so that team coordinators have a clear example of how to fill out swimmer entries without manual sheet creation.
- **As a Meet Administrator**, I want the `"How to fill entries"` sample sheet to be excluded from SDIF/CSV entry sheet export lists, so that sample data is never accidentally exported as real meet entries.

### Functional Requirements
- **FR-1 (UI Checkbox)**: `meet-setup-ui.html` shall display an unchecked checkbox (`id="includeSampleHowToFill"`) in the **Setup Options** section immediately following the `includeOfficials` checkbox, labeled `Include "How to fill entries" sample sheet`.
- **FR-2 (Payload Transmission)**: `submitMeetSetup()` in `meet-setup-ui.html` shall read `document.getElementById('includeSampleHowToFill').checked` and pass `includeSampleHowToFill` in the payload to `saveAndGenerateMeetSetup(formData)`.
- **FR-3 (Conditional Table Generation)**: `saveAndGenerateMeetSetup(formData)` in `tableBuilder.js` shall invoke `createConfiguredTable('SampleIndividualEntries', configOptions)` if and only if `formData.includeSampleHowToFill` is truthy.
- **FR-4 (Export Exclusion)**: `getEntriesSheetNames_(ss)` in `SDIFCreator.js` shall exclude `'how to fill entries'` (case-insensitive) via `reservedNames` so it is not offered as an entry sheet in SDIF or CSV export dialogs.

# Technical Design

### Current Implementation
- `meet-setup-ui.html` (lines 143–149) renders the **Setup Options** section with a single `.checkbox-field` (`#includeOfficials`). `submitMeetSetup()` (lines 177–214) constructs the `payload` object and calls `google.script.run...saveAndGenerateMeetSetup(payload)`.
- `tableConfig.js` (lines 644–899) already defines `SampleIndividualEntries` with `placement: { targetSheet: "How to fill entries", startCell: "A1" }`, `rows: 10`, and `default` values on row 2 columns.
- `tableBuilder.js` defines `createConfiguredTable(tableName, options)` (line 238), which creates the target sheet, writes headers, populates `meta.default` values in `dataStartRow`, and binds the structured table via `ensureStructuredTable_`. `saveAndGenerateMeetSetup(formData)` (lines 377–476) calls `createConfiguredTable('TeamOfficials', configOptions)` when `formData.includeOfficials` is true, but has no branch for `SampleIndividualEntries`.
- `SDIFCreator.js` defines `getEntriesSheetNames_(ss)` (line 1234), which filters out `reservedNames` (`'meet'`, `'events'`, `'schools'`, `'teams'`, `'team officials'`, `'sheet1'`, `'sheet 1'`) and checks `sheetHasSwimmerHeaders_(sheet)`. Because `"How to fill entries"` has `First Name`, `Last Name`, `Gender`, `Date of Birth`, and `School Year` headers, it would otherwise be detected as a candidate entry sheet.

### Key Decisions
- **Decision 1: Reuse `SampleIndividualEntries` directly via `createConfiguredTable`**
  - *Rationale*: `SampleIndividualEntries` is already fully configured in `tableConfig.js` with headers, validations, and sample defaults, and `createConfiguredTable` already writes `meta.default` values to the first data row. No schema or builder changes are required.
- **Decision 2: Exclude `"How to fill entries"` in `getEntriesSheetNames_`**
  - *Rationale*: Since the sample sheet uses standard swimmer entry headers, adding `'how to fill entries'` to `reservedNames` in `SDIFCreator.js` prevents sample data from appearing in SDIF/CSV export dropdowns.

### Proposed Changes

#### 1. `meet-setup-ui.html`
- In the **Setup Options** `<div class="section">` (after `#includeOfficials`), add:
  ```html
  <label class="checkbox-field">
    <input type="checkbox" id="includeSampleHowToFill" />
    <span>Include "How to fill entries" sample sheet</span>
  </label>
  ```
- In `submitMeetSetup()`:
  - Read `const includeSampleHowToFill = document.getElementById('includeSampleHowToFill').checked;`
  - Include `includeSampleHowToFill: includeSampleHowToFill` in `payload`.

#### 2. `tableBuilder.js`
- In `saveAndGenerateMeetSetup(formData)` (after step 5 `if (formData.includeOfficials)`):
  ```javascript
  // 6. Optionally create Sample Individual Entries ("How to fill entries") sheet
  if (formData.includeSampleHowToFill) {
    createConfiguredTable('SampleIndividualEntries', configOptions);
  }
  ```

#### 3. `SDIFCreator.js`
- In `getEntriesSheetNames_(ss)`, add `'how to fill entries'` to `reservedNames`:
  ```javascript
  const reservedNames = new Set([
    'meet',
    'events',
    'schools',
    'teams',
    'team officials',
    'how to fill entries',
    'sheet1',
    'sheet 1'
  ]);
  ```

#### 4. `tests.js` and `docs/`
- Add unit tests in `tests.js` covering `saveAndGenerateMeetSetup` with `includeSampleHowToFill` true/false and `getEntriesSheetNames_` excluding `"How to fill entries"`.
- Update `docs/SWIM_ENTRIES_MANAGER_SPEC.md` and `docs/USER_GUIDE.md` to document the optional sample sheet in Meet Setup and its exclusion from export sheet lists.

### Data Models / Contracts
```javascript
// Payload passed from meet-setup-ui.html to saveAndGenerateMeetSetup(formData)
{
  meetName: string,
  competitionType: string,
  startDate: string,
  endDate: string,
  course: string,
  region: string,
  venue: string,
  city: string,
  ageUpDate: string,
  includeOfficials: boolean,
  includeSampleHowToFill: boolean
}
```

### File Structure
- `meet-setup-ui.html` *(modified)*: Add `#includeSampleHowToFill` checkbox in Setup Options and pass it in `submitMeetSetup()`.
- `tableBuilder.js` *(modified)*: Conditionally call `createConfiguredTable('SampleIndividualEntries', configOptions)` in `saveAndGenerateMeetSetup`.
- `SDIFCreator.js` *(modified)*: Exclude `'how to fill entries'` in `getEntriesSheetNames_`.
- `tests.js` *(modified)*: Add unit tests for Meet Setup sample sheet option and export sheet filtering.
- `docs/SWIM_ENTRIES_MANAGER_SPEC.md` *(modified)*: Document optional sample sheet and export exclusion.
- `docs/USER_GUIDE.md` *(modified)*: Document the new Setup Options checkbox.

# Testing

### Validation Approach
Automated verification will run via `tests.js` in the Node.js GAS sandbox, validating Meet Setup payload handling, table generation delegation, and export candidate sheet filtering.

### Key Scenarios
1. **Meet Setup with `includeSampleHowToFill: true`**:
   - Calling `saveAndGenerateMeetSetup` with `includeSampleHowToFill: true` invokes `createConfiguredTable('SampleIndividualEntries', configOptions)`.
2. **Meet Setup with `includeSampleHowToFill: false` (default)**:
   - Calling `saveAndGenerateMeetSetup` with `includeSampleHowToFill: false` or omitted does not create `SampleIndividualEntries`.
3. **Export Candidate Sheet Filtering**:
   - A workbook containing `"How to fill entries"` alongside real entry sheets excludes `"How to fill entries"` from `getEntriesSheetNames_(ss)`.

# Delivery Steps

### ✓ Step 1: Add Setup Options checkbox in `meet-setup-ui.html` and wire generation in `tableBuilder.js`
Meet Setup allows users to opt into generating the `"How to fill entries"` sample sheet alongside `Team Officials`.

- Add `<input type="checkbox" id="includeSampleHowToFill" />` inside the **Setup Options** section of `meet-setup-ui.html` directly below `#includeOfficials`.
- Update `submitMeetSetup()` in `meet-setup-ui.html` to read `includeSampleHowToFill` and include it in the `payload` sent to `saveAndGenerateMeetSetup`.
- Update `saveAndGenerateMeetSetup(formData)` in `tableBuilder.js` to call `createConfiguredTable('SampleIndividualEntries', configOptions)` when `formData.includeSampleHowToFill` is truthy.

### ✓ Step 2: Exclude sample sheet from export lists in `SDIFCreator.js`, add tests, and update docs
The `"How to fill entries"` sample sheet is excluded from SDIF/CSV export sheet selection, tested, and documented.

- Add `'how to fill entries'` to `reservedNames` in `getEntriesSheetNames_(ss)` in `SDIFCreator.js`.
- Add unit tests in `tests.js` verifying `saveAndGenerateMeetSetup` with `includeSampleHowToFill` true/false and `getEntriesSheetNames_` exclusion of `"How to fill entries"`.
- Run the diagnostic test suite in `tests.js` to confirm all tests pass.
- Update `docs/SWIM_ENTRIES_MANAGER_SPEC.md` and `docs/USER_GUIDE.md` to document the new Meet Setup option and export exclusion.