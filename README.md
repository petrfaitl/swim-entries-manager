# Swim Entries Manager

[![Google Workspace Add-on](https://img.shields.io/badge/Google%20Workspace-Add--on-4285F4?logo=google)](https://workspace.google.com/)
[![Apps Script](https://img.shields.io/badge/Apps%20Script-V8-34A853)](https://developers.google.com/apps-script)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A Google Workspace Add-on for managing swim meet entries efficiently. This tool helps swim meet organisers create structured entry templates, manage school/cluster registrations, and export data directly to SDIF format for Hy-Tek Meet Manager or CSV for other tools.

![Swim Entries Manager](https://www.lvwasc.co.nz/wp-content/uploads/2025/07/Swim-Entries-Manager.png)

## Features

### 🔐 Workspace & Export Setup
- **One-Click Sheet Authorization**: Authorize the active spreadsheet via Google Picker under the non-sensitive `drive.file` scope so `TableApp` (`Sheets.Spreadsheets.batchUpdate`) works without broad spreadsheet permissions
- **Per-Spreadsheet Export Destination**: Choose any My Drive or Shared Drive folder for `.sd3`, `.csv`, and `_ExceptionReport.txt` exports (persisted per spreadsheet, with My Drive root fallback)
- **In-Place Sidepanel Refresh**: Click **Refresh Status** in the sidepanel to immediately reflect authorization and export folder updates

### 🏊 Meet & Template Management
- **Meet Setup (STEP 1)**: Competition-type-aware setup supporting `School`, `College`, `Age Group`, `Open`, and `Masters` meets
- **Auto-Generated Core Sheets**: Creates `Meet` (including `AgeBrackets` for `School`/`College`), `Events`, `Schools` or `Teams`, `Individual Events Template`, and optional `Team Officials` and `How to fill entries` sample sheets
- **Automated Validation & Named Ranges**: Built-in dropdown validations and dynamic named ranges for events, schools/teams, genders, and school years
- **Google Sheets Tables**: Uses the bundled `TableApp` (`TableApp_Local.js`) for structured tables and formatting

### 📋 Sheet Duplication & Relay Tables
- **Create Team Sheets (STEP 2)**: Bulk-duplicate `Individual Events Template` for multiple schools, clusters, or teams with automatic naming, collision detection, and a **View sheet** preview button
- **Create Relay Tables (STEP 3, Optional)**: Generate structured relay entry tables for relay meets

### 📝 Google Form Self-Entry
- **Create / Sync Entry Form**: Generate a public-facing Google Form linked directly to a `Form Entries` response sheet
- **Configurable Form Options**:
  - 1 to 9 `Event` / `Time` slots per swimmer with regex seed-time validation
  - Optional responder email collection and response editing
  - Optional `Date of Birth` field for `School`/`College` meets (omitted when unchecked to avoid confusion with school years)
  - Optional write-in `"Other"` team/club option for `Age Group`, `Open`, and `Masters` meets
- **In-Place Option Syncing**: Re-sync updated `Events`, `Schools`, or `Teams` lists to the existing Form without creating duplicate Forms or duplicate columns in `Form Entries`

### 📊 Data Export & Validation
- **SDIF Export (Primary)**: Direct export to SDIF v3 (`.sd3`) format for Hy-Tek Meet Manager
- **CSV Export (Alternative)**: Traditional `.csv` format for custom processing
- **`Form Entries` Auto-Detection**: Automatically detects and processes `Form Entries` response sheets alongside team entry sheets
- **`Schools` & `Teams` Table Awareness**: Resolves team codes and reports exceptions against `Schools` (`School`/`College`) or `Teams` (`Age Group`/`Open`/`Masters`)
- **Exception Reporting & Duplicate Warnings**: Validates critical swimmer/team fields and flags duplicate-event entries (skipping duplicates with a warning)
- **Multiple Entry Format Support**:
  - Method 1: Event names in cells (e.g., `"Event 1"` column header, `"100m Freestyle"` in cell)
  - Method 2: Event names in headers (e.g., `"100m Freestyle"` header, `"Yes"` or time in cell)
- **Flexible Column Detection**: Automatically handles sheets with or without `#` numbering, `Timestamp`, or `Email Address` columns
- **Smart Time Format Parsing**: Accepts multiple time formats (`MM:SS.SS`, `SS.SS`, `SS:DD`, etc.)
- **Stroke-Aware Pool Conversions**:
  - Supports 33m pool, 18m pool, and Long Course (50m) conversions
  - Individual Medley events use correct source distances (e.g., 133.2m vs 99.9m for 100m in 33m pool)
- **Smart Export Destination**: Saves `.sd3`, `.csv`, and `_ExceptionReport.txt` files to the configured Drive folder (falling back to My Drive root if unconfigured or unwritable)

## Quick Start

### Installation

#### Test Deployment (Current)
1. Open your Google Spreadsheet
2. Navigate to **Extensions** > **Apps Script**
3. Copy all files from this repository into your Apps Script project
4. Deploy as Test Deployment
5. Refresh your spreadsheet
6. Access via the sidebar or **Extensions** > **Swim Entries Manager**

#### From Google Workspace Marketplace (Coming Soon)
1. Visit the [Google Workspace Marketplace](#)
2. Search for "Swim Entries Manager"
3. Click "Install"
4. Grant necessary permissions
5. Open any Google Spreadsheet to use the add-on

### First Use

0. **Workspace & Export Setup (First-Time Sheet Authorization)**:
   - At the top of the **Swim Entries Manager** sidepanel (or via **Extensions** > **Swim Entries Manager** > **Workspace & Export Setup** / **Set Export Destination**), click **Authorize & Setup**.
   - Select the active spreadsheet in Google Picker to grant `drive.file` access for table creation, and optionally choose a My Drive or Shared Drive folder for exports.
   - Click **Refresh Status** in the sidepanel after closing the modal.
1. **Meet Setup (STEP 1)**:
   - Open **Meet Setup** from the sidepanel or menu.
   - Choose your **Competition Type** (`School`, `College`, `Age Group`, `Open`, or `Masters`), enter meet details, and optionally include `Team Officials` and the `How to fill entries` sample sheet.
   - Generates `Meet` (with `AgeBrackets` for `School`/`College`), `Events`, `Schools` or `Teams`, and `Individual Events Template`.
2. **Populate tables**:
   - `Events`: Enter `Distance` (25m, 50m, 100m, 200m, 400m) and `Discipline` (Freestyle, Backstroke, Breaststroke, Butterfly, Individual Medley); the `Events` column builds automatically.
   - `Schools` or `Teams`: Fill participating organisations with `Code` and `Team Name` (plus `School` and `Cluster` for `School`/`College` meets).
   - `Team Officials` (if enabled): Enter team, school, or cluster coordinator contact details.
3. **Collect Entries — Team Sheets (STEP 2) or Entry Form (Optional)**:
   - **Create Team Sheets (STEP 2)**: Verify the dropdowns in `Individual Events Template`, then duplicate one entry sheet per school, cluster, or team (use **View sheet** to preview your source table).
   - **Create / Sync Entry Form (Optional)**: Generate a public Google Form whose responses flow directly into `Form Entries`. Re-sync whenever `Events`, `Schools`, or `Teams` change.
4. **Create Relay Tables (STEP 3, Optional)**:
   - Run **Create Relay Tables** if your meet includes relay events.
5. **Fill swimmer entries**:
   - **Required**: `First Name`, `Last Name`, `Gender`, valid `School`/`Team`/`Club`, plus `Date of Birth` (`Age Group`/`Open`/`Masters`) or `School Year` mapped via `AgeBrackets` (`School`/`College`).
   - **Recommended**: Entry times (`1:30.95`, `45.67`, `34:56` = 34.56 sec) and pool conversion selection (33m, 18m, or LC) when needed.
   - Ensure all organisations exist in the `Schools` or `Teams` table.
6. **Export for Meet Manager (STEP 4)**:
   - Use **Export for Meet Manager** to generate an SDIF (`.sd3`) file (or **Export to CSV** for `.csv`).
   - The exporter automatically detects `Form Entries` and team sheet layouts, checks against `Schools`/`Teams`, flags duplicate events, and saves exports and `_ExceptionReport.txt` to your configured Drive folder (or My Drive root).

## Project Structure

```
SwimEntriesManager/
├── AGENTS.md                      # Agent instructions and project context
├── CONTRIBUTING.md                # Contribution guidelines
├── LICENSE                        # MIT license
├── README.md                      # This file
├── appsscript.json                # Project manifest, OAuth scopes, and services
├── ui.js                          # User interface (sidebar, menu, modal launchers)
├── workspace-setup-ui.html        # Workspace & Export Setup modal (Google Picker)
├── shared-dialog-styles.html      # Shared CSS styles for modal dialogs
├── shared-drive-picker.html       # Export destination bar for export dialogs
├── meet-setup-ui.html             # Meet Setup modal dialog UI
├── core-tables-ui.html            # Legacy core tables modal dialog UI
├── tableConfig.js                 # Table definitions, schemas, and competition presets
├── tableBuilder.js                # Table creation, Meet Setup, and named ranges
├── TableApp_Local.js              # Bundled TableApp library for structured tables
├── duplicateTemplate.js           # Sheet duplication logic with collision detection
├── sheets-from-template-ui.html   # Create Team Sheets modal dialog UI
├── relay-ui-full.html             # Create Relay Tables modal dialog UI
├── formBuilder.js                 # Google Form entry generator & in-place sync
├── form-setup-ui.html             # Create / Sync Entry Form modal dialog UI
├── EntryDataProcessor.js          # Shared data parsing, validation, and normalization
├── resizeSwim.js                  # Pool distance & course conversion algorithms
├── SDIFCreator.js                 # SDIF v3 (.sd3) export with exception reporting
├── SDIFCreator_Dialog.html        # SDIF export modal dialog UI
├── exportToCSV.js                 # CSV export generation and Drive persistence
├── export-csv-ui.html             # CSV export modal dialog UI
├── utils.js                       # Utility functions, Drive folder & Picker helpers
├── tests.js                       # Diagnostic and unit tests
└── docs/                          # Documentation
    ├── SWIM_ENTRIES_MANAGER_SPEC.md # Functional specification
    ├── USER_GUIDE.md                # Comprehensive user guide
    ├── PRIVACY_POLICY.md            # Privacy policy
    └── TERMS_OF_SERVICE.md          # Terms of service
```

## Configuration

### Table Configuration

Tables are defined in `tableConfig.js`. Each table includes:
- **Headers**: Column names
- **Columns**: Data types, validation rules, formulas
- **Options**: Styling, placement, named ranges

Example:
```javascript
Schools: {
  title: 'Schools for Individual Entries Template',
  headers: ['Team Name', 'School', 'Cluster', 'Code'],
  columns: {
    'School': { type: 'text' },
    'Cluster': {
      type: 'text',
      validation: {
        type: 'list',
        args: { values: ['Cluster One', 'Cluster Two'] }
      }
    }
  },
  options: {
    freezeHeader: 1,
    headerBg: '#356853',
    rows: 20,
    placement: { targetSheet: "SchoolsForTemplate", startCell: 'A1' },
    namedRange: { name: 'SchoolsList', columnName: 'School' }
  }
}
```

### Named Ranges

Named ranges are automatically created for:
- **SchoolsList**: References the 'School' column in the Schools table
- **EventsList**: References the 'Events' column in the EventsForTemplate table

These enable dropdown validation across multiple sheets.

## Dependencies

- **Google Sheets API v4**: Advanced spreadsheet operations (`Sheets` service)
- **Google Drive API v3**: File/folder metadata and Drive file persistence (`Drive` service)
- **Google Picker API**: Interactive spreadsheet authorization and Drive/Shared Drive export folder selection
- **TableApp (`TableApp_Local.js`)**: Bundled structured table management for Google Sheets (no external library dependency)

## Permissions

The add-on uses non-sensitive/non-restricted OAuth scopes matching `appsscript.json` (no broad `spreadsheets` scope):
- `spreadsheets.currentonly`: Read/write the active spreadsheet
- `drive.file`: Access the Picker-authorized spreadsheet and create/manage add-on-created exports and Picker-selected folders
- `userinfo.email`: User email identification
- `userinfo.profile`: User profile identification
- `script.container.ui`: Display custom sidepanel and modal UI
- `script.locale`: Localization support
- `forms`: Create and sync Google Entry Forms

To enable Google Picker in **Workspace & Export Setup** (both spreadsheet authorization and **Choose Folder...**), set `GOOGLE_PICKER_API_KEY` and `GOOGLE_CLOUD_PROJECT_NUMBER` in the Apps Script project's **Script Properties** (with the Google Picker API enabled in your Google Cloud project).

## Development

### Local Development

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/swim-entries-manager.git
   ```

2. Install [clasp](https://github.com/google/clasp) (Google Apps Script CLI):
   ```bash
   npm install -g @google/clasp
   ```

3. Login to clasp:
   ```bash
   clasp login
   ```

4. Create a new Apps Script project or link to existing:
   ```bash
   clasp create --type sheets --title "Swim Entries Manager"
   # or
   clasp clone 1sLvLkVjHiO44jf7iKPi3S0FTgA8GLVlCSiGJMlW6eh2PPCDymltl-JoQ
   ```

5. Push changes:
   ```bash
   clasp push
   ```

### Testing

Run the diagnostic suite from the Apps Script editor (or in Node.js with mocked Apps Script services):
```javascript
// In tests.js
runAllDiagnosticTests();
```

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for details.

### Development Workflow

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Recent Updates

### v2.3 - Workspace & Export Setup & Non-Sensitive Scopes
- ✅ **Workspace & Export Setup**: Dedicated Google Picker modal (`workspace-setup-ui.html`) for one-click spreadsheet authorization under `drive.file` and per-spreadsheet My Drive / Shared Drive export folder selection
- ✅ **Non-Sensitive OAuth Scopes**: Removed the broad `spreadsheets` scope in favor of `spreadsheets.currentonly` + `drive.file` (with automatic `SHEET_NOT_AUTHORIZED` guard before table creation)
- ✅ **In-Place Sidepanel Refresh**: Added **Refresh Status** in the sidepanel so authorization and export folder changes update immediately
- ✅ **Smart Export Persistence**: `.sd3`, `.csv`, and `_ExceptionReport.txt` exports save to the configured Drive folder (falling back to My Drive root)

### v2.2 - Meet Setup, Competition Types & Google Form Entry Generator
- ✅ **Competition-Type-Aware Meet Setup**: Supports `School`, `College`, `Age Group`, `Open`, and `Masters` meets, generating `Meet` (`AgeBrackets` for School/College), `Events`, `Schools` or `Teams`, `Individual Events Template`, and optional `Team Officials` / `How to fill entries` sample sheet
- ✅ **Google Form Entry Generator & Sync**: `formBuilder.js` / `form-setup-ui.html` creates and syncs a public Entry Form linked to `Form Entries`, with 1–9 event slots, optional DOB for School/College, optional write-in `"Other"` team, and duplicate-column cleanup
- ✅ **Export & Validation Parity**: Auto-detection of `Form Entries`, `Teams` vs `Schools` table lookup and exception labels, and duplicate-event warnings
- ✅ **Updated In-App Help & Sheet Preview**: Sidepanel Help updated to match Meet Setup / Entry Form / Relay Tables, plus **View sheet** preview in `Create Team Sheets`

### v2.1 - Enhanced Data Processing & Flexibility
- ✅ **Multiple entry format support**: Event names in headers OR in cells
- ✅ **Flexible column detection**: Works with or without "#" numbering column
- ✅ **Smart time parsing**: Handles MM:SS.SS, SS.SS, SS:DD, and other formats
- ✅ **Stroke-aware pool conversions**: Individual Medley uses correct distances
- ✅ **Extended pool support**: 33m, 18m, and Long Course (50m) conversions
- ✅ **Improved time formatting**: Proper decimal padding (.6 → .60, not .06)
- ✅ **Help & About pages**: In-app documentation with links to full guides

### v2.0 - SDIF Export & Exception Reporting
- ✅ Direct SDIF v3 export for Hy-Tek Meet Manager
- ✅ Automatic exception reporting with detailed validation
- ✅ Critical data validation (names, DOB, gender, team codes)
- ✅ Invalid event detection and reporting
- ✅ Shared data processing between CSV and SDIF exports
- ✅ Modernized codebase (ES6+ const/let)

### v1.0 - Core Functionality
- ✅ Template management and table creation
- ✅ Sheet duplication for multiple schools
- ✅ CSV export with auto-detection
- ✅ Named ranges and validation

## Roadmap

- [ ] Google Workspace Marketplace listing
- [ ] Relay entries SDIF export
- [ ] Custom event type configurations
- [ ] Bulk import from existing spreadsheets


## Support

- **Issues**: [GitHub Issues](https://github.com/petrfaitl/swim-entries-manager/issues)
- **Discussions**: [GitHub Discussions](https://github.com/petrfaitl/swim-entries-manager/discussions)
- **Email**: recorder@lvwasc.co.nz

## Legal

- **Privacy Policy**: [docs/PRIVACY_POLICY.md](docs/PRIVACY_POLICY.md)
- **Terms of Service**: [docs/TERMS_OF_SERVICE.md](docs/TERMS_OF_SERVICE.md)

## Acknowledgments

- Built with [Google Apps Script](https://developers.google.com/apps-script)
- Uses [TableApp Library](https://github.com/tanaikech/TableApp) for structured tables
- Icons from [Material Design](https://material.io/resources/icons/)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Donate

If you find this add-on useful, consider supporting its development:

☕ [Buy me a coffee](https://buymeacoffee.com/pfaitl)

---

**Note**: This add-on is currently in test deployment. A public release on the Google Workspace Marketplace is planned.
