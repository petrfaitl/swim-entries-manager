# Swim Entries Manager

[![Google Workspace Add-on](https://img.shields.io/badge/Google%20Workspace-Add--on-4285F4?logo=google)](https://workspace.google.com/)
[![Apps Script](https://img.shields.io/badge/Apps%20Script-V8-34A853)](https://developers.google.com/apps-script)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A Google Workspace Add-on for managing swim meet entries efficiently. This tool helps swim meet organizers create structured entry templates, manage school/cluster registrations, and export data in CSV format compatible with meet management software.

![Swim Entries Manager](https://www.lvwasc.co.nz/wp-content/uploads/2025/07/Swim-Entries-Manager.png)

## Features

### 🏊 Template Management
- **Create Core Tables**: Generate structured tables for events, schools, team officials, and relay entries
- **Automated Validation**: Built-in dropdown validations for consistency (events, school years, genders, distances)
- **Named Ranges**: Automatic creation of named ranges for cross-sheet references
- **Google Sheets Tables**: Uses the TableApp library for advanced table features

### 📋 Sheet Duplication
- **Bulk Sheet Creation**: Duplicate individual entry templates for multiple schools or clusters
- **Automatic Naming**: Sheets are automatically named based on school/cluster names
- **Collision Detection**: Prevents overwriting existing sheets

### 📊 CSV Export
- **Auto-Detection**: Automatically detects event and time columns in any sheet format
- **Flexible Format**: Handles sheets with or without time entries
- **Meet Manager Compatible**: Exports in format suitable for swim meet management software
- **Sheet Selection**: Export any visible sheet from your spreadsheet

## Quick Start

### Installation

#### Test Deployment (Current)
1. Open your Google Spreadsheet
2. Navigate to **Extensions** > **Apps Script**
3. Copy all files from this repository into your Apps Script project
4. Deploy as Test Deployment
5. Refresh your spreadsheet
6. Access via the sidebar or **Extensions** menu

#### From Google Workspace Marketplace (Coming Soon)
1. Visit the [Google Workspace Marketplace](#)
2. Search for "Swim Entries Manager"
3. Click "Install"
4. Grant necessary permissions
5. Open any Google Spreadsheet to use the add-on

### First Use

1. **Open the add-on**: Find it in the right sidebar (you may need to unhide the sidebar first).
2. **Create Core Tables (STEP 1)**:
   - Create all tables marked as required in `tableConfig.js` (`options.required: true`).
   - At minimum, include `Team Officials`, `Events for Individual Entries Template`, and `Schools for Individual Entries Template`.
   - If you are using the Individual Entries Template, set your `Gender` and `School Year` values before creating tables.
3. **Populate tables**:
   - `Team Officials`: Enter your team, school, or cluster details (minimum required data is the first column).
   - `Events for Template`: Enter `Distance` and `Discipline`; the last column builds the event name automatically.
   - `Schools`: Fill `Team Name`, `School`, `Cluster`, and `Code` based on your meet setup.
4. **Create sheets from template (STEP 2)**:
   - Verify the dropdowns in `INDIVIDUAL_EVENTS_TEMPLATE` first.
   - Use **Create Sheets From Template** to duplicate one entry sheet per team/school.
5. **Fill swimmer entries**:
   - Required: `First Name`, `Last Name`, `Gender`, `Date of Birth`, `School`, and selected events (`Event 1` to `Event 9`).
   - Recommended: entry times and school year.
6. **Export entries (STEP 3)**:
   - Export to CSV for meet software workflows.
   - CSV to SDIF conversion tool: https://www.lvwasc.co.nz/tools/csv-sdif-converter/

## Project Structure

```
SwimEntriesManager/
├── appsscript.json          # Project manifest and configuration
├── ui.js                    # User interface (sidebar, dialogs)
├── tableConfig.js           # Table definitions and configurations
├── tableBuilder.js          # Table creation and named range logic
├── duplicateTemplate.js     # Sheet duplication functionality
├── exportToCSV.js          # CSV export functionality
├── utils.js                # Utility functions
├── tests.js                # Test functions
└── README.md               # This file
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

- **Google Sheets API v4**: Advanced spreadsheet operations
- **TableApp Library**: Google's structured table functionality
  - Library ID: `1G4RVvyLtwPjQl6x_p8j3X65-yYVU3w2dMXxHDuzCgorucjs8P3Clv5Qt`
  - Version: 1

## Permissions

The add-on requires the following OAuth scopes:
- `spreadsheets.currentonly`: Read/write current spreadsheet
- `spreadsheets`: Full spreadsheet access
- `drive.file`: Create and manage files
- `drive`: Drive access for file operations
- `userinfo.email`: User identification
- `script.container.ui`: Display custom UI
- `script.locale`: Localization support

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

Run tests from the Apps Script editor:
```javascript
// In tests.js
testTableCreation();
testNamedRanges();
testCSVExport();
```

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for details.

### Development Workflow

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Roadmap

- [ ] Google Workspace Marketplace listing
- [ ] Advanced export options (SDIF format)
- [ ] Custom event type configurations


## Support

- **Issues**: [GitHub Issues](https://github.com/petrfaitl/swim-entries-manager/issues)
- **Discussions**: [GitHub Discussions](https://github.com/petrfaitl/swim-entries-manager/discussions)
- **Email**: recorder@lvwasc.co.nz

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
