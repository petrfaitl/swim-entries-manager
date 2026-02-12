# Contributing to Swim Entries Manager

Thank you for your interest in contributing to Swim Entries Manager! This document provides guidelines and instructions for contributing to the project.

## Code of Conduct

### Our Standards

- Be respectful and inclusive
- Welcome newcomers and help them get started
- Focus on what is best for the community
- Show empathy towards other community members

## How to Contribute

### Reporting Bugs

Before creating bug reports, please check existing issues to avoid duplicates. When creating a bug report, include:

- **Clear title and description**
- **Steps to reproduce** the issue
- **Expected behavior** vs actual behavior
- **Screenshots** if applicable
- **Environment details** (Browser, Google Sheets version, etc.)

**Bug Report Template:**
```markdown
**Describe the bug**
A clear description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Go to '...'
2. Click on '....'
3. See error

**Expected behavior**
What you expected to happen.

**Screenshots**
If applicable, add screenshots.

**Environment:**
 - Browser: [e.g. Chrome, Firefox]
 - Add-on Version: [e.g. 1.0.0]
```

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion, include:

- **Clear title and description**
- **Use case** for the enhancement
- **Possible implementation** if you have ideas
- **Alternative solutions** you've considered

### Pull Requests

1. **Fork the repository** and create your branch from `main`
2. **Make your changes** following the code style guidelines
3. **Test your changes** thoroughly
4. **Update documentation** if needed
5. **Write clear commit messages**
6. **Submit a pull request**

## Development Setup

### Prerequisites

- Google Account
- [Node.js](https://nodejs.org/) (for clasp CLI)
- [clasp](https://github.com/google/clasp) installed globally

### Setup Steps

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/swim-entries-manager.git
   cd swim-entries-manager
   ```

2. **Install clasp:**
   ```bash
   npm install -g @google/clasp
   ```

3. **Login to clasp:**
   ```bash
   clasp login
   ```

4. **Create a new Apps Script project:**
   ```bash
   clasp create --type sheets --title "Swim Entries Manager Dev"
   ```

5. **Push the code:**
   ```bash
   clasp push
   ```

6. **Open in Apps Script editor:**
   ```bash
   clasp open
   ```

## Code Style Guidelines

### JavaScript

- Use **ES6+ syntax** where possible
- Use **const** and **let** instead of **var**
- Use **meaningful variable names**
- **Comment complex logic**
- Keep functions **focused and small**

### Naming Conventions

- **Functions**: `camelCase` (e.g., `createConfiguredTable`)
- **Private functions**: Suffix with `_` (e.g., `ensureStructuredTable_`)
- **Constants**: `UPPER_SNAKE_CASE` (e.g., `DEFAULT_CLUSTERS`)
- **Variables**: `camelCase` (e.g., `tableCfg`)

### Code Example

```javascript
/**
 * Creates a named range for a specific column in a table.
 * @param {string} namedRangeName - Name for the named range
 * @param {Sheet} sheet - The sheet containing the table
 * @param {number} startRow - Starting row (data row, not header)
 * @param {number} columnIndex - Column index (1-based)
 * @param {number} numRows - Number of data rows
 */
function createNamedRangeForColumn_(namedRangeName, sheet, startRow, columnIndex, numRows) {
  const ss = SpreadsheetApp.getActive();

  // Remove existing named range if it exists
  const existingRanges = ss.getNamedRanges();
  for (let i = 0; i < existingRanges.length; i++) {
    if (existingRanges[i].getName() === namedRangeName) {
      existingRanges[i].remove();
      Logger.log('[createNamedRangeForColumn_] Removed existing named range: %s', namedRangeName);
      break;
    }
  }

  // Create the range for the column (excluding header)
  const range = sheet.getRange(startRow, columnIndex, numRows, 1);
  ss.setNamedRange(namedRangeName, range);
  Logger.log('[createNamedRangeForColumn_] Created named range "%s" for %s', namedRangeName, range.getA1Notation());
}
```

### Documentation

- **JSDoc comments** for all public functions
- **Inline comments** for complex logic
- **README updates** for new features
- **User guide updates** for user-facing changes

## Project Structure

### File Organization

```
SwimEntriesManager/
├── appsscript.json          # Project configuration
├── ui.js                    # UI components (sidebar, dialogs)
├── tableConfig.js           # Table definitions
├── tableBuilder.js          # Table creation logic
├── duplicateTemplate.js     # Sheet duplication
├── exportToCSV.js          # CSV export
├── utils.js                # Utility functions
└── tests.js                # Test functions
```

### Module Responsibilities

- **ui.js**: All Card Service UI, dialogs, and menu items
- **tableConfig.js**: Table schemas, validation rules, styling
- **tableBuilder.js**: Table creation, named ranges, structured tables
- **duplicateTemplate.js**: Sheet duplication logic
- **exportToCSV.js**: CSV generation and export
- **utils.js**: Shared utility functions
- **tests.js**: Test cases and validation

## Testing

### Manual Testing

1. **Create a test spreadsheet**
2. **Deploy as test deployment**
3. **Test each feature**:
   - Create core tables
   - Verify validations work
   - Test named ranges
   - Duplicate sheets
   - Export to CSV

### Test Functions

Add test functions in `tests.js`:

```javascript
function testTableCreation() {
  const result = createConfiguredTable('Schools');
  Logger.log('Test result: %s', JSON.stringify(result));
  // Verify table was created correctly
}

function testNamedRanges() {
  const ss = SpreadsheetApp.getActive();
  const namedRanges = ss.getNamedRanges();
  Logger.log('Named ranges: %s', namedRanges.map(r => r.getName()));
  // Verify SchoolsList and EventsList exist
}
```

## Adding New Features

### Adding a New Table

1. **Define in tableConfig.js:**
   ```javascript
   MyNewTable: {
     title: 'My New Table',
     headers: ['Column1', 'Column2'],
     columns: {
       'Column1': { type: 'text' },
       'Column2': { type: 'number' }
     },
     options: {
       freezeHeader: 1,
       headerBg: '#356853',
       placement: { targetSheet: "MyNewTable", startCell: 'A1' }
     }
   }
   ```

2. **Test the table creation**
3. **Update documentation**

### Adding Named Ranges

Configure in table options:
```javascript
options: {
  // ... other options
  namedRange: { name: 'MyNamedRange', columnName: 'ColumnToReference' }
}
```

### Adding UI Elements

1. **Update ui.js** to add new sections or buttons
2. **Create handler functions**
3. **Test the UI flow**
4. **Update user guide**

## Commit Messages

Use clear, descriptive commit messages:

- **feat**: New feature (e.g., `feat: add relay entry table`)
- **fix**: Bug fix (e.g., `fix: resolve named range collision`)
- **docs**: Documentation (e.g., `docs: update user guide`)
- **style**: Code style (e.g., `style: format tableConfig.js`)
- **refactor**: Code refactoring (e.g., `refactor: extract validation logic`)
- **test**: Tests (e.g., `test: add CSV export tests`)
- **chore**: Maintenance (e.g., `chore: update dependencies`)

**Example:**
```
feat: add automatic time conversion for 33m pools

- Add checkbox column for pool conversion
- Implement conversion formula
- Update CSV export to handle converted times
```

## Release Process

1. **Update version** in `appsscript.json`
2. **Update CHANGELOG.md**
3. **Create release branch**: `release/v1.0.0`
4. **Test thoroughly**
5. **Merge to main**
6. **Tag release**: `git tag v1.0.0`
7. **Push tag**: `git push origin v1.0.0`
8. **Create GitHub release**

## Questions?

- **GitHub Discussions**: For questions and general discussion
- **GitHub Issues**: For bugs and feature requests
- **Email**: support@example.com

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to Swim Entries Manager! 🏊
