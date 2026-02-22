# Swim Entries Manager - User Guide

Welcome to the Swim Entries Manager user guide! This comprehensive guide will help you set up and use the add-on to manage your swim meet entries efficiently.

## Table of Contents

1. [Getting Started](#getting-started)
2. [Creating Core Tables](#creating-core-tables)
3. [Populating Your Data](#populating-your-data)
4. [Creating Entry Sheets](#creating-entry-sheets)
5. [Exporting for Meet Manager](#exporting-for-meet-manager)
6. [Understanding Exception Reports](#understanding-exception-reports)
7. [Advanced Configuration](#advanced-configuration)
8. [Troubleshooting](#troubleshooting)
9. [FAQ](#faq)

## Getting Started

### Installation

The add-on is currently available as a test deployment. To use it:

1. Open your Google Spreadsheet where you want to manage swim entries
2. Navigate to **Extensions** > **Apps Script**
3. Copy all project files from the repository
4. Save and deploy as a test deployment
5. Refresh your spreadsheet
6. Access the add-on from **Extensions** > **Swim Entries Manager**

**Alternative Installation**
Get an email invite as a tester.

### First Launch

When you first open the add-on:

1. Click on the sidebar and find the icon, as one of the installed apps. If you do not see the sidebar, unhide/reopen it, 
2. The sidebar will open with two main cards:
   - **Swim Meet Templates**: Create tables and entry sheets
   - **Export Entries**: Export data to CSV
3. Upon further use, click on **Extensions** > **Swim Entries Manager** in the menu.

### Quick Workflow

Use this sequence for a standard meet setup:

1. **Create Core Tables (STEP 1)**
   - Create all tables marked as required.
   - At minimum, create `Team Officials`, `Events for Individual Entries Template`, and `Schools for Individual Entries Template`.
   - If using the Individual Entries Template, define `Gender` and `School Year` values.
2. **Populate tables**
   - `Team Officials`: Enter teams/schools/clusters (minimum required value is the first column).
   - `EventsForTemplate`: Enter `Distance` and `Discipline`; `Events` is auto-generated.
   - `SchoolsForTemplate`: Fill `Team Name`, `School`, `Cluster`, and `Code`.
3. **Verify template dropdowns**
   - Confirm that the `INDIVIDUAL_EVENTS_TEMPLATE` dropdowns are populated before duplication.
4. **Create Sheets from Templates (STEP 2)**
   - Duplicate the individual template for each team/school so each coordinator works in a separate sheet.
5. **Fill swimmer entries**
   - **Required fields**: `First Name`, `Last Name`, `Gender`, `Date of Birth`, and valid `Team Code`
   - **Recommended fields**: entry times and `School Year`
   - Ensure each team exists in the Schools table
6. **Export for Meet Manager (STEP 4)**
   - Use "Export for Meet Manager" to generate an SDIF (.sd3) file
   - Review any exception reports if warnings appear
   - Alternative: Export to CSV for manual processing

## Creating Core Tables

### Overview

The core tables form the foundation of your swim meet management system. These include:

- **Team Officials**: Contact information for team managers and officials
- **Detailed Events**: Complete event list for Meet Manager alignment and participating schools info.
- **Events for Template**: Simplified event names for dropdown validation
- **Schools**: School information with clusters for dropdown validation
- **Individual Events Template**: Template for per-school entry sheets
- **Relay Entry**: Relay team configuration

### Step-by-Step: Creating Tables

1. **Open the Templates Card**
   - In the sidebar, you'll see the "Swim Meet Templates" card
   - Click on **STEP 1: Create Core Tables**

2. **Select Tables to Create**
   - Click the **"Create Core Tables"** button
   - A dialog will appear with checkboxes for each available table
   - Select the tables you need (typically start with EventsForTemplate and Schools)

3. **Customize Dropdown Values (Optional)**
   - **School Years**: Comma-separated list (default: Y5,Y6,Y7,Y8,Y9,Junior,Intermediate,Senior)
   - **Genders**: Comma-separated list (default: Female,Male)
   - These values will populate dropdown validations in your tables

4. **Placement Options (Optional)**
   - **Override Sheet Name**: Specify a custom sheet name for all selected tables
   - **Start Cell**: Specify where the table should start (default: A1)

5. **Create the Tables**
   - Click **"Create Tables"**
   - The system will create structured tables with:
     - Formatted headers
     - Data validation dropdowns
     - Named ranges (where applicable)
     - Frozen header rows

6. **Verify Creation**
   - Check that new sheets were created
   - Verify that dropdowns work in validation columns
   - Named ranges should be visible in Data > Named ranges

### Understanding Named Ranges

Named ranges enable cross-sheet references:

- **SchoolsList**: References the 'School' column in the Schools table
- **EventsList**: References the 'Events' column in the EventsForTemplate table

These allow the Individual Events Template to always reference current school and event lists.

## Populating Your Data

### Schools Table

Located in the **SchoolsForTemplate** sheet:

| Team Name | School | Cluster | Code  |
|-----------|--------|---------|-------|
| Hamilton East | Hamilton East School | Cluster One | HAMES |
| Cambridge | Cambridge Primary | Cluster Two | CAMPS |

**Tips:**
- The **School** column is used for the named range
- Codes for Secondary schools are predefined. For other grades it is advisable to seek confirmation from regional leads.
- Codes are 4-5 letters
- Use short codes for quick identification in exports
- For cluster meets, values should be unique across the four fields where possible
- Repeated values can still be used when needed to distinguish entries from different schools in cluster competitions
- For single-school meets, the **School** column can be used for houses; **Cluster** can be left blank

### Events for Template Table

Located in the **EventsForTemplate** sheet:

| Distance | Discipline | Events |
|----------|------------|--------|
| 50m | Freestyle | 50m Freestyle |
| 100m | Backstroke | 100m Backstroke |

**Tips:**
- The **Events** column uses a formula to combine Distance + Discipline
- Formula: `=IF(AND(B2<>"",A2<>""),A2&" "&B2,"")`
- This column feeds the EventsList named range

### Team Officials Table

Located in the **Team Officials** sheet:

| Cluster or School | Sport Coordinator | Email | Contact Phone | ... |
|-------------------|-------------------|-------|---------------|-----|
| Cluster One | John Smith | john@example.com | 021-123-4567 | ... |

**Tips:**
- Fill in all contact information for each cluster/school
- This information can be used for communications

### Detailed Events Table

Located in the **Events** sheet:

| Events | Gender | Min Age | Max Age | Discipline | Distance | Event Type | Event No. |
|--------|--------|---------|---------|------------|----------|------------|-----------|
| Girls 12-13 50m Freestyle | Female | 12 | 13 | Freestyle | 50 | Individual | 1 |

**Tips:**
- This is your complete event list for Meet Manager
- Event numbers should match your meet program
- Use consistent naming conventions

## Creating Entry Sheets

### Overview

Once your core tables are populated, you can create individual entry sheets for each school or cluster.

### Step-by-Step: Duplicating Templates

1. **Navigate to STEP 2**
   - In the sidebar, find **STEP 2: Create Sheets from Templates**
   - Click the **"Create Sheets from Templates"** button

2. **Configure Duplication**
   - A dialog will appear
   - **Template Sheet**: Select "INDIVIDUAL_EVENTS_TEMPLATE"
   - **Source Sheet**: Select "SchoolsForTemplate" (where your schools are listed)
   - **Range**: Specify the range containing school names (e.g., "B2:B21")
     - This should reference the School column

3. **Create Sheets**
   - Click **"Duplicate Templates"**
   - The system will:
     - Create a new sheet for each school
     - Name each sheet after the school
     - Skip any sheets that already exist
   - You'll see a confirmation message with the count of sheets created

4. **Verify Creation**
   - Check the sheet tabs at the bottom
   - Each school should have its own sheet
   - All sheets should maintain the validation rules

### Using Entry Sheets

Each entry sheet contains:

| # | First Name | Last Name | Date of Birth | Gender | School Year | School | Event 1 | Time 1 | Event 2 | Time 2 | ... |
|---|------------|-----------|---------------|--------|-------------|--------|---------|--------|---------|--------|-----|
| 1 | Sarah | Johnson | 15/01/2012    | Female | Y8 | Hamilton East | 50m Freestyle | 0:35.2 | 100m Backstroke | 1:15.8 | ... |

**Features:**
- **Dropdowns**: School, School Year, Gender, and all Event columns have dropdown validation
- **Time Format**: Use m:s.S format (e.g., 0:35.2 for 35.2 seconds, 1:32.12 for 1 minute and 32.12 seconds)
- **33m Pool Conversion**: Optional checkbox to indicate times from 33m pool

**Tips:**
- Fill in swimmer details first (name, DOB, gender, year)
- Select school from dropdown
- Choose events from validated dropdowns
- Enter times in the correct format
- Leave time cells blank if no entry time available
- 33m conversion uses simple distance truncation. Provides approximate conversion for schools qualifying for 25m pool competitions, when coming from 33m pool.

## Exporting for Meet Manager

### Overview

The Swim Entries Manager provides two export options:

1. **Export for Meet Manager** (Recommended) - Generates an SDIF (.sd3) file ready for direct import into Hy-Tek Meet Manager
2. **Export to CSV** - Creates a CSV file for manual processing or use with other tools

### Recommended: Export for Meet Manager (SDIF)

This is the **primary export method** that generates SDIF v3 format files compatible with Hy-Tek Meet Manager.

#### Step-by-Step Process

1. **Navigate to Export**
   - In the sidebar, click **"Go to Export →"** button
   - Or use **Extensions** > **Swim Entries Manager** > **Export for Meet Manager**

2. **Open Export Dialog**
   - Click **"Export for Meet Manager"** button in the sidebar (STEP 4)
   - A dialog will appear

3. **Configure Export**
   - **Entries Sheet**: Select the sheet containing swimmer entries
   - **Meet Table Name**: Confirm the meet details table (default: "Meet")
   - **Schools Table Name**: Confirm the schools/teams table (default: "Schools")
   - **Output File Name**: Leave blank for auto-generated name, or specify custom name

4. **Generate SDIF File**
   - Click **"Generate SD3 File"**
   - The system will:
     - Validate all critical data (names, DOB, gender, team codes)
     - Process events and times
     - Generate SDIF v3 format file
     - Create an exception report if any issues are found

5. **Review Results**
   - **Success (Green)**: "Successfully generated!" - No issues found
   - **Warning (Yellow)**: "Generated with warnings - please review exception report"
     - Click the exception report link to see details
     - The .sd3 file is still generated but may be missing some data

6. **Download Files**
   - Click **"Download .sd3 File"** to get your Meet Manager import file
   - If warnings appeared, also download and review the Exception Report

#### What Gets Validated

The export process performs critical validations to prevent import failures:

**Critical Data (Must be present or swimmer is excluded):**
- First Name
- Last Name
- Date of Birth
- Gender
- Team Code (must exist in Schools table)
- Team Name (from Schools table)

**Event Validation:**
- Events must be in correct format (e.g., "25m Freestyle", not "Y5")
- Invalid events are skipped but swimmer is still exported

**Warnings (Not critical):**
- Swimmers with no events entered

### Alternative: Export to CSV

Use this option if you need CSV format for custom processing or other tools.

#### Step-by-Step Process

1. **Navigate to Export**
   - In the sidebar, expand **"Alternative: Export to CSV"** section
   - Or use **Extensions** > **Swim Entries Manager** > **Export to CSV**

2. **Open Export Dialog**
   - Click **"Export to CSV"** button
   - A dialog will appear

3. **Select Sheet and Table**
   - **Sheet**: Choose the sheet to export
   - **Team Code Table**: Select the table for team code lookup (default: "Schools")

4. **Export**
   - Click **"Export to CSV"**
   - The system will:
     - Auto-detect Event and Time columns
     - Format times correctly
     - Look up team codes
     - Generate CSV file

5. **Download CSV**
   - Click the download link
   - Save with a descriptive name (e.g., "Hamilton_East_Entries.csv")

#### CSV Format

```csv
Team Code,First Name,Last Name,Date of Birth,Gender,Event Entries,Entry Times,School Year
HAMES,Sarah,Johnson,15/01/2012,Female,"50m Freestyle, 100m Backstroke","0:35.20, 1:15.80",Y8
```

**Features:**
- Auto-detects event/time column pairs
- Handles missing times gracefully
- Provides approximate conversion for 33m pool times
- Team codes are 4-5 letters

## Understanding Exception Reports

### What is an Exception Report?

When you export for Meet Manager, the system validates all critical data required for successful import. If any issues are found, an **Exception Report** is automatically generated alongside your .sd3 file.

The exception report is a text file that details:
- What data is missing or invalid
- Which swimmers or events are affected
- Whether items were excluded or included with warnings
- Action items to fix the issues

### When Exception Reports Are Created

An exception report is generated whenever:
- A swimmer is missing critical data (name, DOB, gender, team code)
- A team code isn't found in the Schools table
- Events are in invalid format
- Swimmers have no events entered

### Reading an Exception Report

#### Report Structure

```
═══════════════════════════════════════════════════════════════
         SDIF EXPORT EXCEPTION REPORT
═══════════════════════════════════════════════════════════════

Generated: 2026-02-22 15:30:00
Meet: Summer Championships

───────────────────────────────────────────────────────────────
SUMMARY
───────────────────────────────────────────────────────────────

Total swimmers processed: 45
Swimmers excluded (missing critical data): 3
Valid events exported: 120
Skipped events: 8
Total issues found: 11
```

The **Summary** section shows:
- How many swimmers were successfully exported
- How many were excluded due to missing critical data
- Event statistics
- Total number of issues found

#### Error Categories

Errors are grouped by category for easy review:

**1. Missing Critical Data**
```
▸ Missing Critical Data (3)

  Swimmer: John [No Last Name]
  Team: ABC
  Missing Field: Missing Last Name
  ⚠ This swimmer was EXCLUDED from the export
```

**Impact**: These swimmers will **NOT appear** in Meet Manager. You must fix these issues and re-export.

**2. Missing Team in Schools Table**
```
▸ Missing Team in Schools Table (2)

  Team Code: XYZ
  Missing Field: Team Code and Team Name
  Issue: Team "XYZ" not found in Schools table...
```

**Impact**: Team information may be incomplete. Swimmers are included but team details may not display correctly.

**3. Invalid Event Format**
```
▸ Invalid Event Format (6)

  Swimmer: Jane Smith
  Team: ABC
  Event: Y5
  Reason: Unrecognised event format
```

**Impact**: Invalid events are **NOT included** in the export. Swimmer is still exported but without these events.

#### Warning Categories

**No Events**
```
▸ No Events (2)

  Swimmer: Tom Brown
  Team: ABC
  Issue: Swimmer has no events entered
```

**Impact**: Swimmer is included in export but has no events. Not critical but may need attention.

#### Action Required Section

The report ends with a clear summary of actions needed:

```
───────────────────────────────────────────────────────────────
ACTION REQUIRED
───────────────────────────────────────────────────────────────

🚨 CRITICAL ERRORS FOUND:

  • 3 swimmer(s) EXCLUDED due to missing critical data
    (First Name, Last Name, Date of Birth, Gender, or Team Code)
    These swimmers will NOT appear in Meet Manager.

  • 2 team(s) missing from Schools table
    Team information may be incomplete in Meet Manager.

  • 6 event(s) NOT included due to invalid format
    (e.g., "Y5" instead of "25m Freestyle")
    Please correct event names and re-export.

  ⚠️  ACTION REQUIRED: Fix the errors above and re-export.
```

### Fixing Common Issues

#### Missing Critical Data

**Problem**: Swimmer excluded due to missing name, DOB, gender, or team code

**Solution**:
1. Open the entries sheet
2. Locate the swimmer mentioned in the report
3. Fill in all required fields:
   - First Name
   - Last Name
   - Date of Birth (DD/MM/YYYY format)
   - Gender (use dropdown)
   - School/Team (use dropdown)
4. Ensure the team exists in Schools table
5. Re-export

#### Missing Team in Schools Table

**Problem**: Team code not found or incomplete in Schools table

**Solution**:
1. Open the **Schools** (or **SchoolsForTemplate**) table
2. Find or add the team row
3. Ensure these fields are filled:
   - **Code**: 4-5 letter team code
   - **Team Name**: Full team name
   - **School**: School name
4. Save and re-export

#### Invalid Event Format

**Problem**: Event name doesn't match expected format

**Solution**:
1. Events must be formatted as: `[Distance] [Stroke]`
   - ✅ Correct: "25m Freestyle", "50m Backstroke", "100m IM"
   - ❌ Incorrect: "Y5", "25 Free", "50 Back"
2. Open the entries sheet
3. Correct the event name using the dropdown
4. Re-export

#### Swimmers with No Events

**Problem**: Swimmer has no events entered (warning only)

**Action**:
- If intentional, no action needed
- If swimmer should have events, add them and re-export

### Best Practices

1. **Review Before Import**: Always review exception reports before importing to Meet Manager
2. **Fix Critical Errors First**: Focus on excluded swimmers and missing teams first
3. **Validate Data Early**: Run a test export early to catch data quality issues
4. **Keep Schools Table Updated**: Ensure all teams are in the Schools table before entries begin
5. **Use Consistent Format**: Train data entry users on proper event naming
6. **Save Reports**: Keep exception reports for your records

### Tips for Clean Exports

- ✅ **DO**: Fill in all required fields before exporting
- ✅ **DO**: Use dropdown values when available
- ✅ **DO**: Add all teams to Schools table first
- ✅ **DO**: Use correct event naming format
- ❌ **DON'T**: Leave required fields blank
- ❌ **DON'T**: Type freeform text when dropdowns are available
- ❌ **DON'T**: Use abbreviated event names

## Advanced Configuration

### Customizing Table Definitions

Tables are defined in `tableConfig.js`. You can customize:

#### Column Validation

```javascript
'School Year': {
  type: 'text',
  validation: {
    type: 'list',
    args: { values: ['Y5','Y6','Y7','Y8'] }
  }
}
```

#### Named Ranges

```javascript
options: {
  namedRange: { name: 'MyNamedRange', columnName: 'ColumnName' }
}
```

#### Table Styling

```javascript
options: {
  freezeHeader: 1,
  headerBg: '#356853',
  rows: 30,
  placement: { targetSheet: "SheetName", startCell: 'A1' }
}
```

### Adding Custom Clusters

Edit the `DEFAULT_CLUSTERS` array in `tableConfig.js`:

```javascript
const DEFAULT_CLUSTERS = [
  'North Region',
  'South Region',
  'East Region',
  'West Region'
];
```

Or use the WBOP clusters by changing line 131:

```javascript
'Cluster': {
  type: 'text',
  validation: {
    type: 'list',
    args: { values: WBOP_CLUSTERS }  // Changed from DEFAULT_CLUSTERS
  }
}
```

### Custom Event Distances

Modify the EventsForTemplate configuration:

```javascript
'Distance': {
  type: 'text',
  validation: {
    type: 'list',
    args: { values: ['25m','50m','100m','200m','400m','800m','1500m'] }
  }
}
```

## Troubleshooting

### Tables Not Creating

**Problem**: Tables fail to create or show error messages

**Solutions**:
1. Check that TableApp library is properly configured
2. Verify you have edit permissions on the spreadsheet
3. Check for name conflicts with existing sheets
4. Review the Apps Script logs (View > Logs)

### Dropdowns Not Working

**Problem**: Validation dropdowns show no options or errors

**Solutions**:
1. Verify named ranges exist (Data > Named ranges)
2. Ensure source tables (Schools, EventsForTemplate) are populated
3. Recreate the named ranges by recreating the source tables
4. Check that the named range formulas reference the correct columns

### Sheet Duplication Fails

**Problem**: Template sheets don't duplicate or show errors

**Solutions**:
1. Verify the template sheet name is correct
2. Check that the source range contains valid data
3. Ensure no special characters in school names
4. Check for sheet name length limits (max 100 characters)

### CSV Export Issues

**Problem**: CSV doesn't generate or format incorrectly

**Solutions**:
1. Verify the sheet has data
2. Check that event/time columns follow the naming pattern
3. Ensure times are in the correct format (m:s.S)
4. Try exporting a different sheet to isolate the issue

### Permission Errors

**Problem**: "Authorization required" or permission errors

**Solutions**:
1. Re-authorize the add-on from Extensions menu
2. Check OAuth scopes in appsscript.json
3. Verify you have necessary Google account permissions
4. Try removing and reinstalling the add-on

## FAQ

### Can I customize the event types?

Yes! Edit `tableConfig.js` to modify the EventsForTemplate table configuration. You can add or remove distances, disciplines, and event types.

### How many entries can I manage?

Google Sheets supports up to 10 million cells. For practical purposes, you can manage hundreds of swimmers across dozens of schools without issues.

### Can I import existing data?

Yes! Copy and paste your existing data into the appropriate tables. Make sure column headers match exactly.

### Does it work offline?

No, this is a Google Workspace Add-on and requires an internet connection to function.

### Can multiple users collaborate?

Yes! Multiple users can work on the same spreadsheet simultaneously, just like any Google Sheet. Use Google Sheets' sharing features to control access.

### How do I backup my data?

Use File > Make a copy to create backups. You can also export to Excel format via File > Download.

### Can I use this for other sports?

With modifications, yes! The structure is flexible enough to adapt to other sports with similar entry requirements. You'd need to customize the table configurations.

### Is my data private?

Yes. The add-on only accesses spreadsheets you explicitly authorize. Data is stored in your Google Drive account and follows Google's security policies.

### How do I report bugs?

Please report bugs via [GitHub Issues](https://github.com/yourusername/swim-entries-manager/issues) with:
- Clear description of the issue
- Steps to reproduce
- Expected vs actual behavior
- Screenshots if applicable

### How can I contribute?

See [CONTRIBUTING.md](../CONTRIBUTING.md) for guidelines on contributing to the project.

## Support

Need help? Here are your options:

- **Documentation**: You're reading it! Check other sections for specific topics
- **GitHub Issues**: [Report bugs or request features](https://github.com/yourusername/swim-entries-manager/issues)
- **GitHub Discussions**: [Ask questions or discuss usage](https://github.com/yourusername/swim-entries-manager/discussions)
- **Email Support**: recorder@lvwasc.co.nz

## Tips for Success

1. **Start Small**: Begin with EventsForTemplate and Schools tables
2. **Populate Data First**: Fill in your schools and events before creating entry sheets
3. **Use Consistent Naming**: Keep school names and event names consistent
4. **Test Export Early**: Export a test sheet early to verify format compatibility
5. **Train Your Team**: Share this guide with others who will use the system

## Next Steps

Now that you understand the basics:

1. Create your core tables (Schools and EventsForTemplate)
2. Populate them with your meet-specific data
3. Create entry sheets for each school
4. Distribute to team managers for completion
5. Export completed sheets to CSV
6. Import into your meet management software

Happy swimming! 🏊
