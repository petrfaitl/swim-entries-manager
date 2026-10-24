# Core Specification: Swim Entries Manager

## Purpose

Swim Entries Manager is a Google Workspace Add-on for Google Sheets that helps swim meet organizers create structured entry templates, manage school/cluster registrations, and export swimmer entry data to SDIF v3 format for Hy-Tek Meet Manager or CSV for other processing tools.

## Capabilities

### 1. Core Table Management

The system shall provide a set of predefined table templates that can be created within a Google Spreadsheet. Each table includes headers, data validation rules, and named ranges for cross-sheet references.

**Tables:**
- **Meet**: Meet details (name, dates, course, venue, contact info)
- **Team Officials**: Team/school/cluster official details
- **Events for Individual Entries Template**: Event definitions (distance, discipline, auto-generated event names)
- **Schools for Individual Entries Template**: School/cluster registrations with team codes
- **Events for Relay Entries Template**: Relay event definitions
- **Schools for Relay Entries Template**: School/cluster registrations for relays

**Requirements:**
- The system shall support configurable school years, genders, countries, and regions via overrides
- The system shall apply data validation (list, number range, date, checkbox) per column configuration
- The system shall create named ranges for each table to enable cross-sheet references
- The system shall support two clear modes: `rebuild` (clear and recreate) and `append` (preserve existing data)
- The system shall use Google Sheets Tables (TableApp library) for advanced table features

### 2. Sheet Duplication from Templates

The system shall duplicate individual entry template sheets for each registered school/cluster, enabling bulk creation of entry forms.

**Requirements:**
- The system shall duplicate a single template sheet once per team/school from the Schools table
- The system shall automatically name each duplicated sheet based on the school/cluster name
- The system shall detect and prevent collisions with existing sheet names
- The system shall preserve dropdown validations referencing named ranges from core tables
- The system shall support both individual and relay entry templates

### 3. Relay Table Management

The system shall provide specialized relay entry table creation with support for multiple relay events and teams.

**Requirements:**
- The system shall generate relay entry tables based on configured relay events and schools
- The system shall support fixed swimmer rows per relay table (default: 4)
- The system shall maintain configurable gaps between tables when multiple relays share a sheet
- The system shall support suggested default relay event names (e.g., 4x25m Freestyle Relays)

### 4. SDIF Export

The system shall export swimmer entry data to SDIF v3 (.sd3) format compatible with Hy-Tek Meet Manager.

**Requirements:**
- The system shall read swimmer entries from duplicated template sheets
- The system shall validate all required fields (first name, last name, gender, date of birth, team code)
- The system shall generate detailed exception reports for data quality issues
- The system shall support two entry formats:
  - Method 1: Event names in cells (e.g., "Event 1" column header with event name in cell), the following column, after the event is an entry time column.
  - Method 2: Event names in headers (e.g., "100m Freestyle" header with time in cell)
- The system shall auto-detect column layout including optional "#" numbering columns
- The system shall parse multiple time formats (MM:SS.SS, SS.SS, SS:DD, etc.)
- The system shall apply stroke-aware pool conversions for non-standard pools:
  - 33m pool conversions with IM-specific distances (e.g., 133.2m vs 99.9m for 100m)
  - 18m pool conversions
  - Long Course (50m) to Short Course conversions
- The system shall use the RESIZESWIM function for time conversions
- The system shall create the .sd3 file in Google Drive with configurable output folder
- The system shall follow SDIF v3.0 specification with FINA organization code

### 5. CSV Export

The system shall export swimmer entry data to traditional CSV format for custom processing.

**Requirements:**
- The system shall export processed swimmer data with consistent column structure
- The system shall support the same entry format auto-detection as SDIF export
- The system shall create the CSV file in Google Drive

### 6. Data Processing & Validation

The system shall provide shared data processing logic for entry validation and transformation.

**Requirements:**
- The system shall capitalize swimmer names
- The system shall normalize school years
- The system shall validate team codes against the Schools table
- The system shall convert pool times using configured conversion tables
- The system shall detect and report critical data issues before export
- The system shall handle B-records, C-records, and D0/D1 records

### 7. User Interface

The system shall provide a Google Sheets Add-on interface with sidebar and dialog interactions.

**Requirements:**
- The system shall create an add-on menu under "Extensions" with actions for each capability
- The system shall provide a sidebar homepage with step-by-step workflow guidance
- The system shall use modal dialogs for table creation, sheet duplication, and export configuration
- The system shall display progress and exception reports during export operations
- The system shall support AuthMode.NONE fallback for limited permission states

## Technical Constraints

- Runtime: Google Apps Script V8 runtime
- Advanced Services: Google Sheets API v4, Google Drive API v3
- Timezone: Pacific/Auckland (configurable)
- OAuth Scopes: spreadsheets.currentonly, drive.file, userinfo.email, userinfo.profile, script.container.ui, script.locale
- The system shall use the TableApp library for table management
- The system shall log exceptions to Stackdriver (Cloud Logging)

## Entry Workflow

The add-on follows a structured four-step workflow:
1. **Create Core Tables** — Generate meet, events, schools, and team officials tables
2. **Populate Tables** — Enter meet details, events, and participating schools
3. **Create Sheets from Template** — Duplicate entry sheets per team/school
4. **Export** — Generate SDIF or CSV files from completed entry sheets
