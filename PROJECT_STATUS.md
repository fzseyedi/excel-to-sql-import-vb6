# Excel To SQL Import

A VB6 desktop tool for importing Excel data into SQL Server tables with column mapping, staging, validation, transaction support, and detailed logging.

---

## Overview

This application helps users import data from Excel files into SQL Server tables in a controlled and safe way.

It supports:

- SQL Server connection with:
  - Windows Authentication
  - SQL Server Authentication
- Browsing non-system databases
- Browsing base tables only
- Searching databases and tables
- Reading Excel files using:
  - ADO/OLEDB
  - Excel Automation
- Excel header loading
- Excel preview rows
- Manual and automatic column mapping
- Mapping save/load
- Mapping validation
- Staging-table based import
- Real import into target table
- Transaction handling with commit/rollback
- Identity insert handling
- Duplicate row handling
- Progress bar and status updates
- Log file generation
- Friendly error messages

---

## Main Workflow

1. Connect to SQL Server
2. Select database and target table
3. Select Excel file
4. Load Excel columns
5. Preview sample Excel rows
6. Create column mappings
7. Validate mapping
8. Start import
9. Data is first inserted into a staging table
10. Data is transferred from staging to the target table inside a transaction
11. On success:
   - transaction is committed
   - staging table is removed
12. On failure:
   - transaction is rolled back
   - error is logged
   - friendly error message is shown

---

## Features

### Connection and Selection
- Supports Windows Authentication
- Supports SQL Server Authentication
- Automatically resets dependent UI state when connection settings change
- Shows only non-system databases
- Shows only base tables
- Database and table search is supported

### Excel Handling
- Supports `.xls` and `.xlsx`
- Can read Excel using ADO/OLEDB
- Can read Excel using Excel Automation
- Loads Excel headers from the first row
- Reads Excel data starting from row 2
- Provides preview of the first few rows

### Mapping
- Manual mapping between Excel columns and SQL columns
- Auto Match based on similar names
- Save mapping into SQL Server
- Load saved mapping
- Prevent duplicate mapping to the same SQL field
- Validate required target fields before import

### Import Engine
- Uses staging table before final import
- Handles data validation before target insert
- Supports `Delete Existing Rows`
- Supports `Continue on Type/Validation Error`
- Supports `Continue on Duplicate`
- Supports `IDENTITY_INSERT` when needed
- Uses transaction for final import
- Supports rollback on failure

### Logging and Errors
- Writes logs into `Logs` folder
- Preserves original error messages
- Shows friendly error messages to user
- Logs technical details for debugging
- Handles duplicate, FK, NULL, validation, and length errors more clearly

---

## Architecture Summary

The application is structured around the following main components:

- `clsSqlServerConnection`
- `clsDatabaseBrowser`
- `clsExcelReader`
- `clsMappingManager`
- `clsStagingManager`
- `clsImportEngine`
- `clsImportLogger`
- `clsImportOptions`
- `clsColumnInfo`
- `clsMappingItem`
- `clsImportRowResult`

UI is centered around:

- SQL connection
- Database/table selection
- Excel file selection
- Mapping tab
- Import options tab
- Execution/status tab

---

## Import Safety Design

This application was designed to reduce import risk:

- data is not inserted directly into the target table
- staging table is used first
- mapping must be validated before import
- final insert uses transaction
- rollback happens when import fails
- duplicate handling is configurable
- important operations are logged

---

## Supported Validation

Current validation includes:

- invalid numeric values
- invalid date values
- invalid bit values
- duplicate target insert errors
- foreign key related errors
- missing required values
- string length overflow
- numeric precision/scale overflow

---

## Current Project Status

This project has reached a **working and testable version**.

Implemented and tested areas include:

- SQL connection
- database and table browsing
- Excel reading
- preview rows
- mapping
- mapping save/load
- validation
- staging import
- final import
- transaction handling
- rollback
- duplicate handling
- progress display
- logging
- friendly error reporting

It can be considered a usable **Version 1.0** for controlled/internal use.

---

## Already Added Improvements Beyond Core 1.0

The following improvements were also added:

- Excel preview rows
- duplicate SQL-field mapping prevention
- stronger string-length validation
- better numeric precision/scale validation
- better NULL handling
- improved bit conversion
- clearer error propagation and logging

---

## Suggested Next Improvements

High-value future improvements:

1. Better final summary report
2. Separate row-level error log file
3. Save last-used connection settings
4. Optional cancel operation during long imports

Optional future enhancements:

- keep staging table for debugging
- better mapping colors/status indicators
- schema filtering
- deployment/setup checklist
- multilingual/resource-based messages

---

## Notes

Before deploying to other systems, verify that required dependencies are available, such as:

- VB6 runtime
- ADO/MDAC
- Common controls
- FlexGrid control
- Excel provider support where needed
- Excel installation if using Excel Automation mode

Also test with several real customer files, especially for:

- duplicate rows
- long text values
- required field issues
- identity columns
- foreign key dependent tables

---

## Logs

Application logs are written into the `Logs` folder under the application path.

Typical log contents include:

- import start
- selected target table
- selected Excel file
- row validation warnings/errors
- duplicate handling
- rollback/commit result
- final status

---

## Final Note

This project is no longer just a prototype UI.  
It includes a real import flow with validation, staging, transaction control, logging, and error handling, making it suitable as a strong internal desktop import tool built with VB6.