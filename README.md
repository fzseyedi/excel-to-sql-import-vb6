# Excel To SQL Import

A VB6 desktop tool for importing Excel data into SQL Server tables with mapping, validation, staging, transaction support, and logging.

## Why this project?

Importing Excel data directly into production tables is risky.  
This project was built to make the process safer and more controlled by using:

- column mapping
- validation
- staging tables
- transaction-based final import
- rollback on failure
- progress tracking
- detailed logging

## Features

- SQL Server connection using:
  - Windows Authentication
  - SQL Server Authentication
- Browse non-system databases
- Browse base tables only
- Search databases and tables
- Read Excel files using:
  - ADO/OLEDB
  - Excel Automation
- Load Excel headers
- Preview first Excel rows
- Manual column mapping
- Auto Match for similar column names
- Save / Load mappings
- Validate required mappings
- Prevent duplicate mapping to the same SQL field
- Import through a staging table
- Transaction-based final import
- `IDENTITY_INSERT` support when needed
- Duplicate handling
- Progress bar and status updates
- Log file generation
- Friendly error reporting

## Import flow

1. Connect to SQL Server
2. Select database and target table
3. Select Excel file
4. Load Excel columns
5. Preview Excel rows
6. Create mappings
7. Validate mapping
8. Start import
9. Load rows into a staging table
10. Move rows from staging to target inside a transaction
11. Commit on success
12. Rollback on failure

## Safety design

This application does **not** insert Excel rows directly into the target table.

Instead, it uses:

- staging table first
- validation before final insert
- transaction for final import
- rollback on failure
- logging for troubleshooting

## Current status

This project is currently in a **working and testable state**.

Implemented and tested areas include:

- connection and authentication
- database and table browsing
- Excel loading
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
- improved error handling

## Tech stack

- Visual Basic 6
- SQL Server
- ADO
- MSFlexGrid
- Common Controls
- Optional Excel Automation

## Requirements

Before running on another machine, verify that the following are available:

- VB6 runtime
- ADO/MDAC
- required OCX controls
- Excel provider support for ADO/OLEDB mode
- Microsoft Excel for Excel Automation mode

## Logs

The application writes logs into the `Logs` folder under the application path.

## Roadmap

Planned improvements include:

- richer final import summary
- separate row-level error log
- save last-used connection settings
- cancel long-running import
- optional keep-staging-table mode
- better mapping visuals
- deployment/setup checklist

## Notes

This project is intended as a practical desktop import tool, not just a UI prototype.  
It already includes a real import pipeline with validation, staging, transaction control, logging, and error handling.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.