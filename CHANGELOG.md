# Changelog

All notable changes to this project will be documented in this file.

## [1.1.0] - 2026-04-24

### Added
- Excel preview rows
- duplicate SQL-field mapping prevention
- stronger string length validation
- better numeric precision/scale validation
- improved NULL handling
- improved bit conversion handling
- cancel support during import
- optional save/load of last connection settings
- search support for databases and tables

### Improved
- clearer friendly error messages
- better preservation of original SQL/validation errors
- better duplicate error reporting
- better staging and target import error logging
- more responsive cancel handling using more frequent `DoEvents`
- better reset behavior when connection settings change
- better reset behavior when file/table selection changes
- improved progress/status behavior during import

### Fixed
- duplicate key errors being replaced by generic VB6 errors
- cancellation errors being overwritten during Excel Automation error handling
- stale database/table selections after failed reconnection or authentication change
- bit values being written incorrectly as `True/False` instead of `1/0`
- long Excel preview reads loading unnecessary full-row data

## [1.0.0] - 2026-04-24

### Added
- SQL Server connection with Windows Authentication
- SQL Server connection with SQL Authentication
- browsing non-system databases
- browsing base tables
- Excel reading using ADO/OLEDB
- Excel reading using Excel Automation
- Excel header loading
- manual mapping
- auto match
- save/load mapping
- required mapping validation
- staging table import
- final import into target table
- transaction handling
- rollback support
- `IDENTITY_INSERT` handling
- duplicate handling
- progress bar and execution status
- logging to file
- friendly error handling foundation