Attribute VB_Name = "modConstants"
Option Explicit

'====================================================
' Application constants
'====================================================
Public Const APP_NAME As String = "Excel To SQL Import"
Public Const APP_LOG_FOLDER As String = "Logs"
Public Const APP_MAPPING_TABLE As String = "ExcelImportMappings"
Public Const APP_MAPPING_SCHEMA As String = "dbo"

'====================================================
' Authentication types
'====================================================
Public Const AUTH_WINDOWS As String = "Windows Authentication"
Public Const AUTH_SQL_SERVER As String = "SQL Server Authentication"

'====================================================
' Excel read modes
'====================================================
Public Const EXCEL_READ_MODE_ADO As Integer = 1
Public Const EXCEL_READ_MODE_AUTOMATION As Integer = 2

'====================================================
' UI / process constants
'====================================================
Public Const PROGRESS_MAX As Long = 100
Public Const DEFAULT_STAGE_PREFIX As String = "_ImportStage_"

'====================================================
' Message constants
'====================================================
Public Const MSG_CONFIRM_DELETE_EXISTING As String = _
    "All existing rows in the selected table will be deleted before import." & vbCrLf & vbCrLf & _
    "Do you want to continue?"

Public Const MSG_NO_DATABASE_SELECTED As String = "Please select a database."
Public Const MSG_NO_TABLE_SELECTED As String = "Please select a table."
Public Const MSG_NO_EXCEL_FILE_SELECTED As String = "Please select an Excel file."
Public Const MSG_NO_MAPPING_DEFINED As String = "Please define at least one valid column mapping."
Public Const MSG_OPERATION_CANCELLED As String = "Operation cancelled by user."

'====================================================
' SQL Server system databases
'====================================================
Public Const SYS_DB_MASTER As String = "master"
Public Const SYS_DB_MODEL As String = "model"
Public Const SYS_DB_MSDB As String = "msdb"
Public Const SYS_DB_TEMPDB As String = "tempdb"

