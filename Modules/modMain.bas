Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    On Error GoTo ErrorHandler

    Load frmImportExcelToSql
    frmImportExcelToSql.Show

    Exit Sub

ErrorHandler:
    MsgBox "Program startup error:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "Startup Error"
End Sub
