Attribute VB_Name = "modLog"
Option Explicit

Public Sub WriteTextLine(ByVal FilePath As String, ByVal TextLine As String)
    Dim FileNo As Integer
    
    On Error GoTo ErrorHandler
    
    EnsureFolderExists GetLogsFolderPath()
    
    FileNo = FreeFile
    Open FilePath For Append As #FileNo
    Print #FileNo, TextLine
    Close #FileNo
    
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If FileNo > 0 Then Close #FileNo
End Sub
