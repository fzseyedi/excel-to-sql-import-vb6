VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExcelReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mReadMode As Integer

Public Property Get ReadMode() As Integer
    ReadMode = mReadMode
End Property

Public Property Let ReadMode(ByVal Value As Integer)
    mReadMode = Value
End Property

Public Sub ValidateExcelFile(ByVal FilePath As String)
    If IsNullOrEmpty(FilePath) Then
        Err.Raise vbObjectError + 1000, "clsExcelReader.ValidateExcelFile", "Excel file path is empty."
    End If
    
    If Not FileExists(FilePath) Then
        Err.Raise vbObjectError + 1001, "clsExcelReader.ValidateExcelFile", "Selected Excel file does not exist."
    End If
    
    If Not IsExcelFileExtensionValid(FilePath) Then
        Err.Raise vbObjectError + 1002, "clsExcelReader.ValidateExcelFile", "Only .xls and .xlsx files are supported."
    End If
End Sub

Public Function LoadHeaders(ByVal FilePath As String) As Collection
    Call ValidateExcelFile(FilePath)
    
    Select Case mReadMode
        Case EXCEL_READ_MODE_AUTOMATION
            Set LoadHeaders = LoadHeadersByAutomation(FilePath)
        Case Else
            Set LoadHeaders = LoadHeadersByAdo(FilePath)
    End Select
End Function

Private Function LoadHeadersByAdo(ByVal FilePath As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim Result As Collection
    Dim SheetName As String
    Dim i As Long
    
    Set Result = New Collection
    Set Conn = New ADODB.Connection
    
    Conn.Open GetExcelConnectionString(FilePath)
    
    SheetName = GetFirstWorksheetName(Conn)
    If Len(Trim$(SheetName)) = 0 Then
        Err.Raise vbObjectError + 1010, "clsExcelReader.LoadHeadersByAdo", "No worksheet could be found in the Excel file."
    End If
    
    'On Error GoTo 0
    Set Rs = New ADODB.Recordset
    Rs.Open "SELECT * FROM " & BracketName(SheetName) & " WHERE 1=0", Conn, adOpenForwardOnly, adLockReadOnly
    
    For i = 0 To Rs.Fields.Count - 1
        Result.Add Trim$(Rs.Fields(i).Name)
    Next i
    
    Rs.Close
    Set Rs = Nothing
    
    Conn.Close
    Set Conn = Nothing
    
    Set LoadHeadersByAdo = Result
    Exit Function

ErrorHandler:
    On Error Resume Next
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    If Not Conn Is Nothing Then
        If Conn.State = adStateOpen Then Conn.Close
    End If
    Set Conn = Nothing
    
    Err.Raise Err.Number, "clsExcelReader.LoadHeadersByAdo", Err.Description
End Function

Private Function LoadHeadersByAutomation(ByVal FilePath As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim ExcelApp As Object
    Dim WorkBook As Object
    Dim WorkSheet As Object
    Dim Result As Collection
    Dim LastCol As Long
    Dim i As Long
    Dim HeaderText As String
    
    Const xlToLeft As Long = -4159
    
    Set Result = New Collection
    
    Set ExcelApp = CreateObject("Excel.Application")
    
    '''Set WorkBook = ExcelApp.Workbooks.Open(FilePath, False, True)
    Set WorkBook = ExcelApp.Workbooks.Open(FilePath)
    Set WorkSheet = WorkBook.Worksheets(1)
    
    LastCol = WorkSheet.Cells(1, WorkSheet.Columns.Count).End(xlToLeft).Column
    
    If LastCol < 1 Then
        Err.Raise vbObjectError + 1020, "clsExcelReader.LoadHeadersByAutomation", "No header row could be detected in the Excel file."
    End If
    
    For i = 1 To LastCol
        HeaderText = Trim$(CStr(WorkSheet.Cells(1, i).Value))
        
        If Len(HeaderText) = 0 Then
            HeaderText = "Column" & CStr(i)
        End If
        
        Result.Add HeaderText
    Next i
    
    WorkBook.Close False
    ExcelApp.Quit
    
    Set WorkSheet = Nothing
    Set WorkBook = Nothing
    Set ExcelApp = Nothing
    
    Set LoadHeadersByAutomation = Result
    Exit Function

ErrorHandler:
    On Error Resume Next
    
    If Not WorkBook Is Nothing Then
        WorkBook.Close False
    End If
    
    If Not ExcelApp Is Nothing Then
        ExcelApp.Quit
    End If
    
    Set WorkSheet = Nothing
    Set WorkBook = Nothing
    Set ExcelApp = Nothing
    
    Err.Raise Err.Number, "clsExcelReader.LoadHeadersByAutomation", Err.Description
End Function

Private Function GetExcelConnectionString(ByVal FilePath As String) As String
    Dim Extension As String
    
    Extension = GetFileExtension(FilePath)
    
    Select Case Extension
        Case ".xls"
            GetExcelConnectionString = _
                "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & FilePath & ";" & _
                "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
                
        Case ".xlsx"
            GetExcelConnectionString = _
                "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & FilePath & ";" & _
                "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"";"
                
        Case Else
            Err.Raise vbObjectError + 1030, "clsExcelReader.GetExcelConnectionString", "Unsupported Excel file extension."
    End Select
End Function

Private Function GetFirstWorksheetName(ByVal Conn As ADODB.Connection) As String
    On Error GoTo ErrorHandler
    
    Dim Rs As ADODB.Recordset
    Dim TableName As String
    
    Set Rs = Conn.OpenSchema(adSchemaTables)
    
    Do While Not Rs.EOF
        TableName = NzString(Rs.Fields("TABLE_NAME").Value)
        
        If IsWorksheetTableName(TableName) Then
            GetFirstWorksheetName = NormalizeWorksheetName(TableName)
            Exit Do
        End If
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    Err.Raise Err.Number, "clsExcelReader.GetFirstWorksheetName", Err.Description
End Function

Private Function IsWorksheetTableName(ByVal TableName As String) As Boolean
    Dim Value As String
    
    Value = UCase$(Trim$(TableName))
    
    If InStr(1, Value, "$", vbTextCompare) = 0 Then
        IsWorksheetTableName = False
        Exit Function
    End If
    
    If InStr(1, Value, "FILTERDATABASE", vbTextCompare) > 0 Then
        IsWorksheetTableName = False
        Exit Function
    End If
    
    If InStr(1, Value, "PRINT_AREA", vbTextCompare) > 0 Then
        IsWorksheetTableName = False
        Exit Function
    End If
    
    IsWorksheetTableName = True
End Function

Private Function NormalizeWorksheetName(ByVal TableName As String) As String
    Dim Result As String
    
    Result = Trim$(TableName)
    
    If Left$(Result, 1) = "'" And Right$(Result, 1) = "'" Then
        Result = Mid$(Result, 2, Len(Result) - 2)
    End If
    
    NormalizeWorksheetName = Result
End Function

Public Function LoadRows(ByVal FilePath As String, Optional ByVal OwnerForm As Object = Nothing, Optional ByVal MaxRows As Long = 0) As Collection
    Call ValidateExcelFile(FilePath)
    
    Select Case mReadMode
        Case EXCEL_READ_MODE_AUTOMATION
            Set LoadRows = LoadRowsByAutomation(FilePath, OwnerForm, MaxRows)
        Case Else
            Set LoadRows = LoadRowsByAdo(FilePath, MaxRows)
    End Select
End Function

Private Function LoadRowsByAdo(ByVal FilePath As String, Optional ByVal MaxRows As Long = 0) As Collection
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim Result As Collection
    Dim RowData As Object
    Dim SheetName As String
    Dim i As Long
    Dim SqlText As String
    Dim LoadedCount As Long
    
    Set Result = New Collection
    Set Conn = New ADODB.Connection
    
    Conn.Open GetExcelConnectionString(FilePath)
    
    SheetName = GetFirstWorksheetName(Conn)
    If Len(Trim$(SheetName)) = 0 Then
        Err.Raise vbObjectError + 1200, "clsExcelReader.LoadRowsByAdo", "No worksheet could be found in the Excel file."
    End If
    
    SqlText = "SELECT * FROM " & BracketName(SheetName)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    LoadedCount = 0
    
    Do While Not Rs.EOF
        Set RowData = CreateObject("Scripting.Dictionary")
        
        For i = 0 To Rs.Fields.Count - 1
            RowData(Trim$(Rs.Fields(i).Name)) = Rs.Fields(i).Value
        Next i
        
        Result.Add RowData
        LoadedCount = LoadedCount + 1
        
        If MaxRows > 0 Then
            If LoadedCount >= MaxRows Then Exit Do
        End If
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
    
    Conn.Close
    Set Conn = Nothing
    
    Set LoadRowsByAdo = Result
    Exit Function

ErrorHandler:
    On Error Resume Next
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    If Not Conn Is Nothing Then
        If Conn.State = adStateOpen Then Conn.Close
    End If
    Set Conn = Nothing
    
    Err.Raise vbObjectError + 1201, "clsExcelReader.LoadRowsByAdo", _
              "Failed to read Excel rows using ADO/OLEDB." & vbCrLf & _
              "You can try Excel Automation mode instead."
End Function

Private Function LoadRowsByAutomation(ByVal FilePath As String, Optional ByVal OwnerForm As Object = Nothing, Optional ByVal MaxRows As Long = 0) As Collection
    On Error GoTo ErrorHandler
    
    Dim ExcelApp As Object
    Dim WorkBook As Object
    Dim WorkSheet As Object
    Dim Result As Collection
    Dim RowData As Object
    Dim LastRow As Long
    Dim LastCol As Long
    Dim EffectiveLastRow As Long
    Dim r As Long
    Dim c As Long
    Dim HeaderText As String
    Dim CellValue As Variant
    Const xlToLeft As Long = -4159
    Const xlUp As Long = -4162
    
    Set Result = New Collection
    
    Set ExcelApp = CreateObject("Excel.Application")
    '''ExcelApp.Visible = False
    
    Set WorkBook = ExcelApp.Workbooks.Open(FilePath, 0, True, , "", "", True)
    Set WorkSheet = WorkBook.Worksheets(1)
    
    LastCol = WorkSheet.Cells(1, WorkSheet.Columns.Count).End(xlToLeft).Column
    LastRow = WorkSheet.Cells(WorkSheet.Rows.Count, 1).End(xlUp).Row
    
    If LastCol < 1 Then
        Err.Raise vbObjectError + 1210, "clsExcelReader.LoadRowsByAutomation", "No columns were found in the Excel file."
    End If
    
    If LastRow < 2 Then
        Set LoadRowsByAutomation = Result
        GoTo SafeExit
    End If
    
    EffectiveLastRow = LastRow
    
    If MaxRows > 0 Then
        If (1 + MaxRows) < EffectiveLastRow Then
            EffectiveLastRow = 1 + MaxRows
        End If
    End If
    
    For r = 2 To EffectiveLastRow
        If Not OwnerForm Is Nothing Then
            If OwnerForm.CancelRequested Then
                Err.Raise vbObjectError + 3004, "clsExcelReader.LoadRowsByAutomation", "Import was cancelled by user while reading Excel rows."
            End If
        End If
        
        Set RowData = CreateObject("Scripting.Dictionary")
        
        For c = 1 To LastCol
            HeaderText = Trim$(CStr(WorkSheet.Cells(1, c).Value))
            If Len(HeaderText) = 0 Then
                HeaderText = "Column" & CStr(c)
            End If
            
            CellValue = WorkSheet.Cells(r, c).Value
            RowData(HeaderText) = CellValue
        Next c
        
        Result.Add RowData
        
        If Not OwnerForm Is Nothing Then
            On Error Resume Next
            
            OwnerForm.lblCurrentStep.Caption = "Reading Excel rows..."
            OwnerForm.lblRowProgress.Caption = CStr(r - 1) & " / " & CStr(EffectiveLastRow - 1)
            
            If EffectiveLastRow > 1 Then
                OwnerForm.prgImport.Value = CLng(((r - 1) / (EffectiveLastRow - 1)) * PROGRESS_MAX)
            Else
                OwnerForm.prgImport.Value = 0
            End If
            
            On Error GoTo ErrorHandler
        End If
        
        If ((r - 1) Mod 5) = 0 Or r = EffectiveLastRow Then
            DoEvents
        End If
    Next r
    
    Set LoadRowsByAutomation = Result

SafeExit:
    On Error Resume Next
    If Not WorkBook Is Nothing Then WorkBook.Close False
    If Not ExcelApp Is Nothing Then ExcelApp.Quit
    
    Set WorkSheet = Nothing
    Set WorkBook = Nothing
    Set ExcelApp = Nothing
    Exit Function

ErrorHandler:
    Dim OriginalErrNumber As Long
    Dim OriginalErrDescription As String
    
    OriginalErrNumber = Err.Number
    OriginalErrDescription = Err.Description
    
    On Error Resume Next
    
    If Not WorkBook Is Nothing Then WorkBook.Close False
    If Not ExcelApp Is Nothing Then ExcelApp.Quit
    
    Set WorkSheet = Nothing
    Set WorkBook = Nothing
    Set ExcelApp = Nothing
    
    Err.Clear
    
    If OriginalErrNumber = vbObjectError + 3004 Then
        Err.Raise vbObjectError + 3004, "clsExcelReader.LoadRowsByAutomation", OriginalErrDescription
    Else
        Err.Raise vbObjectError + 1211, "clsExcelReader.LoadRowsByAutomation", _
                  "Failed to read Excel rows using Excel Automation." & vbCrLf & _
                  "Please make sure the file is accessible." & vbCrLf & _
                  "If needed, copy the file to a normal local folder and try again."
    End If
End Function

