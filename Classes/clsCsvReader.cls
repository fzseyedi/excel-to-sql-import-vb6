VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCsvReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDelimiter As String
Private mEncodingMode As String

Private Const CSV_ENCODING_AUTO As String = "AUTO"
Private Const CSV_ENCODING_UTF8 As String = "UTF-8"
Private Const CSV_ENCODING_ANSI As String = "ANSI"

Private Sub Class_Initialize()
    mDelimiter = ","
    mEncodingMode = CSV_ENCODING_AUTO
End Sub

Public Property Get Delimiter() As String
    Delimiter = mDelimiter
End Property

Public Property Let Delimiter(ByVal Value As String)
    If Len(Value) = 0 Then
        mDelimiter = ","
    Else
        mDelimiter = Left$(Value, 1)
    End If
End Property

Private Function HasUtf8Bom(ByVal FilePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim FileNo As Integer
    Dim b1 As Byte
    Dim b2 As Byte
    Dim b3 As Byte
    
    If Len(Dir$(FilePath)) = 0 Then Exit Function
    If FileLen(FilePath) < 3 Then Exit Function
    
    FileNo = FreeFile
    
    Open FilePath For Binary Access Read As #FileNo
    
    Get #FileNo, 1, b1
    Get #FileNo, 2, b2
    Get #FileNo, 3, b3
    
    Close #FileNo
    
    HasUtf8Bom = (b1 = &HEF And b2 = &HBB And b3 = &HBF)
    Exit Function

ErrorHandler:
    On Error Resume Next
    If FileNo > 0 Then Close #FileNo
    HasUtf8Bom = False
End Function

Public Function LoadHeaders(ByVal FilePath As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim Lines As Collection
    Dim HeaderLine As String
    Dim Values As Collection
    Dim Result As Collection
    Dim i As Long
    
    Set Result = New Collection
    
    Set Lines = ReadAllLines(FilePath)
    
    If Lines.Count = 0 Then
        Err.Raise vbObjectError + 6100, "clsCsvReader.LoadHeaders", "CSV file is empty."
    End If
    
    HeaderLine = CStr(Lines(1))
    mDelimiter = DetectDelimiterFromLine(HeaderLine)
    Set Values = ParseCsvLine(HeaderLine)
    
    For i = 1 To Values.Count
        If Len(Trim$(CStr(Values(i)))) > 0 Then
            Result.Add Trim$(CStr(Values(i)))
        Else
            Result.Add "Column" & CStr(i)
        End If
    Next i
    
    Set LoadHeaders = Result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsCsvReader.LoadHeaders", Err.Description
End Function

Public Function LoadRows(ByVal FilePath As String, Optional ByVal OwnerForm As Object = Nothing) As Collection
    On Error GoTo ErrorHandler
    
    Dim Lines As Collection
    Dim Headers As Collection
    Dim Values As Collection
    Dim Result As Collection
    Dim RowData As Object
    Dim i As Long
    Dim j As Long
    Dim HeaderName As String
    Dim ValueText As String
    
    Set Result = New Collection
    Set Lines = ReadAllLines(FilePath)
    
    If Lines.Count = 0 Then
        Err.Raise vbObjectError + 6101, "clsCsvReader.LoadRows", "CSV file is empty."
    End If
    
    Set Headers = LoadHeaders(FilePath)
    
    For i = 2 To Lines.Count
        If Not OwnerForm Is Nothing Then
            If OwnerForm.CancelRequested Then
                Err.Raise vbObjectError + 3001, "clsCsvReader.LoadRows", "Import was cancelled by user while reading CSV rows."
            End If
        End If
        
        If Len(Trim$(CStr(Lines(i)))) > 0 Then
            Set Values = ParseCsvLine(CStr(Lines(i)))
            Set RowData = CreateObject("Scripting.Dictionary")
            
            For j = 1 To Headers.Count
                HeaderName = CStr(Headers(j))
                
                If j <= Values.Count Then
                    ValueText = CStr(Values(j))
                Else
                    ValueText = ""
                End If
                
                RowData(HeaderName) = ValueText
            Next j
            
            Result.Add RowData
        End If
        
        If (i Mod 50) = 0 Then
            DoEvents
        End If
    Next i
    
    Set LoadRows = Result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsCsvReader.LoadRows", Err.Description
End Function

Public Function LoadPreviewRows(ByVal FilePath As String, ByVal MaxRows As Long) As Collection
    On Error GoTo ErrorHandler
    
    Dim Lines As Collection
    Dim Headers As Collection
    Dim Values As Collection
    Dim Result As Collection
    Dim RowData As Object
    Dim i As Long
    Dim j As Long
    Dim DataRowCount As Long
    Dim HeaderName As String
    Dim ValueText As String
    
    Set Result = New Collection
    Set Lines = ReadAllLines(FilePath)
    
    If Lines.Count = 0 Then
        Err.Raise vbObjectError + 6102, "clsCsvReader.LoadPreviewRows", "CSV file is empty."
    End If
    
    Set Headers = LoadHeaders(FilePath)
    
    DataRowCount = 0
    
    For i = 2 To Lines.Count
        If DataRowCount >= MaxRows Then Exit For
        
        If Len(Trim$(CStr(Lines(i)))) > 0 Then
            Set Values = ParseCsvLine(CStr(Lines(i)))
            Set RowData = CreateObject("Scripting.Dictionary")
            
            For j = 1 To Headers.Count
                HeaderName = CStr(Headers(j))
                
                If j <= Values.Count Then
                    ValueText = CStr(Values(j))
                Else
                    ValueText = ""
                End If
                
                RowData(HeaderName) = ValueText
            Next j
            
            Result.Add RowData
            DataRowCount = DataRowCount + 1
        End If
    Next i
    
    Set LoadPreviewRows = Result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsCsvReader.LoadPreviewRows", Err.Description
End Function

Private Function ReadAllLines(ByVal FilePath As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim Result As Collection
    Dim FileText As String
    Dim Lines() As String
    Dim i As Long
    
    Set Result = New Collection
    
    Select Case mEncodingMode
        Case CSV_ENCODING_UTF8
            FileText = ReadTextFileUtf8(FilePath)
        
        Case CSV_ENCODING_ANSI
            Set ReadAllLines = ReadAllLinesAnsi(FilePath)
            Exit Function
        
        Case Else
            If HasUtf8Bom(FilePath) Then
                FileText = ReadTextFileUtf8(FilePath)
            Else
                Set ReadAllLines = ReadAllLinesAnsi(FilePath)
                Exit Function
            End If
    End Select
    
    FileText = RemoveUnicodeBom(FileText)
    FileText = Replace(FileText, vbCrLf, vbLf)
    FileText = Replace(FileText, vbCr, vbLf)
    
    Lines = Split(FileText, vbLf)
    
    For i = LBound(Lines) To UBound(Lines)
        Result.Add Lines(i)
    Next i
    
    Set ReadAllLines = Result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsCsvReader.ReadAllLines", Err.Description
End Function

Private Function ReadAllLinesAnsi(ByVal FilePath As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim Result As Collection
    Dim FileNo As Integer
    Dim LineText As String
    
    Set Result = New Collection
    
    FileNo = FreeFile
    
    Open FilePath For Input As #FileNo
    
    Do While Not EOF(FileNo)
        Line Input #FileNo, LineText
        Result.Add LineText
    Loop
    
    Close #FileNo
    
    Set ReadAllLinesAnsi = Result
    Exit Function

ErrorHandler:
    On Error Resume Next
    If FileNo > 0 Then Close #FileNo
    
    Err.Raise Err.Number, "clsCsvReader.ReadAllLinesAnsi", Err.Description
End Function

Private Function ParseCsvLine(ByVal LineText As String) As Collection
    Dim Result As Collection
    Dim i As Long
    Dim ch As String
    Dim CurrentValue As String
    Dim InQuotes As Boolean
    Dim NextChar As String
    
    Set Result = New Collection
    
    CurrentValue = ""
    InQuotes = False
    
    i = 1
    
    Do While i <= Len(LineText)
        ch = Mid$(LineText, i, 1)
        
        If ch = """" Then
            If InQuotes Then
                If i < Len(LineText) Then
                    NextChar = Mid$(LineText, i + 1, 1)
                    
                    If NextChar = """" Then
                        CurrentValue = CurrentValue & """"
                        i = i + 1
                    Else
                        InQuotes = False
                    End If
                Else
                    InQuotes = False
                End If
            Else
                InQuotes = True
            End If
        
        ElseIf ch = mDelimiter And Not InQuotes Then
            Result.Add CurrentValue
            CurrentValue = ""
        
        Else
            CurrentValue = CurrentValue & ch
        End If
        
        i = i + 1
    Loop
    
    Result.Add CurrentValue
    
    Set ParseCsvLine = Result
End Function

Public Property Get EncodingMode() As String
    EncodingMode = mEncodingMode
End Property

Public Property Let EncodingMode(ByVal Value As String)
    Value = UCase$(Trim$(Value))
    
    Select Case Value
        Case CSV_ENCODING_UTF8
            mEncodingMode = CSV_ENCODING_UTF8
        
        Case CSV_ENCODING_ANSI
            mEncodingMode = CSV_ENCODING_ANSI
        
        Case Else
            mEncodingMode = CSV_ENCODING_AUTO
    End Select
End Property

Private Function ReadTextFileUtf8(ByVal FilePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim Stream As Object
    
    Set Stream = CreateObject("ADODB.Stream")
    
    Stream.Type = 2        ' adTypeText
    Stream.Charset = "utf-8"
    Stream.Open
    Stream.LoadFromFile FilePath
    
    ReadTextFileUtf8 = Stream.ReadText
    
    Stream.Close
    Set Stream = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not Stream Is Nothing Then
        If Stream.State = 1 Then Stream.Close
    End If
    Set Stream = Nothing
    
    Err.Raise Err.Number, "clsCsvReader.ReadTextFileUtf8", Err.Description
End Function

Private Function RemoveUnicodeBom(ByVal TextValue As String) As String
    If Len(TextValue) > 0 Then
        If AscW(Left$(TextValue, 1)) = &HFEFF Then
            RemoveUnicodeBom = Mid$(TextValue, 2)
            Exit Function
        End If
    End If
    
    RemoveUnicodeBom = TextValue
End Function

Private Function DetectDelimiterFromLine(ByVal LineText As String) As String
    Dim CommaCount As Long
    Dim SemicolonCount As Long
    Dim TabCount As Long
    
    CommaCount = CountDelimiterOutsideQuotes(LineText, ",")
    SemicolonCount = CountDelimiterOutsideQuotes(LineText, ";")
    TabCount = CountDelimiterOutsideQuotes(LineText, vbTab)
    
    If SemicolonCount >= CommaCount And SemicolonCount >= TabCount And SemicolonCount > 0 Then
        DetectDelimiterFromLine = ";"
    ElseIf TabCount >= CommaCount And TabCount > 0 Then
        DetectDelimiterFromLine = vbTab
    Else
        DetectDelimiterFromLine = ","
    End If
End Function

Private Function CountDelimiterOutsideQuotes(ByVal LineText As String, ByVal DelimiterText As String) As Long
    Dim i As Long
    Dim ch As String
    Dim InQuotes As Boolean
    Dim NextChar As String
    
    InQuotes = False
    
    i = 1
    Do While i <= Len(LineText)
        ch = Mid$(LineText, i, 1)
        
        If ch = """" Then
            If InQuotes Then
                If i < Len(LineText) Then
                    NextChar = Mid$(LineText, i + 1, 1)
                    
                    If NextChar = """" Then
                        i = i + 1
                    Else
                        InQuotes = False
                    End If
                Else
                    InQuotes = False
                End If
            Else
                InQuotes = True
            End If
        
        ElseIf ch = DelimiterText And Not InQuotes Then
            CountDelimiterOutsideQuotes = CountDelimiterOutsideQuotes + 1
        End If
        
        i = i + 1
    Loop
End Function
