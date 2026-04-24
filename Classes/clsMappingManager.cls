VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMappingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub EnsureMappingTableExists(ByVal Conn As ADODB.Connection, ByVal DatabaseName As String)
    On Error GoTo ErrorHandler
    
    Dim SqlText As String
    Dim Rs As ADODB.Recordset
    Dim SafeDatabaseName As String
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = ""
    SqlText = SqlText & "SELECT 1 " & vbCrLf
    SqlText = SqlText & "FROM " & SafeDatabaseName & ".sys.tables t " & vbCrLf
    SqlText = SqlText & "INNER JOIN " & SafeDatabaseName & ".sys.schemas s ON t.schema_id = s.schema_id " & vbCrLf
    SqlText = SqlText & "WHERE s.name = '" & APP_MAPPING_SCHEMA & "' " & vbCrLf
    SqlText = SqlText & "  AND t.name = '" & APP_MAPPING_TABLE & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        CreateMappingTable Conn, DatabaseName
    Else
        Rs.Close
        Set Rs = Nothing
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "clsMappingManager.EnsureMappingTableExists", Err.Description
End Sub

Private Sub CreateMappingTable(ByVal Conn As ADODB.Connection, ByVal DatabaseName As String)
    On Error GoTo ErrorHandler
    
    Dim SqlText As String
    Dim SafeDatabaseName As String
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = ""
    SqlText = SqlText & "CREATE TABLE " & SafeDatabaseName & "." & BracketName(APP_MAPPING_SCHEMA) & "." & BracketName(APP_MAPPING_TABLE) & " (" & vbCrLf
    SqlText = SqlText & "    MappingId INT IDENTITY(1,1) NOT NULL PRIMARY KEY," & vbCrLf
    SqlText = SqlText & "    SchemaName NVARCHAR(128) NOT NULL," & vbCrLf
    SqlText = SqlText & "    TableName NVARCHAR(128) NOT NULL," & vbCrLf
    SqlText = SqlText & "    ExcelColumnName NVARCHAR(255) NOT NULL," & vbCrLf
    SqlText = SqlText & "    SqlColumnName NVARCHAR(128) NOT NULL," & vbCrLf
    SqlText = SqlText & "    SqlDataType NVARCHAR(128) NULL," & vbCrLf
    SqlText = SqlText & "    IsActive BIT NOT NULL CONSTRAINT DF_" & APP_MAPPING_TABLE & "_IsActive DEFAULT(1)," & vbCrLf
    SqlText = SqlText & "    CreatedAt DATETIME NOT NULL CONSTRAINT DF_" & APP_MAPPING_TABLE & "_CreatedAt DEFAULT(GETDATE())," & vbCrLf
    SqlText = SqlText & "    UpdatedAt DATETIME NULL" & vbCrLf
    SqlText = SqlText & ")"
    
    Conn.Execute SqlText
    
    SqlText = ""
    SqlText = SqlText & "ALTER TABLE " & SafeDatabaseName & "." & BracketName(APP_MAPPING_SCHEMA) & "." & BracketName(APP_MAPPING_TABLE) & vbCrLf
    SqlText = SqlText & "ADD CONSTRAINT UQ_" & APP_MAPPING_TABLE & "_Schema_Table_ExcelColumn " & vbCrLf
    SqlText = SqlText & "UNIQUE (SchemaName, TableName, ExcelColumnName)"
    
    Conn.Execute SqlText
    
    SqlText = ""
    SqlText = SqlText & "CREATE INDEX IX_" & APP_MAPPING_TABLE & "_Schema_Table " & vbCrLf
    SqlText = SqlText & "ON " & SafeDatabaseName & "." & BracketName(APP_MAPPING_SCHEMA) & "." & BracketName(APP_MAPPING_TABLE) & " (SchemaName, TableName)"
    
    Conn.Execute SqlText
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "clsMappingManager.CreateMappingTable", Err.Description
End Sub

Public Function AutoMatch(ByVal ExcelHeaders As Collection, ByVal TargetColumns As Collection) As Collection
    On Error GoTo ErrorHandler
    
    Dim Result As Collection
    Dim MapItem As clsMappingItem
    Dim ExcelName As String
    Dim ColInfo As clsColumnInfo
    Dim i As Long
    Dim j As Long
    
    Set Result = New Collection
    
    For i = 1 To ExcelHeaders.Count
        ExcelName = NzString(ExcelHeaders(i))
        
        Set MapItem = New clsMappingItem
        MapItem.ExcelColumnName = ExcelName
        MapItem.SqlColumnName = ""
        MapItem.SqlDataType = ""
        MapItem.IsRequired = False
        MapItem.IsMapped = False
        
        For j = 1 To TargetColumns.Count
            Set ColInfo = TargetColumns(j)
            
            If NormalizeNameForMatch(ExcelName) = NormalizeNameForMatch(ColInfo.ColumnName) Then
                MapItem.SqlColumnName = ColInfo.ColumnName
                MapItem.SqlDataType = ColInfo.DataType
                MapItem.IsRequired = ColInfo.IsRequired
                MapItem.IsMapped = True
                Exit For
            End If
        Next j
        
        Result.Add MapItem
    Next i
    
    Set AutoMatch = Result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsMappingManager.AutoMatch", Err.Description
End Function

Private Function NormalizeNameForMatch(ByVal Value As String) As String
    Dim Result As String
    
    Result = LCase$(Trim$(Value))
    Result = Replace(Result, " ", "")
    Result = Replace(Result, "_", "")
    Result = Replace(Result, "-", "")
    Result = Replace(Result, ".", "")
    
    NormalizeNameForMatch = Result
End Function

Public Sub SaveMapping( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal TableName As String, _
    ByVal Mappings As Collection)

    On Error GoTo ErrorHandler
    
    Dim SqlText As String
    Dim SafeDatabaseName As String
    Dim i As Long
    Dim Item As clsMappingItem
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = ""
    SqlText = SqlText & "DELETE FROM " & SafeDatabaseName & "." & BracketName(APP_MAPPING_SCHEMA) & "." & BracketName(APP_MAPPING_TABLE) & vbCrLf
    SqlText = SqlText & "WHERE SchemaName = '" & EscapeSqlText(SchemaName) & "' " & vbCrLf
    SqlText = SqlText & "  AND TableName = '" & EscapeSqlText(TableName) & "'"
    
    Conn.Execute SqlText
    
    For i = 1 To Mappings.Count
        Set Item = Mappings(i)
        
        If Item.IsMapped And Len(Item.SqlColumnName) > 0 Then
            SqlText = ""
            SqlText = SqlText & "INSERT INTO " & SafeDatabaseName & "." & BracketName(APP_MAPPING_SCHEMA) & "." & BracketName(APP_MAPPING_TABLE) & " (" & vbCrLf
            SqlText = SqlText & "    SchemaName, TableName, ExcelColumnName, SqlColumnName, SqlDataType, IsActive, CreatedAt, UpdatedAt" & vbCrLf
            SqlText = SqlText & ") VALUES (" & vbCrLf
            SqlText = SqlText & "    '" & EscapeSqlText(SchemaName) & "'," & vbCrLf
            SqlText = SqlText & "    '" & EscapeSqlText(TableName) & "'," & vbCrLf
            SqlText = SqlText & "    '" & EscapeSqlText(Item.ExcelColumnName) & "'," & vbCrLf
            SqlText = SqlText & "    '" & EscapeSqlText(Item.SqlColumnName) & "'," & vbCrLf
            SqlText = SqlText & "    '" & EscapeSqlText(Item.SqlDataType) & "'," & vbCrLf
            SqlText = SqlText & "    1, GETDATE(), NULL" & vbCrLf
            SqlText = SqlText & ")"
            
            Conn.Execute SqlText
        End If
    Next i
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "clsMappingManager.SaveMapping", Err.Description
End Sub

Public Function LoadMapping( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal TableName As String) As Collection

    On Error GoTo ErrorHandler
    
    Dim SqlText As String
    Dim SafeDatabaseName As String
    Dim Rs As ADODB.Recordset
    Dim Result As Collection
    Dim Item As clsMappingItem
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = ""
    SqlText = SqlText & "SELECT ExcelColumnName, SqlColumnName, SqlDataType " & vbCrLf
    SqlText = SqlText & "FROM " & SafeDatabaseName & "." & BracketName(APP_MAPPING_SCHEMA) & "." & BracketName(APP_MAPPING_TABLE) & vbCrLf
    SqlText = SqlText & "WHERE SchemaName = '" & EscapeSqlText(SchemaName) & "' " & vbCrLf
    SqlText = SqlText & "  AND TableName = '" & EscapeSqlText(TableName) & "' " & vbCrLf
    SqlText = SqlText & "  AND IsActive = 1 " & vbCrLf
    SqlText = SqlText & "ORDER BY ExcelColumnName"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    Set Result = New Collection
    
    Do While Not Rs.EOF
        Set Item = New clsMappingItem
        
        Item.ExcelColumnName = NzString(Rs.Fields("ExcelColumnName").Value)
        Item.SqlColumnName = NzString(Rs.Fields("SqlColumnName").Value)
        Item.SqlDataType = NzString(Rs.Fields("SqlDataType").Value)
        Item.IsMapped = (Len(Item.SqlColumnName) > 0)
        
        Result.Add Item
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
    
    Set LoadMapping = Result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsMappingManager.LoadMapping", Err.Description
End Function

Public Function ValidateRequiredMappings(ByVal TargetColumns As Collection, ByVal Mappings As Collection) As String
    Dim i As Long
    Dim j As Long
    Dim ColInfo As clsColumnInfo
    Dim MapItem As clsMappingItem
    Dim Found As Boolean
    Dim MissingList As String
    
    MissingList = ""
    
    For i = 1 To TargetColumns.Count
        Set ColInfo = TargetColumns(i)
        
        If ColInfo.IsRequired Then
            Found = False
            
            For j = 1 To Mappings.Count
                Set MapItem = Mappings(j)
                
                If MapItem.IsMapped Then
                    If StrComp(Trim$(MapItem.SqlColumnName), Trim$(ColInfo.ColumnName), vbTextCompare) = 0 Then
                        Found = True
                        Exit For
                    End If
                End If
            Next j
            
            If Not Found Then
                MissingList = MissingList & ColInfo.ColumnName & " (" & ColInfo.DataType & ")" & vbCrLf
            End If
        End If
    Next i
    
    ValidateRequiredMappings = MissingList
End Function
