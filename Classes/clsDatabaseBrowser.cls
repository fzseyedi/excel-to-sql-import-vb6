VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabaseBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetUserDatabases(ByVal Conn As ADODB.Connection) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim Rs As ADODB.Recordset
    Dim SqlText As String
    
    SqlText = ""
    SqlText = SqlText & "SELECT name " & vbCrLf
    SqlText = SqlText & "FROM sys.databases " & vbCrLf
    SqlText = SqlText & "WHERE name NOT IN ('master', 'model', 'msdb', 'tempdb') " & vbCrLf
    SqlText = SqlText & "ORDER BY name"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    Set GetUserDatabases = Rs
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsDatabaseBrowser.GetUserDatabases", Err.Description
End Function

Public Function GetBaseTables(ByVal Conn As ADODB.Connection, ByVal DatabaseName As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim Rs As ADODB.Recordset
    Dim SqlText As String
    Dim SafeDatabaseName As String
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = ""
    SqlText = SqlText & "SELECT TABLE_SCHEMA, TABLE_NAME " & vbCrLf
    SqlText = SqlText & "FROM " & SafeDatabaseName & ".INFORMATION_SCHEMA.TABLES " & vbCrLf
    SqlText = SqlText & "WHERE TABLE_TYPE = 'BASE TABLE' " & vbCrLf
    SqlText = SqlText & "ORDER BY TABLE_SCHEMA, TABLE_NAME"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    Set GetBaseTables = Rs
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsDatabaseBrowser.GetBaseTables", Err.Description
End Function

Public Function GetTableColumns( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal TableName As String) As Collection

    On Error GoTo ErrorHandler

    Dim Rs As ADODB.Recordset
    Dim SqlText As String
    Dim SafeDatabaseName As String
    Dim Item As clsColumnInfo
    Dim Result As Collection
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = ""
    SqlText = SqlText & "SELECT " & vbCrLf
    SqlText = SqlText & "    c.TABLE_SCHEMA, " & vbCrLf
    SqlText = SqlText & "    c.TABLE_NAME, " & vbCrLf
    SqlText = SqlText & "    c.COLUMN_NAME, " & vbCrLf
    SqlText = SqlText & "    c.DATA_TYPE, " & vbCrLf
    SqlText = SqlText & "    c.IS_NULLABLE, " & vbCrLf
    SqlText = SqlText & "    c.CHARACTER_MAXIMUM_LENGTH, " & vbCrLf
    SqlText = SqlText & "    c.NUMERIC_PRECISION, " & vbCrLf
    SqlText = SqlText & "    c.NUMERIC_SCALE, " & vbCrLf
    SqlText = SqlText & "    c.ORDINAL_POSITION, " & vbCrLf
    SqlText = SqlText & "    COLUMNPROPERTY(OBJECT_ID(c.TABLE_SCHEMA + '.' + c.TABLE_NAME), c.COLUMN_NAME, 'IsIdentity') AS IsIdentity, " & vbCrLf
    SqlText = SqlText & "    COLUMNPROPERTY(OBJECT_ID(c.TABLE_SCHEMA + '.' + c.TABLE_NAME), c.COLUMN_NAME, 'IsComputed') AS IsComputed, " & vbCrLf
    SqlText = SqlText & "    CASE WHEN c.COLUMN_DEFAULT IS NULL THEN 0 ELSE 1 END AS HasDefault " & vbCrLf
    SqlText = SqlText & "FROM " & SafeDatabaseName & ".INFORMATION_SCHEMA.COLUMNS c " & vbCrLf
    SqlText = SqlText & "WHERE c.TABLE_SCHEMA = '" & EscapeSqlText(SchemaName) & "' " & vbCrLf
    SqlText = SqlText & "  AND c.TABLE_NAME = '" & EscapeSqlText(TableName) & "' " & vbCrLf
    SqlText = SqlText & "ORDER BY c.ORDINAL_POSITION"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    Set Result = New Collection
    
    Do While Not Rs.EOF
        Set Item = New clsColumnInfo
        
        Item.SchemaName = NzString(Rs.Fields("TABLE_SCHEMA").Value)
        Item.TableName = NzString(Rs.Fields("TABLE_NAME").Value)
        Item.ColumnName = NzString(Rs.Fields("COLUMN_NAME").Value)
        Item.DataType = NzString(Rs.Fields("DATA_TYPE").Value)
        Item.IsNullable = (UCase$(NzString(Rs.Fields("IS_NULLABLE").Value)) = "YES")
        Item.CharacterMaxLength = SafeCLng(Rs.Fields("CHARACTER_MAXIMUM_LENGTH").Value, 0)
        Item.NumericPrecision = SafeCLng(Rs.Fields("NUMERIC_PRECISION").Value, 0)
        Item.NumericScale = SafeCLng(Rs.Fields("NUMERIC_SCALE").Value, 0)
        Item.IsIdentity = (SafeCLng(Rs.Fields("IsIdentity").Value, 0) = 1)
        Item.IsComputed = (SafeCLng(Rs.Fields("IsComputed").Value, 0) = 1)
        Item.HasDefault = (SafeCLng(Rs.Fields("HasDefault").Value, 0) = 1)
        Item.OrdinalPosition = SafeCLng(Rs.Fields("ORDINAL_POSITION").Value, 0)
        Item.IsRequired = (Not Item.IsNullable) And (Not Item.HasDefault) And (Not Item.IsComputed)
        
        Result.Add Item
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
    
    Set GetTableColumns = Result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsDatabaseBrowser.GetTableColumns", Err.Description
End Function

