VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSqlServerConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mConnection As ADODB.Connection
Private mConnectionString As String

Public Property Get Connection() As ADODB.Connection
    Set Connection = mConnection
End Property

Public Property Get ConnectionString() As String
    ConnectionString = mConnectionString
End Property

Public Property Get IsConnected() As Boolean
    On Error Resume Next
    
    If mConnection Is Nothing Then
        IsConnected = False
    Else
        IsConnected = (mConnection.State = adStateOpen)
    End If
End Property

Public Function BuildConnectionString( _
    ByVal ServerName As String, _
    ByVal AuthenticationType As String, _
    Optional ByVal UserName As String = "", _
    Optional ByVal Password As String = "" _
) As String

    Dim Result As String
    
    Result = "Provider=SQLOLEDB;" & _
             "Data Source=" & ServerName & ";" & _
             "Initial Catalog=master;"
    
    If AuthenticationType = AUTH_WINDOWS Then
        Result = Result & "Integrated Security=SSPI;"
    Else
        Result = Result & "User ID=" & UserName & ";" & _
                          "Password=" & Password & ";"
    End If
    
    mConnectionString = Result
    BuildConnectionString = Result
End Function

Public Function OpenConnection( _
    ByVal ServerName As String, _
    ByVal AuthenticationType As String, _
    Optional ByVal UserName As String = "", _
    Optional ByVal Password As String = "" _
) As Boolean

    On Error GoTo ErrorHandler
    
    Call CloseConnection
    
    Set mConnection = New ADODB.Connection
    mConnection.ConnectionTimeout = 15
    mConnection.CommandTimeout = 60
    mConnection.CursorLocation = adUseClient
    
    mConnectionString = BuildConnectionString(ServerName, AuthenticationType, UserName, Password)
    mConnection.Open mConnectionString
    
    OpenConnection = (mConnection.State = adStateOpen)
    Exit Function

ErrorHandler:
    OpenConnection = False
    Err.Raise Err.Number, "clsSqlServerConnection.OpenConnection", Err.Description
End Function

Public Sub CloseConnection()
    On Error Resume Next
    
    If Not mConnection Is Nothing Then
        If mConnection.State <> adStateClosed Then
            mConnection.Close
        End If
    End If
    
    Set mConnection = Nothing
End Sub

Private Sub Class_Terminate()
    Call CloseConnection
End Sub
