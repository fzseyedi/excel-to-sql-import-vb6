Attribute VB_Name = "modIniFile"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Public Function ReadIniValue(ByVal SectionName As String, ByVal KeyName As String, ByVal DefaultValue As String, ByVal FilePath As String) As String
    Dim Buffer As String
    Dim ResultLen As Long
    
    Buffer = String$(1024, vbNullChar)
    ResultLen = GetPrivateProfileString(SectionName, KeyName, DefaultValue, Buffer, Len(Buffer), FilePath)
    
    If ResultLen > 0 Then
        ReadIniValue = Left$(Buffer, ResultLen)
    Else
        ReadIniValue = DefaultValue
    End If
End Function

Public Sub WriteIniValue(ByVal SectionName As String, ByVal KeyName As String, ByVal ValueText As String, ByVal FilePath As String)
    WritePrivateProfileString SectionName, KeyName, ValueText, FilePath
End Sub

Public Sub DeleteIniKey(ByVal SectionName As String, ByVal KeyName As String, ByVal FilePath As String)
    WritePrivateProfileString SectionName, KeyName, ByVal 0&, FilePath
End Sub

