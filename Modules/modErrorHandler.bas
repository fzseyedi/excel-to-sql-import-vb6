Attribute VB_Name = "modErrorHandler"
Public Function GetFriendlyImportErrorMessage(ByVal ErrorNumber As Long, ByVal ErrorDescription As String) As String
    Dim DescText As String
    Dim Msg As String
    
    DescText = LCase$(Trim$(ErrorDescription))
    Msg = ""
    
    If InStr(1, DescText, "text value is too long", vbTextCompare) > 0 _
       Or InStr(1, DescText, "max length=", vbTextCompare) > 0 _
       Or InStr(1, DescText, "actual length=", vbTextCompare) > 0 Then
        
        Msg = "Import failed because the length of one or more Excel values is greater than the allowed length of the target SQL field."
        Msg = Msg & vbCrLf & ErrorDescription
        GetFriendlyImportErrorMessage = Msg
        Exit Function
    End If
    
    If InStr(1, DescText, "duplicate", vbTextCompare) > 0 _
       Or InStr(1, DescText, "unique key", vbTextCompare) > 0 _
       Or InStr(1, DescText, "primary key", vbTextCompare) > 0 Then
        
        Msg = "Import failed because a duplicate record was found."
        Msg = Msg & vbCrLf & ErrorDescription
        GetFriendlyImportErrorMessage = Msg
        Exit Function
    End If
    
    If InStr(1, DescText, "foreign key", vbTextCompare) > 0 _
       Or InStr(1, DescText, "reference constraint", vbTextCompare) > 0 Then
        
        Msg = "Import failed because one or more values do not match valid parent records."
        Msg = Msg & vbCrLf & ErrorDescription
        GetFriendlyImportErrorMessage = Msg
        Exit Function
    End If
    
    If InStr(1, DescText, "cannot insert the value null", vbTextCompare) > 0 _
       Or InStr(1, DescText, "null value", vbTextCompare) > 0 Then
        
        Msg = "Import failed because a required field does not have a value."
        Msg = Msg & vbCrLf & ErrorDescription
        GetFriendlyImportErrorMessage = Msg
        Exit Function
    End If
    
    If InStr(1, DescText, "identity_insert", vbTextCompare) > 0 Then
        
        Msg = "Import failed because IDENTITY_INSERT was not allowed or was not configured correctly."
        Msg = Msg & vbCrLf & ErrorDescription
        GetFriendlyImportErrorMessage = Msg
        Exit Function
    End If
    
    If InStr(1, DescText, "permission", vbTextCompare) > 0 _
       Or InStr(1, DescText, "denied", vbTextCompare) > 0 Then
        
        Msg = "Import failed because the SQL Server user does not have sufficient permissions."
        Msg = Msg & vbCrLf & ErrorDescription
        GetFriendlyImportErrorMessage = Msg
        Exit Function
    End If
    
    If InStr(1, DescText, "data validation error", vbTextCompare) > 0 _
       Or InStr(1, DescText, "data type", vbTextCompare) > 0 _
       Or InStr(1, DescText, "conversion", vbTextCompare) > 0 _
       Or InStr(1, DescText, "overflow", vbTextCompare) > 0 _
       Or InStr(1, DescText, "invalid value", vbTextCompare) > 0 Then
        
        Msg = "Import failed because one or more Excel values are not valid for the target SQL columns."
        Msg = Msg & vbCrLf & ErrorDescription
        GetFriendlyImportErrorMessage = Msg
        Exit Function
    End If
    
    Msg = "Import failed."
    Msg = Msg & vbCrLf & ErrorDescription
    
    GetFriendlyImportErrorMessage = Msg
End Function
