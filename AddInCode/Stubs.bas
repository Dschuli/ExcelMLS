Attribute VB_Name = "Stubs"
Option Explicit
'Privies stubs for functions that if/wanted needed have to be implemented by the user

Function IsPlatformDeveloper(Optional dummy As Boolean) As Boolean
'Checks if user is allowed to alter add-in data

    IsPlatformDeveloper = True
    
    'Example based on user name
'    If InStr(LCase(Application.UserName), "xxx") > 0 Then
'        IsPlatformDeveloper = True
'        Exit Function
'    End If
End Function
Function IsDeveloper(Optional dummy As Boolean) As Boolean
'Checks if user is allowed to use development functions

    IsDeveloper = True
    
    'Example based on user name
'    If InStr(LCase(Application.UserName), "xxx") > 0 Then
'        IsPlatformDeveloper = True
'        Exit Function
'    End If
End Function

Function GetSheetProtectionPassword(Optional sh As Worksheet) As String
'Returns a user/application specific passwoed to protect/unprotect the worksheet

    GetSheetProtectionPassword = c_noPWrequired

End Function

Function GetBookProtectionPassword() As String
'Returns a user/application specific passwoed to protect/unprotect the workbook

    GetBookProtectionPassword = c_noBookPWrequired

End Function
