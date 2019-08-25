Attribute VB_Name = "NlsUsingFunctions"
Option Explicit

Public Enum msgLevel                        'For showmMessage
    smInfo
    smWarning
    smError
    smSystemError
End Enum

Public Enum confirmLevel                    'For showConfirm
    cfInfo
    cfWarning
End Enum

Sub ShowMessage(module As String, Optional identifier As String, Optional level As msgLevel = smError, Optional p1 As String = "", Optional p2 As String = "", Optional p3 As String = "", Optional p4 As String = "")
    'Levels: 0 Info 1 Warning 2 Error 3 SystemError
    'Will provide option to handle system error differently in future
    'if only module is provided it will be used as the message text
    Dim lv(), cap()
    Dim addOn As String
    Dim ret As Variant
    
    level = Application.Median(0, 3, level)
    lv = Array(vbInformation, vbExclamation, vbCritical, vbCritical)
    cap = Array("Info", "Warning", "Error", "System")
    If level = 3 Then addOn = vbCrLf & GetNlsText("MsgBox", "SystemErrorAddOn")
    
    If identifier = "" Then
        ret = MsgBox(module & addOn, vbOKOnly + lv(level), GetNlsText("MsgBox", CStr(cap(level))))
    Else
        ret = MsgBox(GetNlsText(module, identifier, p1, p2, p3, p4) & addOn, vbOKOnly + lv(level), GetNlsText("MsgBox", CStr(cap(level))))
    End If
    
End Sub

Function ShowConfirm(module As String, Optional identifier As String, Optional level As confirmLevel = cfWarning, Optional p1 As String = "", Optional p2 As String = "", Optional p3 As String = "") As Boolean
    'Levels: 0 Question 1 Warning
    'if only module is provided it will be used as the message text
    Dim lv(), cap()
    level = Application.Median(0, 1, level)
    lv = Array(vbQuestion, vbExclamation)
    cap = Array("Question", "Warning")
    Dim ret As Variant
    
    If identifier = "" Then
        ret = MsgBox(module, vbYesNo + lv(level), GetNlsText("MsgBox", CStr(cap(level))))
    Else
        ret = MsgBox(GetNlsText(module, identifier, p1, p2, p3), vbYesNo + lv(level), GetNlsText("MsgBox", CStr(cap(level))))
    End If
    
    If ret = vbYes Then ShowConfirm = True
    
End Function
Function IsNo(str, Optional acceptBlanks As Boolean, Optional opt As String) As Boolean
'Check is a string represents a "no" value as defined in opt or a parm separated value in NLSText System/optionForNo
'If acceptBlank is true then an empty input will return true
    If Len(trim(str)) = 0 Then
        IsNo = acceptBlanks
        Exit Function
    End If
    If Len(opt) = 0 Then
        opt = GetNlsText(module:="system", identifier:="optionsForNo", mandatory:=True)
        If Len(opt) = 0 Then Exit Function
    End If
    IsNo = InStr(LCase(opt) & ",", LCase(CStr(str)) & ",") > 0
End Function
Function IsYes(str, Optional acceptBlanks As Boolean, Optional opt As String) As Boolean
'Check is a string represents a "yes" value as defined in opt or a parm separated value in NLSText System/optionForYes
'If acceptBlank is true then an empty input will return true
    If Len(trim(str)) = 0 Then
        IsYes = acceptBlanks
        Exit Function
    End If
    If Len(opt) = 0 Then
        opt = GetNlsText(module:="system", identifier:="optionsForYes", mandatory:=True)
        If Len(opt) = 0 Then Exit Function
    End If
    IsYes = InStr(LCase(opt) & ",", LCase(CStr(str)) & ",") > 0
End Function
