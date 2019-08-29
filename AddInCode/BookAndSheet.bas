Attribute VB_Name = "BookAndSheet"
Option Explicit

Sub BookOpen(Optional dummy As Boolean)

    Call checkNamedReferencesToWorkbook(inBook:=ThisWorkbook)

    Call SetLanguage

    Call checkAndCreateSystemSheets                                   'Check that some internally used sheets are present and create them if not
    
    Call setSheetProtectionToUserInterfaceOnly                        'Set the userInterfaceOnly protection parameter for all protected sheets
    
End Sub
Sub BookActivate()
    
    With Application
        .OnKey "^+m", "CreateDisplayDevelopmentPopUpMenu"
        .OnKey "^+t", "ShowNLSTable"
    End With
    
    Call SetNLSData
    
    Call SetSeparators
    
End Sub
Sub BookDeactivate()

    Call resetFastMode

    With Application
     .OnKey "^+m"
     .OnKey "^+t"
     .UseSystemSeparators = True
    End With
    
End Sub
Sub bookBeforeClose()
    g_NLSData.bookName = ""                     'Devalidate g_NLSData
End Sub
Sub sheetChange(sh As Worksheet, target As Range)

    
    On Error GoTo ErrorExit
    
    Select Case sh.Name
    
        Case "NLS"
            Call InvalidateNlsText
        
        Case "Demo"
            If target.Address = Range("SelectedLanguage").Address Then
                If ActiveWorkbook.Worksheets("Demo").Shapes("AutoSetLanguage").OLEFormat.Object.Value = xlOn Then Call SetLanguage
            End If
            
    End Select
    
ErrorExit:
    
End Sub
Function SheetBeforeRightClick(sh As Worksheet, target As Range) As Boolean
'Code that runs before right click
   
    Call ShowCustomizedCellContextMenu(sh, target)
    SheetBeforeRightClick = True
        
End Function

Sub checkAndCreateSystemSheets(Optional dummy As Variant)
'Check that some internally used sheets are present and create them if not
    Dim bookWasProtected As Boolean
    Dim sh As Worksheet
    
    On Error Resume Next
    
    Set sh = ActiveSheet
    
    Err.Clear
    dummy = Worksheets(c_infoSheetName).Name
    
    If Err.Number > 0 Then
       bookWasProtected = DoUnProtectWorkBook
       Err.Clear
       ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)).Name = c_infoSheetName
       If Err.Number > 0 Then GoTo ErrorExit
       ActiveSheet.Visible = xlHidden
    End If
    
    Err.Clear
    dummy = Worksheets(c_sortSheetname).Name
    
    If Err.Number > 0 Then
       bookWasProtected = DoUnProtectWorkBook
       ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)).Name = c_sortSheetname
       ActiveSheet.Visible = xlHidden
    End If
    
    If bookWasProtected Then Call DoProtectWorkBook
    
    If ActiveSheet.Name <> sh.Name Then sh.Activate
    
    Exit Sub
    
ErrorExit:
        
End Sub

Sub setSheetProtectionToUserInterfaceOnly(Optional dummy As Boolean)
'Sets sheet protection to userInterfaceOnly for protected sheets as this property is not persisted
    Dim sh As Worksheet
    Dim pw As String
    
    On Error Resume Next
    Err.Clear
                
    For Each sh In ActiveWorkbook.Worksheets
        If sh.ProtectContents Then
            pw = GetSheetProtectionPassword(sh)
            If pw <> c_noPWrequired Then
                If Len(pw) = 0 Then pw = AskForPassWord
                If Len(pw) = 0 Then
                    Call ShowMessage("System", "functionAborted", smError, "setSheetProtectionToUserInterfaceOnly")
                Else
                    sh.Protect Password:=pw, userinterfaceonly:=True
                End If
            Else
                sh.Protect userinterfaceonly:=True
            End If
            
            If Err.Number > 0 Then
                Call ShowMessage("System", "protectionError", smError, Err.Description)
                Exit Sub
            End If
        End If
    Next sh
    
End Sub
Function DoUnProtectWorkBook() As Boolean
' Unprotects the active workbook
    Dim pw As String
    
    If ActiveWorkbook.ProtectStructure Then
            
        On Error Resume Next
        
        pw = GetBookProtectionPassword
    
        If pw > c_noBookPWrequired Then
            If Len(pw) = 0 Then pw = AskForPassWord(confirm:=False)
            ActiveWorkbook.Unprotect Password:=pw
        Else
            ActiveWorkbook.Unprotect
        End If
        
        If Err.Number > 0 Then
            Call ShowMessage("System", "protectBookError", smError, ActiveWorkbook.Name, Err.Description)
            Exit Function
        End If
        
        DoUnProtectWorkBook = True
        
    End If
    
End Function

Function DoProtectWorkBook() As Boolean
' Protects the active workbook
    Dim pw As String

    If Not ActiveWorkbook.ProtectStructure Then
            
        On Error Resume Next
        
        pw = GetBookProtectionPassword
            
        If pw <> c_noBookPWrequired Then
        
            If Len(pw) = 0 Then
                pw = AskForPassWord
                If Len(pw) = 0 Then
                    Call ShowMessage("System", "functionaborted", smInfo, "Protect workbook")
                    Exit Function
                End If
            End If
            
            ActiveWorkbook.Protect Structure:=True, Windows:=False, Password:=pw
        Else
            ActiveWorkbook.Protect Structure:=True, Windows:=False
        End If
        
        If Err.Number > 0 Then
            Call ShowMessage("System", "protectBookError", smError, ActiveWorkbook.Name, Err.Description)
            Exit Function
        End If
        
        DoProtectWorkBook = True
        
    Else
    
        DoProtectWorkBook = True
        
    End If
    
End Function
Sub DoProtect(sh As Worksheet)
' Protects a sheet
    Dim pw As String

    If Not sh.ProtectContents Then
    
        On Error Resume Next
        Err.Clear

        pw = GetSheetProtectionPassword(sh)
                        
        If pw <> c_noPWrequired Then
        
            If Len(pw) = 0 Then
                pw = AskForPassWord
                If Len(pw) = 0 Then
                    Call ShowMessage("System", "functionaborted", smInfo, "Protect sheet")
                    Exit Sub
                End If
            End If
            
            sh.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=False, AllowSorting:=True, AllowFiltering:=True, _
                AllowFormattingColumns:=True, AllowFormattingRows:=True, userinterfaceonly:=True
        Else
            sh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=False, AllowSorting:=True, AllowFiltering:=True, _
                AllowFormattingColumns:=True, AllowFormattingRows:=True, userinterfaceonly:=True
        End If
        
        If Err.Number > 0 Then
            Call ShowMessage("System", "protectSheetError", smError, sh.Name, Err.Description)
            Exit Sub
        End If
            
    End If

End Sub
Function DoUnProtect(sh As Worksheet) As Boolean
' UnProtects a sheet
    Dim pw As String
    
    If sh.ProtectContents Then
        On Error Resume Next
        Err.Clear
        
        pw = GetSheetProtectionPassword(sh)
                        
        If pw <> c_noPWrequired Then
            If Len(pw) = 0 Then pw = AskForPassWord(confirm:=False)
            sh.Unprotect Password:=pw
        Else
            sh.Unprotect
        End If
        
        If Err.Number > 0 Then
            Call ShowMessage("System", "protectSheetError", smError, sh.Name, Err.Description)
            Exit Function
        End If
        
        DoUnProtect = True
        
    End If

End Function
Function checkNamedReferencesToWorkbook(inBook As Workbook) As Boolean
'Checks if references to ranges in a workbook (usually in this add-in) are valid and corrects them if not
'This is necessary to prevent isssues when file is moved/copied as the the refersTo will swith to fully qualified (with path)
'   instead of just using the workbook name, which works when the file is open

    Dim dummy As Variant
    Dim nm As Name
    Dim rg As Range
    Dim str As Variant
    Dim sheetName As String, rangeName As String
    Dim correctionsMade As Boolean, eventState As Boolean
    
    On Error Resume Next

    For Each nm In ActiveWorkbook.Names
        If InStr(nm.RefersTo, inBook.Name) > 0 Then
            Err.Clear
            Set rg = nm.RefersToRange
            If Err.Number > 0 Then
                str = Split(nm.RefersTo, "]")
                If UBound(str) = 1 Then                             'Scope of range is worksheet
                    str = Split(str(1), "!")
                    If UBound(str) = 1 Then
                        sheetName = Replace(str(0), "'", "")
                        rangeName = str(1)
                        Err.Clear
                        nm.RefersTo = "=[" & inBook.Name & "]" & sheetName & "!" & rangeName
                        correctionsMade = True
                    End If
                Else                                               'Scope of range is workbook
                    str = Split(nm.RefersTo, "!")
                    If UBound(str) = 1 Then
                        rangeName = str(1)
                        Err.Clear
                        nm.RefersTo = inBook.Name & "!" & rangeName
                        correctionsMade = True
                    End If
                
                End If
                Err.Clear
                Set rg = nm.RefersToRange
                If rg Is Nothing Then
                    If g_NLSData.isSet Then
                        Call ShowMessage("System", "invalidReference", smSystemError, , inBook.Name)
                    Else
                        MsgBox "Fatal Error. Invalid reference to range <" & nm.Name & "> in workbook <" & inBook.Name & ">. . Pls. do not proceed and contact the developer.", vbCritical + vbOKOnly, "Fatal Error"
                    End If
                End If
            End If
        End If
    Next nm
    
    If correctionsMade Then
        eventState = Application.EnableEvents
        Application.EnableEvents = False
        ActiveWorkbook.Save
        Application.EnableEvents = eventState
    End If
    
    checkNamedReferencesToWorkbook = True
    
End Function
Sub SetSeparators()
    Dim sep As String
    With Application
    
        On Error GoTo ErrorExit
        
        sep = Range("DecimalSeparator").Value2
        If Len(sep) = 0 Then GoTo ErrorExit
        .DecimalSeparator = sep
        
        sep = Range("ThousandsSeparator").Value2
        If Len(sep) = 0 Then GoTo ErrorExit
        .ThousandsSeparator = sep
        
        .UseSystemSeparators = False
    
    Exit Sub
    
ErrorExit:
        .UseSystemSeparators = True
    End With
    
End Sub
