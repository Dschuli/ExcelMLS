Attribute VB_Name = "TestCode"
Option Explicit
Sub ShowHideInfoSheet()
    Dim wasProtected As Boolean
    Dim vis As XlSheetVisibility
    Dim sh As Worksheet
    Dim modeChanged As Boolean
    
    Set sh = ActiveSheet
    vis = ActiveWorkbook.Worksheets(c_NlsSheetName).Visible
    
    modeChanged = SetFastMode
    
    wasProtected = DoUnProtectWorkBook
    
    If vis = xlSheetVisible Then
        ActiveWorkbook.Worksheets(c_infoSheetName).Visible = xlSheetHidden
        ActiveWorkbook.Worksheets(c_NlsSheetName).Visible = xlSheetHidden
'        ActiveWorkbook.Worksheets(c_sortSheetname).Visible = xlSheetHidden
    Else
        ActiveWorkbook.Worksheets(c_infoSheetName).Visible = xlSheetVisible
        ActiveWorkbook.Worksheets(c_NlsSheetName).Visible = xlSheetVisible
'        ActiveWorkbook.Worksheets(c_sortSheetname).Visible = xlSheetVisible
    End If
    
    If ActiveSheet.Name <> sh.Name Then sh.Activate
    
    If wasProtected Then Call doProtectWorkbook
    
    If modeChanged Then Call resetFastMode
    
End Sub

Sub ShowUserMessage()
    Call showMessage("Test", "DemoMessage", smInfo, Application.UserName)
End Sub
