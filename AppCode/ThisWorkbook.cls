VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Activate()
    Call BookActivate
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call bookBeforeClose
End Sub

Private Sub Workbook_Deactivate()
    Call BookDeactivate
End Sub

Private Sub Workbook_Open()
    Call BookOpen
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
    Cancel = SheetBeforeRightClick(sh, Target)
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    Call SheetChange(sh, Target)
End Sub
