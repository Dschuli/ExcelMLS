VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NLSTableEditClass"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public WithEvents NLSTextBox As MSForms.TextBox
Attribute NLSTextBox.VB_VarHelpID = -1


Private Sub NLSTextBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim txt As String, newText As String
    Dim pos As Long, length As Long
    
    If Shift = 2 Then                   'Ctrl key was pressed
        
        Select Case KeyCode
        
            Case Asc("1"), Asc("2"), Asc("3"), Asc("4")
                txt = NLSTextBox.Value
                pos = NLSTextBox.SelStart
                length = NLSTextBox.SelLength
                If pos > 0 Then
                    newText = Left(txt, pos)
                    txt = Mid(txt, pos + 1)
                End If
                newText = newText & "__&" & Chr(KeyCode) & "__"
                If length > 0 Then txt = Mid(txt, length + 1)
                NLSTextBox.Value = newText & txt
                NLSTextBox.SelStart = pos + 6
        End Select
        
    End If
    
End Sub
