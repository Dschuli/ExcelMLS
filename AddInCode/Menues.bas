Attribute VB_Name = "Menues"
Option Explicit

'************ Enumerations ****************************
Public Enum appArea                 'Actions available for NSL entries
    aaOther
    aaInfo
End Enum

'************ Datatypes ****************************
Type cellMenuElement                                        'Data structure for cellMenu elements
    caption As String
    action As String
    parameter As Variant
    faceId As Long
    beginGroup As Boolean
End Type

Sub Dm(Optional dummy As Boolean)
    If Not (IsDeveloper Or IsPlatformDeveloper) Then Exit Sub
    Call CreateDisplayDevelopmentPopUpMenu
End Sub
Sub Nt(Optional dummy As Boolean)
    If Not (IsDeveloper Or IsPlatformDeveloper) Then Exit Sub
    Call ShowNlsTable
End Sub


Sub CreateDisplayDevelopmentPopUpMenu(Optional dummy As Boolean)
    Dim menuName As String
    
    menuName = "Test"
    
    'Delete PopUp menu if it exist
    On Error Resume Next
    Application.CommandBars(menuName).Delete
    On Error GoTo 0

    'Create the PopUpmenu
    Call DevelopmentPopUpMenu(menuName)

    'Show the PopUp menu
    On Error Resume Next
    Application.CommandBars(menuName).ShowPopup
    On Error GoTo 0
End Sub

Sub DevelopmentPopUpMenu(Optional menuName As String, Optional parentMenu As Variant)
    'Add PopUp menu for development  - Used t,n
    Dim MenuItem As CommandBarPopup
    
    If IsMissing(parentMenu) Then Set parentMenu = Application.CommandBars.Add(Name:=menuName, position:=msoBarPopup, MenuBar:=False, Temporary:=True)
    
    With parentMenu

        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .caption = "&NLS"
            
            With .Controls.Add(Type:=msoControlButton)
                .caption = "Show NLS&Table"
                .faceId = 584
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "showNlsTable"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .caption = "Set &Language"
                .faceId = 274
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "Setlanguage"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .caption = "Show NLS&Calls"
                .faceId = 371
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "ShowNlsCalls"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .caption = "Show &unused NLS table entries"
                .faceId = 533
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "ShowWidows"
            End With
            
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .caption = "&ToggleAddIn"
            .faceId = 184
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "ToggleAddIn"
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .caption = "&Reset Mode"
            .faceId = 157
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "resetFastMode"
        End With
        
    End With
End Sub
Sub ShowCustomizedCellContextMenu(sh As Worksheet, target As Range)
'Change cellContextMenus
    Dim contextMenu As CommandBar
    Dim ctl As CommandBarControl
    Dim cellMenuElements() As cellMenuElement
    Dim menuElement As Variant
    
    Dim faceId() As Long, elements As Long, count As Long, i As Long
    Dim fidStr As String, context As String
    Dim pt As PivotTable
    
    If Not target.ListObject Is Nothing Then                        'Get context and set ContextMenu accordingly
        context = "List Range Popup"                                    'Listobject
    Else
        On Error Resume Next
        Set pt = target.PivotTable
        If Err.Number = 0 Then
            context = "PivotTable Context Menu"                         'PivotTable
        Else
            If target.Address = target.EntireColumn.Address Then
                context = "Column"
            Else
                If target.Address = target.EntireRow.Address Then context = "Row" Else context = "Cell"
            End If
        End If
    End If
       
    On Error GoTo 0
    
    Set contextMenu = Application.CommandBars(context)
    
    contextMenu.Reset
    
    If sh.Name = c_infoSheetName Then
        If target.CurrentRegion.row = 3 And target.row > 3 And Len(target.Value2) > 0 Then      'Assured the current cell is in the data range
            If InStr(ActiveSheet.Cells(1, 1).Value2, "NLS") > 0 Then                                    'Assure the title (Range A1) includes NLS
                count = count + 1
                Call SetMenuElement(contextMenu:=contextMenu, functionName:="showNlsTable", position:=count)
                If target.CurrentRegion.Cells(1, 2) = "Area" Then
                    count = count + 1
                    Call SetMenuElement(contextMenu:=contextMenu, functionName:="LocateNLSCall", position:=count)
                End If
            End If

        End If
    End If
    
    If IsDeveloper Then
        count = count + 1
        Set menuElement = SetMenuElement(contextMenu, "DevMenu", msoControlPopup, count)
        Call DevelopmentPopUpMenu("", menuElement)
        contextMenu.Controls(count).beginGroup = True
    End If


'     Add a separator to the Cell context menu.
    contextMenu.Controls(count + 1).beginGroup = True
      
    contextMenu.ShowPopup
    
    contextMenu.Reset
    
End Sub
Function SetMenuElement(contextMenu As CommandBar, functionName As String, Optional controlType As MsoControlType = msoControlButton, Optional position As Long = 1, Optional parameter As Variant, Optional beginGroup As Boolean) As Variant
'Fills a cellMenuElement with values based on functionName
    Dim fidStr As String

    Set SetMenuElement = contextMenu.Controls.Add(Type:=controlType, before:=position)
    
    With SetMenuElement

        .caption = GetNlsText(c_moduleForMenus, functionName)
        
        If controlType = msoControlButton Then
            .OnAction = functionName
            fidStr = GetNlsAddition(c_moduleForMenus, functionName)
            If IsNumeric(fidStr) Then .faceId = CLng(fidStr)
        End If

        .beginGroup = beginGroup

        On Error Resume Next
        If Not IsMissing(parameter) Then .parameter = parameter

    End With

End Function
Sub FaceIdsToolBar(Optional dummy As Boolean)
' Shows all valid faceIds in menu Add-Ins
Dim sym_bar As CommandBar
Dim cmd_bar As CommandBar
' =========================
Dim i_bar As Integer
Dim n_bar_ammt As Integer
Dim i_bar_start As Integer
Dim i_bar_final As Integer
' =========================
Dim icon_ctrl As CommandBarControl
' =========================
Dim i_icon As Integer
Dim n_icon_step As Integer
Dim i_icon_start As Integer
Dim i_icon_final As Integer
' =========================
n_icon_step = 10
' =========================
i_bar_start = 1
n_bar_ammt = 500
' i_bar_start = 501
' n_bar_ammt =  1000
' i_bar_start = 1001
' n_bar_ammt =  1500
' i_bar_start = 1501
' n_bar_ammt =  2000
' i_bar_start = 2001
' n_bar_ammt =  2543
i_bar_final = i_bar_start + n_bar_ammt - 1
' =========================
' delete toolbars
' =========================
For Each cmd_bar In Application.CommandBars
    If InStr(cmd_bar.Name, "Symbol") <> 0 Then
        cmd_bar.Delete
    End If
Next
' =========================
' create toolbars
' =========================
For i_bar = i_bar_start To i_bar_final
    On Error Resume Next
    Set sym_bar = Application.CommandBars.Add _
        ("Symbol" & i_bar, msoBarFloating, Temporary:=True)
    ' =========================
    ' create buttons
    ' =========================
    i_icon_start = (i_bar - 1) * n_icon_step + 1
    i_icon_final = i_icon_start + n_icon_step - 1
    For i_icon = i_icon_start To i_icon_final
        Set icon_ctrl = sym_bar.Controls.Add(msoControlButton)
        icon_ctrl.faceId = i_icon
        icon_ctrl.TooltipText = i_icon
    Next i_icon
    sym_bar.Visible = True
Next i_bar
End Sub
Sub DeleteFaceIdsToolbar(Optional dummy As Boolean)
' Removes the faceId add-in
Dim cmd_bar As CommandBar
For Each cmd_bar In Application.CommandBars
    If InStr(cmd_bar.Name, "Symbol") <> 0 Then
        cmd_bar.Delete
    End If
Next
End Sub

Sub Add_Name_To_Contextmenus(Optional dummy As Boolean)
    Dim Cbar As CommandBar
    For Each Cbar In Application.CommandBars
        With Cbar
            If .Type = msoBarTypePopup Then
                On Error Resume Next
                With .Controls.Add(Type:=msoControlButton)
                    .caption = "Name for VBA = " & Cbar.Name
                    .Tag = "NameButtonInContextMenu"
                End With
                On Error GoTo 0
            End If
        End With
    Next
End Sub


Sub Delete_Name_From_Contextmenus(Optional dummy As Boolean)
    Dim Cbar As CommandBar
    Dim ctrl As CommandBarControl

    For Each Cbar In Application.CommandBars
        With Cbar
            If .Type = msoBarTypePopup Then
                For Each ctrl In .Controls
                    If ctrl.Tag = "NameButtonInContextMenu" Then
                        ctrl.Delete
                    End If
                Next ctrl
            End If
        End With
    Next
End Sub
Sub testMenu()
    Dim macroName As String, fidStr As String
    Dim parentMenu As CommandBar
    Dim menuName As String
    
    menuName = "Test"
    
    On Error Resume Next
    Application.CommandBars(menuName).Delete
    On Error GoTo 0
    
    Set parentMenu = Application.CommandBars.Add(Name:="Test", position:=msoBarPopup, MenuBar:=False, Temporary:=True)
    
    With parentMenu
    
        With .Controls.Add(Type:=msoControlButton)
            macroName = "TestMacro"
            .caption = GetNlsText(c_moduleForMenus, macroName)
            fidStr = GetNlsAddition(c_moduleForMenus, macroName)
            If IsNumeric(fidStr) Then .faceId = CLng(fidStr)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
        End With
        
    End With
    
    parentMenu.ShowPopup
    
    parentMenu.Delete

End Sub

Sub TestMacro()
    Call ShowMessage("Test", "Macro", smInfo)
End Sub
