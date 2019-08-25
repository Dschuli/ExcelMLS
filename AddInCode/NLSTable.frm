VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NLSTable 
   OleObjectBlob   =   "NLSTable.frx":0000
   Caption         =   "NLS Table"
   ClientHeight    =   9630.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19290
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   253
End
Attribute VB_Name = "NLSTable"
Attribute VB_Base = "0{8DD3C24A-ECEE-4656-9275-DAD32D0C5EB7}{4C5A9FAE-3329-4741-A56C-351DE6105E76}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


    Dim baselist As Variant, filteredList As Variant, headerText As Variant, rws As Variant, levelIdentifiers As Variant
    Dim platFormDeveloper As Boolean, addModus As Boolean, allowAction As Boolean
    
    Public colWidths As String, entryModule As String, entryIdentifier As String, module As String, identifier As String
    Public colCnt As Long, mainLanguage As Long
    Public wasCanceled As Boolean
    Public cols

Private Sub B_Add_Click()
    NLSTableEdit.func = ntaAdd
    NLSTableEdit.show
End Sub

Private Sub B_Cancel_Click()
    Me.Hide
    wasCanceled = True
End Sub

Private Sub B_ClearFilter_Click()
    Dim i As Long
    Dim extendedKey As String
    

    Call SaveFilterValues(True)             'Clear and save filter values
    
    With Me.C_NLSList
    
        If .ListIndex > -1 Then extendedKey = .list(.ListIndex, 0) & .list(.ListIndex, 1) & .list(.ListIndex, 2)
    
        Call SetFilteredList(LCase(extendedKey))
                
        If .ListCount > 0 Then
            If .ListIndex = -1 Then .ListIndex = 0
            .SetFocus
        End If
        
    End With
    
End Sub

Private Sub B_UseSelection_Click()
    Dim text As String
    
    With Me.C_NLSList
        If .ListIndex = -1 Then Exit Sub
        module = .list(.ListIndex, 1)
        identifier = .list(.ListIndex, 2)
    End With
        
    wasCanceled = False
    Me.Hide
End Sub

Private Sub B_Delete_Click()
    Dim saveIndex As Long
   
    With Me.C_NLSList
        
        saveIndex = .ListIndex
    
        NLSTableEdit.func = ntaDelete
        NLSTableEdit.show
    
        If .ListCount = 0 Then Call B_ClearFilter_Click
        .ListIndex = Application.Median(0, saveIndex - 1, .ListCount - 1)
    End With
    
End Sub

Private Sub B_Edit_Click()
   Call EditEntry
End Sub
Public Sub EditEntry()
    If allowAction Then NLSTableEdit.func = ntaEditClone Else NLSTableEdit.func = ntaClone
    NLSTableEdit.show
End Sub

Private Sub B_SetFilter_Click()
    Dim i As Long
    Dim extendedKey As String
    
    Call SaveFilterValues
    
    With Me.C_NLSList
        
        If .ListIndex > -1 Then extendedKey = .list(.ListIndex, 0) & .list(.ListIndex, 1) & .list(.ListIndex, 2)
        
        Call SetFilteredList(LCase(extendedKey))
        
        If .ListCount > 0 Then
            If .ListIndex = -1 Then .ListIndex = 0
            .SetFocus
        Else
            Me.LB_AllLanguages.Visible = False
        End If
        
    End With
    
End Sub
Sub SaveFilterValues(Optional Clear As Boolean)
'Saves filter values to g_nlsData - optional clear before save
    Dim ctl As MSForms.control
    Dim i As Long

    With g_NLSData
        For i = 1 To UBound(.filterFieldNames)                                     'Save Filtervalues
            Set ctl = Me.Controls(.filterFieldNames(i))
            If Clear Then ctl.Value = ""
            .filterVals(i) = ctl.Value
        Next i
    End With

End Sub
Sub SetFilteredList(Optional extendedKey As String, Optional isLocal As Boolean, Optional baseRecNo As Long)
'Puts a filtered list into the control C_NLSList
'If extended key is provided the corresponding record will get selected
'If isLocal / baseRecNo is provided it is assured that the corresponding record is visible

    Dim rws2 As Variant
    Dim i As Long, j As Long, cnt As Long, recNo As Long, Shift As Long
    Dim likeValue As String
    
    Me.C_NLSList.Clear
    
    With g_NLSData
    
        If baseRecNo > 0 Then
            If .hasLocalText And Not isLocal Then
                recNo = UBound(.localText, 1) + baseRecNo
            Else
                recNo = baseRecNo
            End If
        End If

        baselist = .combinedText
        
        rws = Application.Evaluate("TRANSPOSE(Row(1:" & UBound(baselist, 1) & "))")         'Build initial index vector
        
        For i = colCnt To 1 Step -1
        
            If Len(.filterVals(i)) > 0 Then
            
                ReDim rws2(1 To UBound(rws))
                likeValue = "*" & LCase(.filterVals(i)) & "*"
                cnt = 0
            
                For j = 1 To UBound(rws)
                        
                    If LCase(baselist(rws(j), cols(i - 1))) Like likeValue Then
                        
                        cnt = cnt + 1
                        rws2(cnt) = rws(j)
                        
                    End If
                        
                Next j
                
                If cnt = 0 Then Exit Sub
            
                ReDim Preserve rws2(1 To cnt)
                rws = rws2
                
            End If
            
        Next i
    
    End With
    
    If recNo > 0 And UBound(rws) <> UBound(baselist, 1) Then            'If list is filtered and a mandatory reco is provided, assure its visible
    
        For i = 1 To UBound(rws)
            If rws(i) = recNo Then GoTo RecNoIsVisible                      'If recNo is in the list do nothing
        Next i
        
        ReDim rws2(1 To UBound(rws) + 1)                                    'Extend and insert recNo at the correct position
        
        For i = 1 To UBound(rws)
            If rws(i) > recNo And Shift = 0 Then
                rws2(i) = recNo
                Shift = 1
            End If
            rws2(i + Shift) = rws(i)
        Next i
        If Shift = 0 Then rws2(i) = recNo
        
        rws = rws2
        
    End If
    
RecNoIsVisible:
    filteredList = spliceMatrix(baselist, rws, cols)
    Me.C_NLSList.list = filteredList
    
    If Len(extendedKey) > 0 Then
        For i = 1 To UBound(filteredList, 1)
            If LCase(filteredList(i, 1) & filteredList(i, 2) & filteredList(i, 3)) = LCase(extendedKey) Then Exit For
        Next i
    End If
    
    If Me.C_NLSList.ListCount >= i Then Me.C_NLSList.ListIndex = i - 1


End Sub

Private Sub B_AllLanguages_Click()
    With Me.LB_AllLanguages
        If .Visible Then
            .Visible = False
            Me.Repaint
        Else
            If Me.C_NLSList.ListIndex >= 0 Then
                .height = .font.Size
                .Visible = True
                Call FillAllLanguageBox
            End If
        End If
    End With
End Sub
Private Sub FillAllLanguageBox()
    Dim pos As Long, relPos As Long, maxSize As Single, langSize As Single
    Dim i As Long
    Dim top As Single
    Dim key As String
    
    langSize = 60
    
    With Me.LB_AllLanguages
    
        If Not .Visible Then Exit Sub
    
        .Clear
    
        With Me.C_NLSList
            pos = .ListIndex
            relPos = pos - .TopIndex + 1
            top = .top
        End With
        
        For i = 0 To g_NLSData.languageCount - 1
            .AddItem
            .list(i, 0) = g_NLSData.header(1, i + g_NLSData.offSet)
            Me.TB_Sizer.Value = baselist(rws(pos + 1), i + g_NLSData.offSet)
            maxSize = Application.Max(maxSize, Me.TB_Sizer.width)
            .list(i, 1) = Me.TB_Sizer.Value
        Next i
        
        .ColumnWidths = langSize & ";" & maxSize
        
        .height = Application.Min(.font.Size * g_NLSData.languageCount + 10, Me.height * 0.5)
        .top = top + relPos * .font.Size + 5
        
        If langSize + maxSize > .width Then .height = .height + 12
        
    End With
    
    
End Sub

Private Sub B_ShowUses_Click()
    wasCanceled = True
    Me.Hide
    Call ShowNlsCalls(show:=True)
    Application.OnTime Now(), "SwitchToExcel", , True
End Sub




Private Sub C_NLSList_Click()
    
    'Manage options / captions of action buttons - "Clone" is adding a new record using another record as example
    
    allowAction = platFormDeveloper Or Me.C_NLSList.list(Me.C_NLSList.ListIndex, 0) = g_NLSData.appLevelIdentifier
    
    Me.B_Add.Enabled = platFormDeveloper Or g_NLSData.hasLocalText              'Add is allowed for platform developer or if we have local text
    
    Me.B_Delete.Enabled = allowAction
    
    Me.B_Edit.Enabled = allowAction Or g_NLSData.hasLocalText
    
    If allowAction Then Me.B_Edit.caption = "Edit/Clone" Else Me.B_Edit.caption = "Clone"
    
    Call FillAllLanguageBox
    
    Call setCurrentRecordNumber
    
End Sub

Private Sub C_NLSList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Select Case KeyCode
        
        Case vbKeyL, vbKeyL
            Call B_AllLanguages_Click
    
    End Select
        
End Sub


Private Sub LB_AllLanguages_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Me.LB_AllLanguages.Visible = False
End Sub

Public Sub B_ShowWidows_Click()
'Show unused entries only
    Dim baselist As Variant, rws As Variant, rws2 As Variant
    Dim j As Long
    Dim modeChanged As Boolean
    Dim sh As Worksheet
    Dim wasProtected As Boolean
    
    Call ShowNlsCalls(show:=False)

    With g_NLSData
    
        baselist = .combinedText
        
        rws = Application.Evaluate("TRANSPOSE(Row(1:" & UBound(baselist, 1) & "))")         'Build initial index vector
        
            ReDim rws2(1 To UBound(rws))
            cnt = 0
        
            For j = 1 To UBound(rws)
                    
                If FindRow(.xRef, 8, baselist(j, 1)) = 0 Then
                    
                    cnt = cnt + 1
                    rws2(cnt) = rws(j)
                    
                End If
                    
            Next j
            
            If cnt = 0 Then
                Call ShowMessage("NLS", "noWidows", smInfo)
                Exit Sub
            End If
            ReDim Preserve rws2(1 To cnt)
            rws = rws2
                
    End With
    filteredList = spliceMatrix(baselist, rws, cols)
    
    Set sh = ActiveWorkbook.Worksheets(c_infoSheetName)
    sh.Cells.Clear
    On Error GoTo 0
        
    modeChanged = setFastMode
    
    With sh
    
        With .Cells(1, 1)
            .Cells(1).Value = "Unused NLS Table entries"
            With .font
                .Size = 16
                .Bold = True
                .color = RGB(255, 0, 0)
            End With
        End With
        
        .Cells(1, 4).Value = "Caution: Some NLS table entries, that are used by passing a a variable and not a literal to the respective call, might (incorrectly) show up as unused in here."
        .Cells(1, 4).font.Bold = True
    
        Set rg = ArrayToRange(.Cells(3, 1), Array("Level", "Module", "Identifier", "Type", "Additional", "Text"))
        .rows(3).font.Bold = True
        
        Set rg = ArrayToRange(.Cells(4, 1), filteredList)
        rg.offSet(-1).Resize(rg.rows.count + 1).AutoFilter
        
    End With
    
    If modeChanged Then Call resetFastMode
    
    If Not sh.Visible = xlSheetVisible Then
        wasProtected = DoUnProtectWorkBook
        sh.Visible = xlSheetVisible
        If wasProtected Then Call DoProtectWorkBook
    End If
        
    sh.Activate
    sh.Cells(1, 1).Select
        
    wasCanceled = True
    Me.Hide
        
    Application.OnTime Now(), "SwitchToExcel", , True
    
End Sub

Private Sub UserForm_Activate()
    Dim colWidth As Variant
    Dim nextPos As Single, filterLeft As Single
    Dim i As Long
    Dim filterCtrl As MSForms.control, editCtrl As MSForms.control
    
    NLSTableEdit.colCnt = colCnt
    NLSTableEdit.colWidths = colWidths
    NLSTableEdit.mainLanguage = mainLanguage
    
    platFormDeveloper = IsPlatformDeveloper
    
    colWidth = Split(colWidths, ";")
    
    levelIdentifiers = GetObjectOfType("NLSLevelIdentifier", objtRange, , ThisWorkbook).Value2
    
    Me.LB_AllLanguages.Visible = False
    
    filterLeft = 3
    
    nextPos = filterLeft
    
    With g_NLSData
            
        Call FillComboBoxLists(Array(2, 4))                                                      'Set comboBox Lists - currently used for fields 2 and 4 - also acts on NLSTableEdit form
                    
        For i = 1 To colCnt
        
            Set filterCtrl = Me.Controls(.filterFieldNames(i))
            
            filterCtrl.Left = nextPos
            
            If i = colCnt Then                                                              'Set field width - last one is remainder
                filterCtrl.width = filterLeft + Me.C_NLSList.width - nextPos
            Else
                filterCtrl.width = colWidth(i - 1)
            End If
            
            If i = 1 Then                                                                               'Handle level fields according to user level and local text availability
                filterCtrl.Clear
                If .hasLocalText Then
                    filterCtrl.Enabled = True
                    filterCtrl.list = levelIdentifiers
                    filterCtrl.Value = .filterVals(i)
                    B_Add.Enabled = True
                Else
                    filterCtrl.Enabled = False
                    filterCtrl.Value = ""
                    B_Add.Enabled = platFormDeveloper
                End If
            Else                                                                            'Use saved filter values
                On Error Resume Next
                filterCtrl.Value = .filterVals(i)
                On Error GoTo 0
            End If
            
            nextPos = nextPos + filterCtrl.width
            
        Next i
        
    End With
    
    headerText = TwoDtoOneD(spliceMatrix(g_NLSData.header, Array(1), cols))                         'Configure and fill header box
    With Me.C_Header
        .columnCount = colCnt
        .ColumnWidths = colWidths
        .AddItem
        For i = 0 To UBound(headerText)
            .list(0, i) = headerText(i)
        Next i
    End With

    With Me.C_NLSList                                                                               'Configure fill message list
        .columnCount = colCnt
        .ColumnWidths = colWidths
        
        Select Case getNLSKeyStatus(entryModule, entryIdentifier)
        
            Dim a As nlsKeyStatus
            
            Case nksInvald
                Call SetFilteredList
                
            Case nksNew
                Call SetFilteredList
                If ShowConfirm("NLS", "addNewRecord", cfInfo) Then
                    NLSTableEdit.entryModule = entryModule
                    NLSTableEdit.entryIdentifier = entryIdentifier
                    Call B_Add_Click
                End If
                
            Case nksAdd, nksBoth
                Call SetFilteredList(g_NLSData.appLevelIdentifier & entryModule & entryIdentifier)
                
            Case nksPlt
                Call SetFilteredList(g_NLSData.platFormLevelIdentifier & entryModule & entryIdentifier)
                
            Case Else
                Call SetFilteredList
                
        End Select
        
        If .ListCount > 0 Then
            If .ListIndex = -1 Then .ListIndex = 0
            .SetFocus
        End If
        
    End With
'    call GetNlsText("nwe","hh")
    
End Sub
Sub setCurrentRecordNumber()
'Sets a flag if current record is local or platform and the record number

    Dim level As String, key As String
    Dim recNo As Long, listRecNo As Long
    
    listRecNo = Me.C_NLSList.ListIndex
    If listRecNo = -1 Then Exit Sub
    
    level = Me.C_NLSList.list(listRecNo, 0)
    key = LCase(Me.C_NLSList.list(listRecNo, 1) & Me.C_NLSList.list(listRecNo, 2))
        
    With g_NLSData
        
        If level = .platFormLevelIdentifier Then
            NLSTableEdit.isLocal = False
            NLSTableEdit.currentRecNo = FindRowSorted(.text, 1, key)
        Else
            NLSTableEdit.isLocal = True
            NLSTableEdit.currentRecNo = FindRowSorted(.localText, 1, key)
        End If
    
    End With
    
End Sub
Sub FillComboBoxLists(fieldNums As Variant)
'Fill combobox lists for applicable fields
    Dim filterCtrl As MSForms.control, editCtrl As MSForms.control
    Dim i As Long, fieldNum As Long
    Dim cbList As Variant
        
    With g_NLSData
    
        For i = LBound(fieldNums) To UBound(fieldNums)
        
            fieldNum = fieldNums(i)
            
            
            Set filterCtrl = Me.Controls(.filterFieldNames(fieldNum))
            Set editCtrl = NLSTableEdit.Controls(.editFieldnames(fieldNum))
            
            cbList = SortArray(Application.Index(.combinedText, , fieldNum + .sysCols), , , , , , True)
            filterCtrl.list = cbList
            editCtrl.list = cbList
                
        Next i
        
    End With
        
End Sub
