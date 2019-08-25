VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NLSTableEdit 
   OleObjectBlob   =   "NLSTableEdit.frx":0000
   Caption         =   "NLS Table"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18825
   TypeInfoVer     =   288
End
Attribute VB_Name = "NLSTableEdit"
Attribute VB_Base = "0{66670D55-FA39-4355-9B9E-A5EBEBB7B3F3}{D5A17297-DFB1-4381-9DA8-07B44CC9844B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

    Option Explicit
    
    Dim TBArray() As New NLSTableEditClass
    Dim levelIdentifiers As Variant
    Dim platFormDeveloper As Boolean, reActivate As Boolean, isActivated As Boolean
    Dim entryKey As String
    Dim editCtrl As MSForms.control
    
    Public colWidths As String, entryModule As String, entryIdentifier As String
    Public func As nlsTableAction
    Public colCnt As Long, mainLanguage As Long, currentRecNo As Long
    Public isLocal As Boolean
    Public cols
Private Sub B_Act_Click()
'Do actions related to the Act Button
    Dim tb As ListObject
    Dim i As Long, row As Long, j As Long, recNo As Long
    Dim key As String, level As String, module As String, identifier As String
    Dim modeChanged As Boolean, isLocal As Boolean
    Dim keyStatus As nlsKeyStatus
    Dim txt As Variant
    Dim rg As Range

    Select Case func
        
        Case ntaAdd
            'Button not active
        
        Case ntaClone
            'Button not active
            
        Case ntaEditClone
        
            With g_NLSData
    
                level = trim(Me.CB_E_Level.Value)
                If Len(level) = 0 Then Exit Sub                                  'Saveguards - button should be disabled in this case anyhow
                module = trim(Me.CB_E_Module.Value)
                identifier = trim(Me.TB_E_Identifier.Value)
        
                key = module & identifier
        
                keyStatus = getNLSKeyStatus(module, identifier)
                
                If keyStatus = nksInvald Then Exit Sub                            'Saveguards - button should be disabled in this case anyhow
                
                If level = .platFormLevelIdentifier Then
                    If Not platFormDeveloper Or keyStatus <> nksPlt Then Exit Sub
                Else
                    If Not .hasLocalText Or keyStatus <> nksAdd Then Exit Sub
                End If
                
                If level = .appLevelIdentifier Then                                 'Get table to act on
                    Set tb = GetObjectOfType("LocalNLSText", objtListObject)
                    isLocal = True
                Else
                    Set tb = GetObjectOfType("NLS_Text", objtListObject, , ThisWorkbook)
                End If
                
'                modeChanged = DisableEventsAndScreenUpdate
                
                If isLocal Then                                         'Get recNo
                    recNo = FindRowSorted(.localText, 1, key)
                Else
                    recNo = FindRowSorted(.text, 1, key)
                End If
                
                If recNo = 0 Then Exit Sub
                
                Set rg = tb.ListRows(recNo).Range
                
                For i = 1 To colCnt
                    If i < colCnt Then
                        rg.Cells(1, g_NLSData.sysCols + i) = Me.Controls(g_NLSData.editFieldnames(i)).Value
                    Else
                        For j = 1 To .languageCount
                            rg.Cells(1, g_NLSData.offSet + mainLanguage + j - 1) = Me.Controls(g_NLSData.editFieldnames(i) & "_" & j).Value
                        Next j
                    End If
                Next i
        
'                If modeChanged Then RestoreEventsAndScreenUpdate
        
                Call SetNLSData(True)
                
                Call NLSTable.SaveFilterValues
        
                Call NLSTable.SetFilteredList(LCase(level & key), isLocal, recNo)
        
            End With
        
        Case ntaDelete
            With g_NLSData

                level = Me.CB_E_Level.Value
                If Len(level) = 0 Then Exit Sub                                  'Saveguards - button should be disabled in this case anyhow
                module = Me.CB_E_Module.Value
                identifier = Me.TB_E_Identifier.Value
                
                key = module & identifier
        
                keyStatus = getNLSKeyStatus(module, identifier)
                
                If keyStatus = nksInvald Or keyStatus = nksNew Then Exit Sub      'Saveguards - button should be disabled in this case anyhow
                
                If level = .platFormLevelIdentifier Then
                    If Not platFormDeveloper Then Exit Sub
                Else
                    If Not .hasLocalText Then Exit Sub
                End If
                
                If level = .appLevelIdentifier Then                                 'Get table to act on
                    Set tb = GetObjectOfType("LocalNLSText", objtListObject)
                    txt = .localText
                    isLocal = True
                Else
                    Set tb = GetObjectOfType("NLS_Text", objtListObject, , ThisWorkbook)
                    txt = .text
                End If
                
'                modeChanged = DisableEventsAndScreenUpdate
                
                recNo = FindRowSorted(txt, 1, key)
                If recNo = 0 Then
                    Call ShowMessage("NLS", "missingEntry", smSystemError)
                    Exit Sub
                End If
                
                tb.ListRows(recNo).Delete
                        
'                If modeChanged Then RestoreEventsAndScreenUpdate
        
                Call SetNLSData(True)
                
                Call NLSTable.SaveFilterValues
                
                Call NLSTable.SetFilteredList
        
            End With
 
        Case Else
        
    End Select
            
    Me.Hide
    
    If Me.CB_EditNext And func = ntaEditClone Then                      'Chain edit mode is active - so edit next record in list
        With NLSTable
            If .C_NLSList.ListIndex < .C_NLSList.ListCount - 1 Then
                .C_NLSList.ListIndex = .C_NLSList.ListIndex + 1
                Application.OnTime Now(), "NLSTableEditEntry", , True
            Else
                Me.CB_EditNext = False
            End If
        End With
    End If
    
End Sub

Private Sub B_Cancel_Click()
    Me.CB_EditNext = False
    Me.Hide
End Sub

Private Sub CB_E_Level_Change()
    Call ReactOnKeyState
End Sub

Private Sub CB_E_Module_Change()
    Call ReactOnKeyState
End Sub



Private Sub TB_E_Identifier_Change()
    Call ReactOnKeyState
End Sub

Private Sub B_SaveNew_Click()
    Dim tb As ListObject
    Dim newRow As ListRow
    Dim rg As Range
    Dim i As Long, row As Long, j As Long, recNo As Long
    Dim key As String, level As String, module As String, identifier As String
    Dim modeChanged As Boolean, isLocal As Boolean
    Dim keyStatus As nlsKeyStatus

    With g_NLSData
    
        level = trim(Me.CB_E_Level.Value)
        If Len(level) = 0 Then Exit Sub                                  'Saveguards - button should be disabled in this case anyhow
        module = trim(Me.CB_E_Module.Value)
        identifier = trim(Me.TB_E_Identifier.Value)

        key = module & identifier

        keyStatus = getNLSKeyStatus(module, identifier)
        
        If keyStatus = nksInvald Or keyStatus = nksBoth Then Exit Sub      'Saveguards - button should be disabled in this case anyhow
        
        If level = .platFormLevelIdentifier Then
            If Not platFormDeveloper Or keyStatus = nksPlt Then Exit Sub
        Else
            If Not .hasLocalText Or keyStatus = nksAdd Then Exit Sub
        End If
        
        If level = .appLevelIdentifier Then                                 'Get table to act on
            Set tb = GetObjectOfType("LocalNLSText", objtListObject)
            isLocal = True
        Else
            Set tb = GetObjectOfType("NLS_Text", objtListObject, , ThisWorkbook)
        End If
        
'        modeChanged = DisableEventsAndScreenUpdate
        
        Set newRow = tb.ListRows.Add
        Set rg = newRow.Range
        
        For i = 1 To colCnt
            If i < colCnt Then
                rg.Cells(1, g_NLSData.sysCols + i) = Me.Controls(g_NLSData.editFieldnames(i)).Value
            Else
                For j = 1 To .languageCount
                    rg.Cells(1, g_NLSData.offSet + mainLanguage + j - 1) = Me.Controls(g_NLSData.editFieldnames(i) & "_" & j).Value
                Next j
            End If
        Next i

'        If modeChanged Then RestoreEventsAndScreenUpdate

        Call SetNLSData(True)
        
        Call NLSTable.SaveFilterValues
        
        If isLocal Then                                         'Get recNo
            recNo = FindRowSorted(.localText, 1, key)
        Else
            recNo = FindRowSorted(.text, 1, key)
        End If

        Call NLSTable.SetFilteredList(LCase(level & key), isLocal, recNo)

    End With

    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Dim nextPos As Single, editLeft As Single, width As Single, top As Single, buttonTop As Single
    Dim i As Long, j As Long
    Dim editCtrl As MSForms.control, language As MSForms.TextBox, ctl As MSForms.TextBox, firstCtl As MSForms.TextBox
    Dim colWidth As Variant
    Dim Clear As Boolean
    
    isActivated = False                                               'Flag to indicated activation in progress
    
    platFormDeveloper = IsPlatformDeveloper
    
    colWidth = Split(colWidths, ";")
    
    levelIdentifiers = GetObjectOfType("NLSLevelIdentifier", objtRange, , ThisWorkbook).Value2
    
    editLeft = 3
    
    nextPos = editLeft
    
    Select Case func
        
        Case ntaAdd
            Me.caption = "Add NLS table entry"
            Me.B_Act.Visible = False
            Me.B_SaveNew.Visible = True
            Clear = True
            Me.CB_E_Level.Locked = False
            Me.CB_E_Module.Locked = False
            Me.TB_E_Identifier.Locked = False
            Me.CB_EditNext.Visible = False
        
        Case ntaClone
            Me.caption = "Clone NLS table entry"
            Me.B_Act.Visible = False
            Me.B_SaveNew.Visible = True
            Me.CB_E_Level.Locked = False
            Me.CB_E_Module.Locked = False
            Me.TB_E_Identifier.Locked = False
            Me.CB_EditNext.Visible = False
            
        Case ntaEditClone
            Me.caption = "Edit/Clone NLS table entry"
            Me.B_Act.caption = "Save"
            Me.B_Act.Visible = True
            Me.B_SaveNew.Visible = True
            Me.CB_E_Level.Locked = False
            Me.CB_E_Module.Locked = False
            Me.TB_E_Identifier.Locked = False
            Me.CB_EditNext.Visible = True
            
        Case ntaDelete
            Me.caption = "Delete NLS table entry"
            Me.B_Act.caption = "Delete"
            Me.B_Act.Visible = True
            Me.B_SaveNew.Visible = False
            Me.CB_E_Level.Locked = True
            Me.CB_E_Module.Locked = True
            Me.TB_E_Identifier.Locked = True
            Me.CB_EditNext.Visible = False
            
        Case Else
            Me.caption = "??????"
            Me.B_Act.caption = "?????"
            Me.B_SaveNew.caption = "?????"
            Me.B_Act.Visible = True
            Me.B_SaveNew.Visible = True
            Me.CB_EditNext.Visible = False
        
    End Select
        
    With g_NLSData
    
        ReDim TBArray(1 To .languageCount)                                                  'Prepare Array for use of classs file NLSTableEdit
        
        For i = 1 To colCnt
        
            If i = colCnt Then                                                              'Set field width - last one is remainder
                Set firstCtl = Me.Controls(.editFieldnames(i) & "_1")
                Set TBArray(1).NLSTextBox = firstCtl                                                   'Implement class file behaviour
                Set language = Me.Controls("TB_Language_1")
                language.Left = nextPos
                nextPos = nextPos + language.width
                language.text = .header(1, .offSet)
                firstCtl.Left = nextPos
                width = editLeft + Me.width - nextPos - 18
                top = firstCtl.top
                firstCtl.width = width
                                      
                Me.Controls("TB_Language_1").font.Bold = (mainLanguage = 0)
                
                Me.height = .languageCount * (editCtrl.height + 3) + 75
                                              
                If Clear Or currentRecNo = 0 Then                                   'Clear or set value
                    firstCtl.Value = ""
                Else
                    If isLocal Then
                        firstCtl.Value = .localText(currentRecNo, .offSet)
                    Else
                        firstCtl.Value = .text(currentRecNo, .offSet)
                    End If
                End If
                
                For j = 2 To .languageCount
                            
                    If reActivate Then
                        Set ctl = Me.Controls("TB_Language_" & j)
                    Else
                        Set ctl = Me.Controls.Add("Forms.TextBox.1", "TB_Language_" & j)
                        Call CopyControlProperties(language, ctl)
                    End If
                    
                    ctl.Left = language.Left
                    top = top + ctl.height + 2
                    ctl.top = top
                    ctl.text = .header(1, j + .offSet - 1)
                    Me.Controls(ctl.Name).font.Bold = (mainLanguage = j - 1)
                     
                    If reActivate Then
                        Set ctl = Me.Controls(.editFieldnames(i) & "_" & j)
                    Else
                        Set ctl = Me.Controls.Add("Forms.TextBox.1", .editFieldnames(i) & "_" & j)
                        Call CopyControlProperties(firstCtl, ctl)
                    End If
                    
                    Set TBArray(j).NLSTextBox = ctl                                    'Implement class file behaviour
                    
                    ctl.top = top
                    
                    If Clear Or currentRecNo = 0 Then                               'Clear or set value
                        ctl.Value = ""
                    Else
                        If isLocal Then
                            ctl.Value = .localText(currentRecNo, .offSet + j - 1)
                        Else
                            ctl.Value = .text(currentRecNo, .offSet + j - 1)
                        End If
                    End If
                    
                Next j
            Else
                Set editCtrl = Me.Controls(.editFieldnames(i))
                editCtrl.Left = nextPos
                editCtrl.width = colWidth(i - 1)
                                 
                If Clear Or currentRecNo = 0 Then                               'Clear or set value
                    editCtrl.Value = ""
                Else
                    If isLocal Then
                        editCtrl.Value = .localText(currentRecNo, NLSTable.cols(i - 1))
                    Else
                        editCtrl.Value = .text(currentRecNo, NLSTable.cols(i - 1))
                    End If
                End If
            End If
            
            If i = 1 Then                                                       'Handle level fields according to user level and local text availability
                editCtrl.Clear
                If platFormDeveloper Then editCtrl.AddItem .platFormLevelIdentifier
                If .hasLocalText Then editCtrl.AddItem .appLevelIdentifier
                 
                If editCtrl.ListCount > 1 Then editCtrl.Enabled = True
                
                If Clear Or currentRecNo = 0 Then                               'Clear or set value
                    editCtrl.Value = ""
                    On Error Resume Next
                    editCtrl.Value = NLSTable.C_NLSList.list(NLSTable.C_NLSList.ListIndex, 0)
                    On Error GoTo 0
                Else
                    If isLocal Then
                        editCtrl.Value = .localText(currentRecNo, NLSTable.cols(0))
                    Else
                        editCtrl.Value = .text(currentRecNo, NLSTable.cols(0))
                    End If
                End If
                                
            End If
                
            nextPos = nextPos + editCtrl.width
            
        Next i
        
    End With
    
    buttonTop = Me.height - (35 + Me.B_Cancel.height)
    
    Me.B_Cancel.top = buttonTop
    Me.B_Act.top = buttonTop
    Me.B_SaveNew.top = buttonTop
    Me.CB_EditNext.top = buttonTop
    
    reActivate = True
    isActivated = True
    
    entryKey = Me.CB_E_Level.Value & Me.CB_E_Module.Value & Me.TB_E_Identifier.Value                'save key values at entry time
    
    If Len(entryModule) > 0 Or Len(entryIdentifier) > 0 Then                                        'Entry Values were passed
        Me.CB_E_Module.Value = entryModule
        Me.TB_E_Identifier.Value = entryIdentifier
        entryModule = ""
        entryIdentifier = ""
    End If
    
    Call ReactOnKeyState
    
End Sub


Sub ReactOnKeyState()
    Dim keyStatus As nlsKeyStatus
    Dim level As String, module As String, identifier As String
    
    If Not reActivate Then Exit Sub
    
    level = trim(Me.CB_E_Level.Value)
    module = trim(Me.CB_E_Module.Value)
    identifier = trim(Me.TB_E_Identifier.Value)

    If Len(level) = 0 Then
        keyStatus = nksInvald
    Else
        keyStatus = getNLSKeyStatus(Me.CB_E_Module.Value, Me.TB_E_Identifier)
    End If
    
    With g_NLSData
    
        Select Case keyStatus
            
            Case nksInvald                                      'Invalid Key - no action allowed
                Me.L_KeyState = "Invalid/Undefined record"
                Me.B_Act.Enabled = False
                Me.B_SaveNew.Enabled = False
                
            Case nksNew                                         'New key - only save new is allowed
                Me.L_KeyState = "New record"
                Me.B_Act.Enabled = False
                If level = .platFormLevelIdentifier Then
                    Me.B_SaveNew.Enabled = platFormDeveloper
                Else
                    Me.B_SaveNew.Enabled = .hasLocalText
                End If
            
            Case nksPlt
                Me.L_KeyState = "Record exists in platform NLS table"
                Me.B_Act.Enabled = platFormDeveloper
                If level = .platFormLevelIdentifier Then
                    Me.B_SaveNew.Enabled = False
                Else
                    Me.B_SaveNew.Enabled = .hasLocalText
                End If
                
            Case nksAdd
                Me.L_KeyState = "Record exists in private NLS table"
                Me.B_Act.Enabled = True
                If level = .platFormLevelIdentifier Then
                    Me.B_SaveNew.Enabled = platFormDeveloper
                Else
                    Me.B_SaveNew.Enabled = False
                End If
                                
            Case nksBoth
                Me.L_KeyState = "Record exists in both NLS tables"
                Me.B_Act.Enabled = True
                Me.B_SaveNew.Enabled = False
                
        End Select
      
    End With
    
    If func = ntaDelete Then Me.L_KeyState = "Do you really want to delete this entry?"                  'Display Delete message
    
    If func = ntaEditClone Then Me.B_Act.Enabled = level & module & identifier = entryKey                'Disable Act button (save) if key was changed

End Sub

Private Sub UserForm_Initialize()

    With NLSTable           'set initial position of window
        Me.top = .top + .F_Filter.top + 35
        Me.Left = .Left + .F_Filter.Left + 20
    End With

End Sub
