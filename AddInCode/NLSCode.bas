Attribute VB_Name = "NLSCode"
Option Explicit
Option Compare Text

'************ Enumerations ****************************
Public Enum nlsTableAction                  'Actions available for NSL entries
    ntaAdd
    ntaEditClone
    ntaClone
    ntaDelete
End Enum

Public Enum nlsKeyStatus                  'Possible stati of an NLSTable key value
    nksInvald                                   'Key is not fully defined
    nksNew                                      'Key does not exist
    nksPlt                                      'Key exists (only) in platform table
    nksAdd                                      'Key exists (only) in private table
    nksBoth                                     'Key exists in both tables
End Enum

'************ Custom Datatypes ***********************
Type NLSData                                        'Holds all NLS relevant settings & data
    isSet As Boolean
    bookName As String                              'Name of originating book
    systemLCID As String                            'Language code of excel system language
    language As Long                                'Application Language used 0 = Englisch 1 = Deutsch....
    systemLanguage As Long                          'same but for system language
    offSet As Long                                  'Offset for language 0
    languageCount As Long                           'Number of languages in table
    hasLocalText As Boolean                         'Flag that local text is used
    localText As Variant                            'local text in the active workbook - takes precedence - not used in mandatory
    text As Variant                                 'Standard text from add-in
    combinedText As Variant                         'localText + text combined
    header As Variant                               'NLS table header
    NLSTableColumnCount As Long                     'Number of columns to display in NLSTable
    
    filterVals() As String                          'for saving the showNLStable filter values
    filterFieldNames() As String                    'Filter field names in NLSTable
    editFieldnames() As String                      'Edit field names in NLSTable
    
    platFormLevelIdentifier As String               'Level identifier
    appLevelIdentifier As String
    
    sysCols As Long                                 'Number of upfront calculated columns
    levelCol As Long                                'Position of Level column
    modCol As Long                                  'Position of Module column
    identCol As Long                                'Position of Identifier column
    typeCol As Long                                 'Position of Type column
    addCol As Long                                  'Position of Addition column
     
    udfNames() As Variant                           'Names of NLS related UDFs that use the module/identifier syntax
    xRef As Variant                                 'XRef of module/identifier use
    
    parmSep As String                               'String that is used to seperate parameters in case a full parameter string is passed as module value  - should never be part of a text
    notFoundMessages As Variant                     'Base message in case entry is not found
    moduleForShapes As String                       'Module used for shape text
    useFirstLanguageIfBlank As Boolean              'Flag to control if the first language result should replace a blank entry in another language
End Type

'************ Global Variables ***********************
Public g_NLSData As NLSData


'************ Code ***********************
Sub SetNLSData(Optional force As Boolean = False)
'Sets all NLS relevant info & data
    Dim rg As Range
    Dim nlsText As ListObject
    Dim row As Long, NLSTableColumnCount As Long
    Dim cbSaved As Boolean, modeChanged As Boolean
    Dim systemLanguages As Variant
    Dim levelIdentifiers As Range
    
    
    With g_NLSData
    
        If .isSet And Not force And .bookName = ActiveWorkbook.Name Then Exit Sub
        
        .moduleForShapes = c_moduleForShapes
        .parmSep = c_parmSep
        .useFirstLanguageIfBlank = c_useFirstLanguageIfBlank
        
        .notFoundMessages = Array("Unknown message", "Unbekannte Nachricht", "Message inconnu")
        
        .bookName = ActiveWorkbook.Name
        .udfNames = Array("GetNlsText", "GetMandatoryNlsText", "ShowMessage", "ShowConfirm", "SetMenuElement")       'Names of NLS related UDFs that use the module/identifier syntax - SetMenuElement has a special treatment and has to be the last position
        .sysCols = 2
        .levelCol = .sysCols + 1
        .modCol = .levelCol + 1
        .identCol = .modCol + 1
        .typeCol = .identCol + 1
        .addCol = .typeCol + 1
        Set levelIdentifiers = GetObjectOfType("NLSLevelIdentifier", objtRange, , ThisWorkbook)
        .platFormLevelIdentifier = levelIdentifiers.Cells(1, 1).Value2
        .appLevelIdentifier = levelIdentifiers.Cells(2, 1).Value2
    End With
    
    '*********** Get text from addd-in
    Set nlsText = GetObjectOfType("NLS_Text", objtListObject, , ThisWorkbook)
    If nlsText Is Nothing Then
    
        MsgBox "Fatal Error. Cannot find range <NLS_Text>. Pls. do not proceed and contact the developer.", vbCritical + vbOKOnly, "Fatal Error"
        End
        
    End If
    
'    cbSaved = SaveClipBoard
'
'    modeChanged = DisableEventsAndScreenUpdate
'
    On Error GoTo ErrExit
        
    With nlsText
    
        .AutoFilter.ShowAllData
    
        On Error Resume Next
        .DataBodyRange.Calculate
        
        .DataBodyRange.Columns(1).Value = .DataBodyRange.Columns(2).Value2
        .DataBodyRange.Columns(3).Value = g_NLSData.platFormLevelIdentifier
        
        .Sort.SortFields.Clear
        .Range.Sort key1:=.ListColumns(1).DataBodyRange.Cells(1), _
            order1:=xlAscending, DataOption1:=xlSortNormal, _
            SortMethod:=xlPinYin, MatchCase:=False, header:=xlYes
        On Error GoTo ErrExit
            
    End With
    
    Set rg = nlsText.DataBodyRange
     
    If Not CheckDataSet(rg) Then GoTo ErrExit
 
    With g_NLSData
        
        On Error Resume Next
        .language = Range("language")
        On Error GoTo ErrExit
        
        .text = rg.Value2
        .header = nlsText.HeaderRowRange.Value2
        .offSet = FindCol(rg.rows(0).Value2, 1, c_firstLanguageName, False, "NLS_Text", True)         'find column of first language usually English
        
        If Not CheckNlsTableSetup(nlsText:=nlsText, offSet:=.offSet) Then Exit Sub
        
        .languageCount = UBound(.text, 2) - .offSet + 1
        .NLSTableColumnCount = .offSet - .sysCols                                          'Number of non-system columns in nlstext
        ReDim .filterVals(1 To .NLSTableColumnCount)                                                'Redim some arrays according to this
        ReDim .filterFieldNames(1 To .NLSTableColumnCount)
        ReDim .editFieldnames(1 To .NLSTableColumnCount)
        
        If .offSet = 0 Then GoTo ErrExit
        
        .systemLCID = Application.Dec2Hex(Application.LanguageSettings.LanguageID(msoLanguageIDUI), 4)                                      'Get UI language code
        
        systemLanguages = GetArrayFromRangeName("LCID", , ThisWorkbook.Name)                                                                'Get valid actions for undo - use english (pos 1) as default
        row = FindRowSorted(systemLanguages, 1, .systemLCID)
        If row = 0 Then
            Call ShowMessage("System", "undefinedSystemLanguage", smSystemError)
            .systemLanguage = 0
        Else
            .systemLanguage = systemLanguages(row, 5)
        End If
        
        .filterFieldNames(1) = "CB_F_Level"                                                         'Names of the NLSTable filter fields
        .filterFieldNames(2) = "CB_F_Module"
        .filterFieldNames(3) = "TB_F_Identifier"
        .filterFieldNames(4) = "CB_F_Type"
        .filterFieldNames(5) = "TB_F_Add"
        .filterFieldNames(6) = "TB_F_MainText"
        
        .editFieldnames(1) = "CB_E_Level"                                                         'Names of the NLSTable edit fields
        .editFieldnames(2) = "CB_E_Module"
        .editFieldnames(3) = "TB_E_Identifier"
        .editFieldnames(4) = "CB_E_Type"
        .editFieldnames(5) = "TB_E_Add"
        .editFieldnames(6) = "TB_E_MainText"
        
        .isSet = True
    End With
    
    On Error GoTo Finally
    
    '*********** Get local text - if available
    
    Set nlsText = GetObjectOfType("LocalNLSText", objtListObject, , ActiveWorkbook)
    
    If nlsText Is Nothing Then GoTo Combined
        
    With nlsText
    
        .AutoFilter.ShowAllData
        
        On Error Resume Next
        .DataBodyRange.Calculate
        
        .DataBodyRange.Columns(1) = .DataBodyRange.Columns(2).Value2
        .DataBodyRange.Columns(3).Value = g_NLSData.appLevelIdentifier
    
        .Sort.SortFields.Clear
        .Range.Sort key1:=.ListColumns(1).DataBodyRange.Cells(1), _
            order1:=xlAscending, DataOption1:=xlSortNormal, _
            SortMethod:=xlPinYin, MatchCase:=False, header:=xlYes
        On Error GoTo Finally
            
    End With
    
    Set rg = nlsText.DataBodyRange
     
    If Not CheckDataSet(rg) Then GoTo ErrExit
 
    With g_NLSData
    
        .localText = rg.Value2
        .hasLocalText = True
        
    End With
    
Combined:
    
    With g_NLSData
    
        If .hasLocalText Then
            .combinedText = CombineTwoDArrays(.localText, .text)
        Else
            .combinedText = .text
        End If
    
    End With
    
Finally:
'
'    If modeChanged Then Call RestoreEventsAndScreenUpdate
'
'    If cbSaved Then Call RestoreClipBoard

    g_NLSData.isSet = True
    
    Exit Sub
    
ErrExit:

'    If modeChanged Then Call RestoreEventsAndScreenUpdate
    
    MsgBox "Error in sub <setNLSData>. " & Err.Description, vbCritical, "Error"
End Sub
Function CheckNlsTableSetup(nlsText As ListObject, offSet As Long) As Boolean
'Checks that App and Pls level tables match in structure/header and match the language options range
    Dim localText As ListObject
    Dim languageOptions As Range
    Dim i As Long
    
    Set localText = GetObjectOfType("LocalNLSText", objtListObject, , ActiveWorkbook)
    
    If Not localText Is Nothing Then
        
        If localText.ListColumns.count <> nlsText.ListColumns.count Then
            MsgBox "Fatal Error. Platform and Application level NLS table structures do not match. Pls. do not proceed and contact the developer.", vbCritical + vbOKOnly, "Fatal Error"
            End
        End If
        
        For i = 1 To nlsText.ListColumns.count
            If localText.HeaderRowRange.Cells(1, i).Value2 <> nlsText.HeaderRowRange.Cells(1, i).Value2 Then
                MsgBox "Fatal Error. Platform and Application level NLS table headers do not match. Pls. do not proceed and contact the developer.", vbCritical + vbOKOnly, "Fatal Error"
                End
            End If
        Next i
        
    End If
    
    Set languageOptions = GetObjectOfType("LanguageOptions", objtRange, , ThisWorkbook)
    
    If languageOptions Is Nothing Then
        MsgBox "Fatal Error. Required range <LanguageOptions> is missing in the add-in. Pls. do not proceed and contact the developer.", vbCritical + vbOKOnly, "Fatal Error"
        End
    End If
    

    For i = offSet To nlsText.ListColumns.count
         If localText.HeaderRowRange.Cells(1, i).Value2 <> languageOptions.Cells(i - offSet + 1, 1).Value2 Then
            MsgBox "Fatal Error. Languages as setup in range <LanguageOptions> do not match NLS table languages. Pls. do not proceed and contact the developer.", vbCritical + vbOKOnly, "Fatal Error"
            End
         End If
    Next i

    CheckNlsTableSetup = True


End Function

Function GetNlsText(ByVal module As String, Optional ByVal identifier As String, Optional ByVal p1 As String, Optional ByVal p2 As String, Optional ByVal p3 As String, Optional ByVal p4 As String, Optional useSystemLanguage As Boolean, _
    Optional mandatory As Boolean, Optional quiet As Boolean = False) As String
'   Returns NLS text with max 4 optional string parms, defined as __&x__ in the string definition table NLS_Text; | is used to represent a new line character
'   If only one parameter is used (identifier = blank) then it can be either a language independent text to be used as is or a parm sep (default °°, defined in setNLSData) separated parm string
'       Optional useSystemLanguage  ......... If true, the excel system language will determine which language text will be returned
'       Optional mandatory .................. If true, a null string will be returned as indication of an error (not found on a mandatory, system critical entry) and sisplay of an error message, unless
'       Optional quiet .......................If false, an error meesage will be shown, if a mandatory entry cant be found

    Dim rw As Long
    Dim strParts As Variant
    Dim language As Long
    Dim nlsText As String
    Dim text As Variant
    
    With g_NLSData
        If Not .isSet Then
            Call SetNLSData
            If Not .isSet Then
                GetNlsText = "NLS Error"
                Exit Function
            End If
        End If
        
        If Len(identifier) = 0 Then                                    'Only module is provided - it could be just a non-MLS enabled text - or a parameter string
            
            strParts = Split(module, g_NLSData.parmSep)
            GetNlsText = strParts(0)
            If UBound(strParts) = 0 Then Exit Function              'does not contain the parameter separator - so its just a fixed text to be returned
            
            On Error Resume Next                                    'it contains a parameter seperator - its a parmstring , so get the parms
            
            module = strParts(0)
            identifier = strParts(1)
            p1 = strParts(2)
            p2 = strParts(3)
            p3 = strParts(4)
            p4 = strParts(5)
            
            On Error GoTo 0
           
        End If
        
        If .hasLocalText Then
            text = .localText
            rw = FindRowSorted(text, 1, LCase(module & identifier))                   'First check in local text
        End If
        
        If rw = 0 Then                                                                'then in text from add-in
            text = .text
            rw = FindRowSorted(text, 1, LCase(module & identifier))
        End If
        
        
        If rw = 0 Then
            If mandatory Then
                If Not quiet Then
                    MsgBox "Mandatory entry in NLSText for <" & module & "/" & identifier & "> is missing. Pls. contact the developer"
                End If
            Else
                GetNlsText = "Unknown message"                          'default error message
                On Error Resume Next
                GetNlsText = .notFoundMessages(.language)       'customized, language dependent error message
                GetNlsText = GetNlsText & " <" & module & "/" & identifier & ">"
            End If
        Else
            If useSystemLanguage Then language = .systemLanguage Else language = .language
            nlsText = text(rw, language + .offSet)
            If .useFirstLanguageIfBlank And Len(nlsText) = 0 Then
                nlsText = text(rw, .offSet)
            End If
            GetNlsText = Replace(Replace(Replace(Replace(Replace(nlsText, "__&1__", p1), "__&2__", p2), "__&3__", p3), "__&4__", p4), "|", vbNewLine)
        End If
    End With
End Function


Sub InvalidateNlsText(Optional dummy As Boolean)
'Forces a reload of the NLS data at next use
    g_NLSData.isSet = False
End Sub
Sub PrintNlsTextCol(col As Long)
'Helper function
    Dim i As Long
    For i = 1 To UBound(g_NLSData.text, 1)
        Debug.Print g_NLSData.text(i, col)
    Next i
End Sub
Function GetNlsAddition(module As String, identifier As String) As String
    'Returns "Additionl" column value form NLS text matrix - which is placed just before the language columns
    Dim rw As Long
    
    Call SetNLSData
    
        With g_NLSData
        
        If Not .isSet Then Exit Function
        
        rw = FindRowSorted(.text, 1, LCase(module & identifier))
        
        If rw > 0 Then GetNlsAddition = .text(rw, g_NLSData.addCol)
        
    End With
    
End Function
Function getNLSKeyStatus(module As String, identifier As String) As nlsKeyStatus
'Checks status of NLS key values
    Dim pltStatus As Boolean, addStatus As Boolean
    
    If Len(trim(module)) = 0 Or Len(trim(identifier)) = 0 Then Exit Function
    
    getNLSKeyStatus = nksNew
    
    With g_NLSData
        
        If Not .isSet Then Call SetNLSData
        
        pltStatus = FindRowSorted(.text, 1, module & identifier) > 0
        
        If .hasLocalText Then addStatus = FindRowSorted(.localText, 1, module & identifier) > 0
        
        If pltStatus Then
            If addStatus Then getNLSKeyStatus = nksBoth Else getNLSKeyStatus = nksPlt
        Else
            If addStatus Then getNLSKeyStatus = nksAdd
        End If
        
    End With

End Function
Sub ShowNlsTable(Optional initializeOnly As Boolean, Optional ByRef selectedModule As Variant, Optional ByRef selectedIdentifier As Variant)
    Dim colWidths As String, module As String, identifier As String
    Dim finds() As Long
    Dim res As dissectResult
    Dim hasEntryVals As Boolean
    Dim modeChanged As Boolean
    Dim listSeparator As String
    Dim calledFromVBE As Boolean
    Dim selected As Variant
    
    calledFromVBE = VbeHasFocus
    Set selected = Selection
    
    colWidths = c_defaultColWidths

    With g_NLSData
        If Not .isSet Then
            Call SetNLSData
            If Not .isSet Then Exit Sub
        End If
    End With
    
    modeChanged = setFastMode

    hasEntryVals = getModuleAndIdentifier(module, identifier)

    With NLSTable
        .mainLanguage = c_mainLanguage
        .colCnt = g_NLSData.NLSTableColumnCount
        .colWidths = colWidths
        .cols = Array(3, 4, 5, 6, 7, g_NLSData.offSet + .mainLanguage)
        
        If hasEntryVals Then
            .entryModule = module
            .entryIdentifier = identifier
        End If
        
        If Not initializeOnly Then
        
            .show
        
            If Not .wasCanceled Then
                If Not IsMissing(selectedModule) And Not IsMissing(selectedIdentifier) Then                     'If requested by optional parms then return module & identifier, else put into clipboard
                    selectedModule = .module
                    selectedIdentifier = .identifier
                Else
                    Call ClearClipBoard
                    If calledFromVBE Then
                        Call PutTextToClipboard(Chr(34) & .module & Chr(34) & "," & Chr(34) & .identifier & Chr(34))
                    Else
                        If TypeName(selected) = "Range" Then
                            Call PutTextToClipboard(Chr(34) & .module & Chr(34) & Application.International(xlListSeparator) & Chr(34) & .identifier & Chr(34))
                        Else
                            Call PutTextToClipboard(.identifier)
                        End If
                    End If
                End If
            End If
            
'            Unload NLSTable
'            Unload NLSTableEdit
        
        End If
        
        If modeChanged Then Call resetFastMode
        
        selected.Select
        
    End With
End Sub

Function getModuleAndIdentifier(ByRef module As String, ByRef identifier As String, Optional ByVal codeLine As String, Optional pos As Long, Optional literalsOnly As Boolean = True) As Boolean
'Returns module and identifier info when
'   a line of code is provided or is
'   positioned on a VBE activePane line with appropriate function calls or
'   a cell with a relevant UDF is slected or a shape(autoshape) is selected.
'   Only calls with two parms (module, identifier) as literals will be considered here

    Dim i As Long, hit As Long, delta As Long, length As Long
    Dim parmInfo() As parameterInfo
    Dim wasProtected As Boolean
    Dim shape As Variant
    
    If Not g_NLSData.isSet Then
        Call SetNLSData
    End If
    
    If Len(codeLine) = 0 Then
    
        If VbeHasFocus Then
            If Not getActiveVBECodeLine(codeLine, pos, length) Then Exit Function
        Else
            pos = 1
            If TypeName(Selection) = "Range" Then
                If Selection.Parent.Name = c_infoSheetName Then
                    With ActiveWorkbook.Worksheets(c_infoSheetName)
                    
                        On Error Resume Next                    'Find the respective columns of module and identifier if on Info sheet
                        module = .Cells(Selection.row, .rows(3).Find(what:="module", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).column).Value2
                        identifier = .Cells(Selection.row, .rows(3).Find(what:="identifier", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).column).Value2
                        On Error GoTo 0
                        
                        If Len(module) > 0 And Len(identifier) > 0 Then getModuleAndIdentifier = True
                        Exit Function
                    End With
                Else
                    codeLine = GetFormula(Selection)
                End If
            Else
                On Error Resume Next
                Err.Clear
                Set shape = Selection.Parent.Shapes(Selection.Name)
                If Err.Number > 0 Then Exit Function
                If shape.Type <> msoAutoShape Then Exit Function
                module = c_moduleForShapes
                identifier = shape.Name
                getModuleAndIdentifier = True
                Exit Function
            End If
        End If
        
    End If
    
    If Len(codeLine) = 0 Then Exit Function
    
    parmInfo = getCallParametersFromString(codeLine, g_NLSData.udfNames)
         
    For i = 1 To UBound(parmInfo)               'Find the closest occurence left of the position
        If pos >= parmInfo(i).position Then
            hit = i
        End If
    Next i
    
    hit = Application.Min(hit, UBound(parmInfo))
    
    If hit > 0 Then
        With parmInfo(hit)
            If UBound(.parms) < 2 Then Exit Function        'Less than two parameters - do not consider a regular call, though one parm is allowed (passthrough)
    
            If LCase(.funCall) = "setmenuelement" Then                  'special treatment for setmenuelement
                module = c_moduleForMenus
                identifier = TrimAll(.parms(2), Chr(34))                'Remove quotes
                GoTo Finally
            End If

            If literalsOnly Then
                If Not (IsLiteral(.parms(1)) And IsLiteral(.parms(2))) Then Exit Function
                module = TrimAll(.parms(1), Chr(34))                                                'Remove quotes
                identifier = TrimAll(.parms(2), Chr(34))
            Else
                module = .parms(1)
                identifier = .parms(2)
            End If
            
        End With
    End If
    
Finally:
    getModuleAndIdentifier = Len(trim(module)) > 0 And Len(trim(identifier)) > 0
    
    Debug.Print getModuleAndIdentifier, module, identifier, codeLine

End Function
Sub ShowNlsCalls(Optional show As Boolean = True)
    Dim cnt As Long
    Dim sh As Worksheet
    Dim rg As Range
    Dim modeChanged As Boolean, wasProtected As Boolean
            
    Set sh = ActiveWorkbook.Worksheets(c_infoSheetName)
    sh.Cells.Clear
    sh.Cells.WrapText = False
        
    modeChanged = setFastMode
        
    ReDim g_NLSData.xRef(1 To 10000, 1 To 9)
    
    Call GetNlsOccurrenciesInWorkBookCells(ActiveWorkbook, g_NLSData.xRef, cnt, True)               'Get all calls in Cells in application workbook
    
    Call GetNlsOccurrenciesInCode(ActiveWorkbook, g_NLSData.xRef, cnt)                              'Get all calls in VBA code in application workbook
    
    Call GetNlsOccurrenciesInCode(ThisWorkbook, g_NLSData.xRef, cnt)                                'Get all calls in VBA code in add-in workbook

    With sh
    
        With .Cells(1, 1)
            .Cells(1).Value = "NLS calls, cells, shapes etc using NLS table entries"
            With .font
                .Size = 16
                .Bold = True
                .color = RGB(255, 0, 0)
            End With
        End With
        
        .Cells(1, 6).Value = "By filtering column NLSText for '" & GetNlsText("NLS", "Unknown") & "', you can identify missing NLS table entries"
        .Cells(1, 6).font.Bold = True
    
        Set rg = ArrayToRange(anchor:=.Cells(3, 1), arr:=Array("WorkBook", "Area", "Module/Sheet", "Line/Address/Name", "Call", "Module", "Identifier", "Key", "NLS Text"))
        .rows(3).font.Bold = True
        .rows(3).VerticalAlignment = xlTop
        
        Set rg = ArrayToRange(anchor:=sh.Cells(4, 1), arr:=g_NLSData.xRef, rws:=cnt, noWrapText:=True)             'Put on info sheet, sort by "Key" and reload into xRef
        rg.Sort key1:=rg.Cells(1, 8)
        g_NLSData.xRef = GetArrayFromRange(rg)
        
        .Cells.HorizontalAlignment = xlLeft
        
        rg.offSet(-1).Resize(rg.rows.count + 1).AutoFilter
        
    End With

    If modeChanged Then Call resetFastMode
    
    If show Then
        If Not sh.Visible Then
            wasProtected = DoUnProtectWorkBook
            sh.Visible = True
            If wasProtected Then Call DoProtectWorkBook
        End If
        sh.Activate
        sh.Cells(1, 1).Select
        Call SwitchToExcel
    End If
    
End Sub
Sub GetNlsOccurrenciesInCode(wb As Workbook, ByRef refs As Variant, ByRef cnt As Long)
'Find and decodes all NLS relevant calls in a workbook that have literals for module and identifier and are not in a comment
'Places all matches ito the Array "refs" as defined after the DeodeNLS function call
'Special treatment is provided for the last element in  udfNames (SetMenuElement) ****** Consider splitting that into a seperate function
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.codeModule
    Dim findWhat As String
    Dim startLine As Long, endLine As Long, startColumn As Long, endColumn As Long
    Dim i As Long, posOfCommentChar As Long
    Dim Found As Boolean
    Dim file As String, module As String, codeLine As String, identifier As String
        
    Err.Clear
    On Error Resume Next
    Set vbProj = wb.VBProject
    If Err.Number > 0 Then Exit Sub
    
    On Error GoTo 0
    
    For Each vbComp In vbProj.VBComponents
    
       Set CodeMod = vbComp.codeModule

        For i = 0 To UBound(g_NLSData.udfNames)
                
            findWhat = g_NLSData.udfNames(i) & "("
    
            With CodeMod
                startLine = 1
                endLine = .CountOfLines
                startColumn = 1
                endColumn = 255
                Found = .Find(target:=findWhat, startLine:=startLine, startColumn:=startColumn, _
                    endLine:=endLine, endColumn:=endColumn, _
                    wholeword:=False, MatchCase:=False, patternsearch:=False)
                Do Until Found = False
                    
                    codeLine = .lines(startLine, 1)
                    
                    posOfCommentChar = InStr(codeLine, "'")
                    
                    If posOfCommentChar = 0 Or startColumn < posOfCommentChar Then                                                          'Check to exclude commented calls
                        Select Case findWhat
                            Case "SetMenuElement("
                                If getModuleAndIdentifier(module, identifier, Mid(codeLine, startColumn), 1, False) Then
                                    If IsLiteral(identifier) Then
                                        cnt = cnt + 1
                                        refs(cnt, 1) = wb.Name
                                        refs(cnt, 2) = "VBA Code"
                                        refs(cnt, 3) = vbComp.Name
                                        refs(cnt, 4) = startLine
                                        refs(cnt, 5) = g_NLSData.udfNames(i)
                                        refs(cnt, 6) = c_moduleForMenus
                                        refs(cnt, 7) = TrimAll(identifier, Chr(34))
                                        refs(cnt, 8) = LCase(refs(cnt, 6) & refs(cnt, 7))
                                        refs(cnt, 9) = GetNlsText(refs(cnt, 6), refs(cnt, 7), "__&1__", "__&2__", "__&3__", "__&4__")
                                    End If
                                End If
                            Case Else
                                If getModuleAndIdentifier(module, identifier, Mid(codeLine, startColumn), 1) Then
                                    cnt = cnt + 1
                                    refs(cnt, 1) = wb.Name
                                    refs(cnt, 2) = "VBA Code"
                                    refs(cnt, 3) = vbComp.Name
                                    refs(cnt, 4) = startLine
                                    refs(cnt, 5) = g_NLSData.udfNames(i)
                                    refs(cnt, 6) = module
                                    refs(cnt, 7) = identifier
                                    refs(cnt, 8) = LCase(refs(cnt, 6) & refs(cnt, 7))
                                    refs(cnt, 9) = GetNlsText(refs(cnt, 6), refs(cnt, 7), "__&1__", "__&2__", "__&3__", "__&4__")
                                End If
                        End Select
                    End If
                    
                    
                    endLine = .CountOfLines
                    startColumn = endColumn + 1
                    endColumn = 255
                    Found = .Find(target:=findWhat, startLine:=startLine, startColumn:=startColumn, _
                        endLine:=endLine, endColumn:=endColumn, _
                        wholeword:=False, MatchCase:=False, patternsearch:=False)
                Loop
            End With
        
        Next i
        
    Next vbComp
    
End Sub
Sub GetNlsOccurrenciesInWorkBookCells(wb As Workbook, ByRef refs As Variant, ByRef cnt As Long, Optional addButtons As Boolean)
'Gets all occurrencies of getNLSText in wb
'Adds all butttons if requested
    Dim matchCell As Range, firstCell As Range
    Dim wasProtected As Boolean
    Dim module As String, identifier As String, currentFiltRange As String
    Dim sh As Worksheet
    Dim parms() As parameterInfo
    Dim i As Long
    Dim shape As shape, shape2 As shape
    Dim wasFiltered As Boolean
    Dim filterArray As Variant
    
    Call SetNLSData
    
    For Each sh In wb.Worksheets                      'Only getNLSText is supported - all other uses are strictly code related
        
        wasProtected = False
        
        With GetUsedRange(sh)
        
            If sh.FilterMode Then                                                        'Sheet is filtered - need to unfilter to allow full search
                Call SaveAutoFilter(sh, currentFiltRange, filterArray)
                sh.AutoFilterMode = False
                wasFiltered = True
            End If
            
            Set matchCell = .Find(what:="getNLSText", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
            If Not matchCell Is Nothing Then
                Set firstCell = matchCell
                
                If Not wasProtected Then wasProtected = DoUnProtect(sh)
                
                Do
                    parms = getCallParametersFromString(matchCell.Formula, g_NLSData.udfNames(0))
                    For i = 1 To UBound(parms)
                        cnt = cnt + 1
                        refs(cnt, 1) = wb.Name
                        refs(cnt, 2) = "Cell"
                        refs(cnt, 3) = matchCell.Parent.Name
                        refs(cnt, 4) = matchCell.Address
                        refs(cnt, 5) = parms(i).funCall
                        refs(cnt, 6) = TrimAll(parms(i).parms(1), Chr(34))
                        refs(cnt, 7) = TrimAll(parms(i).parms(2), Chr(34))
                        refs(cnt, 8) = LCase(refs(cnt, 6) & refs(cnt, 7))
                        refs(cnt, 9) = GetNlsText(refs(cnt, 6), refs(cnt, 7), "__&1__", "__&2__", "__&3__", "__&4__")
                    Next i
                        
                    Set matchCell = .FindNext(matchCell)
                Loop While Not matchCell Is Nothing And Not matchCell.Address = firstCell.Address
            End If
        
            If wasFiltered Then
                Call RestoreAutofilter(sh, currentFiltRange, filterArray)
                wasFiltered = False
            End If
            
        End With
        
        If addButtons Then

            For Each shape In sh.Shapes                                     'Add module Button with identifier shape Name + indicator if button is NLS enabled - do only for msoAutoShapes
                Select Case shape.Type
                    Case msoAutoShape
                        cnt = cnt + 1
                        refs(cnt, 1) = wb.Name
                        refs(cnt, 2) = "Shape"
                        refs(cnt, 3) = sh.Name
                        refs(cnt, 4) = shape.Name
                        If getNLSKeyStatus(c_moduleForShapes, shape.Name) > 1 Then
                            refs(cnt, 5) = "-"
                            refs(cnt, 6) = c_moduleForShapes
                            refs(cnt, 7) = shape.Name
                            refs(cnt, 8) = LCase(refs(cnt, 6) & refs(cnt, 7))
                            refs(cnt, 9) = GetNlsText(refs(cnt, 6), refs(cnt, 7), "__&1__", "__&2__", "__&3__", "__&4__")
                        End If
                    Case msoGroup                                           'Handle grouped shapes
                        For Each shape2 In shape.GroupItems
                            cnt = cnt + 1
                            refs(cnt, 1) = wb.Name
                            refs(cnt, 2) = "Shape"
                            refs(cnt, 3) = sh.Name
                            refs(cnt, 4) = shape2.Name
                            If getNLSKeyStatus(c_moduleForShapes, shape2.Name) > 1 Then
                                refs(cnt, 5) = "In NLSTable"
                                refs(cnt, 6) = c_moduleForShapes
                                refs(cnt, 7) = shape2.Name
                                refs(cnt, 8) = LCase(refs(cnt, 6) & refs(cnt, 7))
                                refs(cnt, 9) = GetNlsText(refs(cnt, 6), refs(cnt, 7), "__&1__", "__&2__", "__&3__", "__&4__")
                            End If
                        Next shape2
                    Case Else
                    
                    
                End Select
            Next shape
        
        End If

        If wasProtected Then Call DoProtect(sh)

    Next sh

End Sub

Sub ShowWidows(Optional dummy As Boolean)
'Shows the list of (probably) unused NLS table entries

    Call ShowNlsTable(initializeOnly:=True)
    
    Call NLSTable.B_ShowWidows_Click
       
    Unload NLSTable

End Sub
Sub LocateNLSCall()
'Move active area to NLS call
    
    Dim line As Variant
    Dim module As String, identifier As String
    Dim rowNum As Long, startLine As Long, startCol As Long, endLine As Long, endCol As Long
    Dim book As Workbook
    Dim sh As Worksheet
    Dim cell As Range, defRange As Range
    Dim shape As shape
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codePane As VBIDE.codePane
    Dim wasProtected As Boolean
    
    On Error GoTo ErrorExit
    
    line = Intersect(ActiveCell.CurrentRegion, ActiveSheet.rows(ActiveCell.row)).Value2
    
    Set book = Workbooks(line(1, 1))
    
    Select Case LCase(line(1, 2))
        Case "cell"
            Set sh = book.Worksheets(line(1, 3))
            
            If Not sh.Visible = xlSheetVisible Then
                If Not (IsDeveloper Or IsPlatformDeveloper) Then
                    Call ShowMessage("system", "hiddenSheet", smWarning, sh.Name)
                    Exit Sub
                End If
                wasProtected = DoUnProtectWorkBook
                sh.Visible = xlSheetVisible
                If wasProtected Then Call DoProtectWorkBook
            End If
            
            Set cell = sh.Range(line(1, 4))
            book.Activate
            sh.Activate
            cell.Select
    
        Case "shape"
            Set sh = book.Worksheets(line(1, 3))
            
            If Not sh.Visible = xlSheetVisible Then
                If Not (IsDeveloper Or IsPlatformDeveloper) Then
                    Call ShowMessage("system", "hiddenSheet", smWarning, sh.Name)
                    Exit Sub
                End If
                wasProtected = DoUnProtectWorkBook
                sh.Visible = xlSheetVisible
                If wasProtected Then Call DoProtectWorkBook
            End If
            
            Set shape = sh.Shapes(line(1, 4))
            
            If Not shape.Visible = msoTrue Then
                If ShowConfirm("NLS", "unhideShape", cfWarning) Then shape.Visible = msoTrue
            End If
            
            book.Activate
            sh.Activate
            shape.Select
        
        Case "vba code"
            Set vbProj = book.VBProject
            Set vbComp = vbProj.VBComponents(line(1, 3))
            Set codePane = vbComp.codeModule.codePane
            codePane.show
            startLine = line(1, 4)
            startCol = 1
            endLine = startLine
            endCol = 512
            codePane.SetSelection startLine, startCol, endLine, endCol
        
        Case Else
    End Select
    
    
    Exit Sub
    
ErrorExit:
    Call ShowMessage("NLS", "callNotFound", smError, Err.Description)
End Sub
Sub SetLanguage()
' Sets the language as requested, including for the shapes as defined below
    Dim sh As Worksheet
    Dim cl As Range
    Dim matchCell As Range
    Dim firstCell As Range
    Dim modeChanged As Boolean
    Dim shape As shape
    Dim nlsKeyStatus As nlsKeyStatus
    Dim wasFiltered As Boolean
    Dim filterArray As Variant
    Dim currentFiltRange As String
    
    Call SetNLSData(force:=True)                                    'Safety
    
    Range("ActiveLanguage").Value = Range("SelectedLanguage").Value2
    
    modeChanged = setFastMode
    
    g_NLSData.language = Range("language")
    
    Set cl = ActiveCell
    
    On Error Resume Next
    
    For Each sh In ActiveWorkbook.Worksheets                      'Replace Shape text in Main Sheets + Basistabellen and setdirty for getNLSText
    
        With GetUsedRange(sh)
        
            If sh.FilterMode Then                                                        'Sheet is filtered - need to unfilter to allow full search
                Call SaveAutoFilter(sh, currentFiltRange, filterArray)
                sh.AutoFilterMode = False
                wasFiltered = True
            End If
            
            Set matchCell = .Find(what:="GetNLSText", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)             'find all uses of getNLSText and mark as dirty
            If Not matchCell Is Nothing Then
                Set firstCell = matchCell
                Do
                    matchCell.Dirty
                    Set matchCell = .FindNext(matchCell)
                Loop While Not matchCell Is Nothing And Not matchCell.Address = firstCell.Address
            End If
        
            If wasFiltered Then
                Call RestoreAutofilter(sh, currentFiltRange, filterArray)
                wasFiltered = False
            End If
            
        End With
        
        On Error Resume Next
        
        For Each shape In sh.Shapes                                     'Handle shapes - if nls record with module = button and identifier = shape.name then change text
            nlsKeyStatus = getNLSKeyStatus(c_moduleForShapes, shape.Name)
            If nlsKeyStatus > 1 Then
                shape.TextFrame.Characters.text = GetNlsText(c_moduleForShapes, shape.Name)
            End If
        Next shape
    
        On Error GoTo 0
    
    Next sh
    
    Call SetSeparators
    
    Application.Goto cl
    
    If modeChanged Then Call resetFastMode
    
End Sub
Sub SetToSystemLanguage()
    Dim systemLCID As String, systemLanguage As String
    Dim row As Long, sysLangNumber As Long
    Dim systemLanguages As Variant


    systemLCID = Application.Dec2Hex(Application.LanguageSettings.LanguageID(msoLanguageIDUI), 4)                                      'Get UI language code
        
    systemLanguages = GetArrayFromRangeName("LCID", , ThisWorkbook.Name)
    row = FindRowSorted(systemLanguages, 1, systemLCID)
    sysLangNumber = systemLanguages(row, 5)
    systemLanguage = Range("LanguageOptions").Cells(sysLangNumber + 1, 1).Value2
    
    With Range("SelectedLanguage")
        If .Value2 <> systemLanguage Then .Value = systemLanguage
    End With
    
    Exit Sub
    
ErrorExit:
        Call ShowMessage("System", "undefinedSystemLanguage", smSystemError)
End Sub
Function Text2(dateVal, format) As String
    Dim dateLocale As String
    
    On Error Resume Next
    dateLocale = Range("DateLocale").Value2
    Text2 = Application.text(dateVal, dateLocale & format)
End Function
Sub NLSTableEditEntry()
'Used for chain editing entries in NLS table to avoid stack regression
    Call NLSTable.EditEntry
End Sub
