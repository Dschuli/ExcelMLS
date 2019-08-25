Attribute VB_Name = "SupportFunctions"
Option Explicit

'*************Declarations
Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
Public Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'************ Enumerations
Public Enum objectType                      'Type of object for check/find operations
    objtRange
    objtPivotTable
    objtListObject
    objtShape
    objtSheet
    objtBook
End Enum

'************ Custom datatypes

Type dissectResult
    before As String
    hit As String
    after As String
End Type

Type parameterInfo                                      'Structure to hold info about parameters of a sub or function call string
    position As Long
    funCall As String
    parms() As String
End Type

'********** Code

Function IsObjectOfType(objName As String, objType As objectType, Optional sh As Worksheet, Optional wb As Workbook) As Boolean
'checks the existance of an object of a given type in a workbook/sheet
        
        On Error Resume Next
        
        Err.Clear
        
        If wb Is Nothing Then Set wb = ActiveWorkbook
        
        Select Case objType
        
        Case objtRange
            If sh Is Nothing Then
                If wb Is ActiveWorkbook Then
                    IsObjectOfType = Not Range(objName) Is Nothing
                Else
                    For Each sh In wb.Worksheets
                        Err.Clear
                        IsObjectOfType = Not sh.Range(objName) Is Nothing
                        If Err.Number = 0 Then Exit Function
                    Next sh
                End If
            Else
                IsObjectOfType = Not sh.Range(objName) Is Nothing
            End If
                
        Case objtPivotTable
            If sh Is Nothing Then
                For Each sh In wb.Worksheets
                    Err.Clear
                    IsObjectOfType = Not sh.PivotTables(objName) Is Nothing
                    If Err.Number = 0 Then Exit Function
                Next sh
            Else
                IsObjectOfType = Not sh.PivotTables(objName) Is Nothing
            End If
            
        Case objtListObject
            If sh Is Nothing Then
                For Each sh In wb.Worksheets
                    Err.Clear
                    IsObjectOfType = Not sh.ListObjects(objName) Is Nothing
                    If Err.Number = 0 Then Exit Function
                Next sh
            Else
                IsObjectOfType = Not sh.ListObjects(objName) Is Nothing
            End If
        
        Case objtShape
            If sh Is Nothing Then
                For Each sh In wb.Worksheets
                    Err.Clear
                    IsObjectOfType = Not sh.Shapes(objName) Is Nothing
                    If Err.Number = 0 Then Exit Function
                Next sh
            Else
                IsObjectOfType = Not sh.Shapes(objName) Is Nothing
            End If
            
        Case objtSheet
            IsObjectOfType = Not wb.Worksheets(objName) Is Nothing
            
        Case objtBook
            IsObjectOfType = Not Application.Workbooks(objName) Is Nothing
        
        Case Else
            MsgBox "Object type not implemented in <isObjectOfType> function"
            
        End Select
    
End Function
Function GetObjectOfType(objName As String, objType As objectType, Optional sh As Worksheet, Optional wb As Workbook) As Variant
'Checks the existance of an object of a given type in a workbook/sheet
'For ranges
        
        On Error Resume Next
        
        Err.Clear
        
        If wb Is Nothing Then Set wb = ActiveWorkbook
        
        Set GetObjectOfType = Nothing
        
        Select Case objType
        
        Case objtRange
            If sh Is Nothing Then
                If wb Is ActiveWorkbook Then
                    Set GetObjectOfType = Range(objName)
                Else
                    Set GetObjectOfType = Range("'" & wb.Name & "'!" & objName)
                End If
            Else
                Set GetObjectOfType = sh.Range(objName)
            End If
                
        Case objtPivotTable
            If sh Is Nothing Then
                For Each sh In wb.Worksheets
                    Err.Clear
                    Set GetObjectOfType = sh.PivotTables(objName)
                    If Err.Number = 0 Then Exit Function
                Next sh
            Else
                Set GetObjectOfType = sh.PivotTables(objName)
            End If
            
        Case objtListObject
            If sh Is Nothing Then
                For Each sh In wb.Worksheets
                    Err.Clear
                    Set GetObjectOfType = sh.ListObjects(objName)
                    If Err.Number = 0 Then Exit Function
                Next sh
            Else
                Set GetObjectOfType = sh.ListObjects(objName)
            End If
        
        Case objtShape
            If sh Is Nothing Then
                For Each sh In wb.Worksheets
                    Err.Clear
                    Set GetObjectOfType = sh.Shapes(objName)
                    If Err.Number = 0 Then Exit Function
                Next sh
            Else
                Set GetObjectOfType = sh.Shapes(objName)
            End If
            
        Case objtSheet
            Set GetObjectOfType = wb.Worksheets(objName)
            
        Case objtBook
            Set GetObjectOfType = Application.Workbooks(objName)
        
        Case Else
            MsgBox "Object type not implemented in <GetObjectOfType> function"
            
        End Select
    
End Function
Function FindRowSorted(lowerCaseArr As Variant, col As Long, val As Variant, Optional silent As Boolean = True, Optional where As String = "?") As Long
    'Finds the row which contains the value <val> in Column <col> in an arry or range <arr>
    '0 indicates not found
    'lowerCaseArr has to be all lowerCase to avoid sorting issue
        
    On Error GoTo Err_Exit
    Dim intTop As Long, intMiddle As Long, intBottom As Long, i As Long, first As Long, last As Long
    Dim inverseOrder As Boolean
    Dim middle As Double

    first = LBound(lowerCaseArr, 1)
    last = UBound(lowerCaseArr, 1)

    val = LCase(CStr(val))

    ' deduct direction of sorting
    inverseOrder = (lowerCaseArr(first, col) > lowerCaseArr(last, col))

    ' assume searches failed
    FindRowSorted = 0
    
    Do
        middle = (first + last) \ 2
        If LCase(CStr(lowerCaseArr(middle, col))) = val Then
            FindRowSorted = middle
            Exit Do
        ElseIf ((LCase(CStr(lowerCaseArr(middle, col))) < val) Xor inverseOrder) Then
            first = middle + 1
        Else
            last = middle - 1
        End If
    Loop Until first > last
    
    If FindRowSorted > 0 Then Exit Function
    
    If silent Then Exit Function
    Call ShowMessage("System", "notFound", smWarning, CStr(val), where, "FindRowSorted")
    Exit Function
Err_Exit:
    Call ShowMessage("System", "opsFail", smSystemError, "FindRowSorted", Err.Description)
End Function

Sub CopyControlProperties(master As MSForms.control, ByRef slave As MSForms.control)
'Dim copy property values of a master control to a slave control
'Not a full implementation
    On Error Resume Next
    slave.Left = master.Left
    slave.top = master.top
    slave.height = master.height
    slave.width = master.width
    slave.Enabled = master.Enabled
    slave.caption = master.caption
    Set slave.font = master.font
    slave.ForeColor = master.ForeColor
    slave.BackColor = master.BackColor
    slave.Alignment = master.Alignment
    slave.Locked = master.Locked
    slave.MatchEntry = master.MatchEntry
    slave.MatchRequired = master.MatchRequired
    slave.TextAlign = master.TextAlign
    slave.BoundColumn = master.BoundColumn
    slave.TextColumn = master.TextColumn
    slave.Style = master.Style
    slave.TripleState = master.TripleState
End Sub
Function CheckDataSet(rg As Range, Optional confirm As Boolean = False, Optional silent As Boolean = False, Optional severity As msgLevel = smError) As Boolean
'To prevent malfunctions if a dataset contains errors this function checks if a cell in a range contains an error
'Stops with error messsage at first error
'Returns false on error else true

    Dim vals() As Variant, val As Variant
    Dim i As Long, j As Long
    Dim cellAddress As String
    Dim area As Range
    
    On Error Resume Next
    rg.Calculate
    On Error GoTo 0
    
    If rg Is Nothing Then Exit Function
    
    For Each area In rg.Areas
    
        If area.Cells.CountLarge = 1 Then
            ReDim vals(1 To 1, 1 To 1)
            vals(1, 1) = area.Value2
        Else
            vals = area.Value2
        End If
        
        For i = 1 To UBound(vals, 1)
            For j = 1 To UBound(vals, 2)
                If IsError(vals(i, j)) Then
                    cellAddress = CLetter(j + rg.column - 1) & CStr(i + rg.row - 1)
                    If confirm Then
                        CheckDataSet = ShowConfirm("System", "dataSetHasErrorsConfirm", cfWarning, rg.Parent.Name, cellAddress)
                    Else
                        If Not silent Then Call ShowMessage("System", "DataSetHasErrors", smError, rg.Parent.Name, cellAddress)
                    End If
                    Exit Function
                End If
            Next j
        Next i
    
    Next area
    
    CheckDataSet = True
    
End Function
Function FindRow(arr As Variant, col As Long, val As Variant, Optional silent As Boolean = True, Optional where As String = "?", Optional ignoreCase As Boolean = False) As Long
    'Finds the row which contains the value <val> in Column <col> in an array
    '0 indicates not found
    
    On Error GoTo Err_Exit
    Dim lb As Long, ub As Long, i As Long
    
    lb = LBound(arr, 1)
    ub = UBound(arr, 1)
    
    Err.Clear
    On Error GoTo Err_Exit
    
    If ignoreCase Then
        For i = lb To ub
            If LCase(arr(i, col)) = LCase(val) Then
                FindRow = i
                Exit Function
            End If
        Next i
    Else
        For i = lb To ub
            If arr(i, col) = val Then
                FindRow = i
                Exit Function
            End If
        Next i
    End If
    
    If silent Then Exit Function
    Call ShowMessage("System", "notFound", smWarning, CStr(val), where, "FindRow")
    Exit Function
Err_Exit:
    Call ShowMessage("System", "opsFail", smSystemError, "FindRow", Err.Description)
End Function

Function FindCol(arr As Variant, rw As Long, val As Variant, Optional silent As Boolean = True, Optional where As String = "?", Optional ignoreCase As Boolean = False) As Long
    'Finds the column which contains the value <val> in row <rw> in an array
    '0 indicates not found
    
    On Error GoTo Err_Exit
    Dim lb As Long, ub As Long, i As Long, j As Long
    
    lb = LBound(arr, 2)
    ub = UBound(arr, 2)

    If ignoreCase Then
        val = LCase(val)
        For j = lb To ub
            If LCase(arr(rw, j)) = val Then
                FindCol = j
                Exit Function
            End If
        Next j
    Else
        For j = lb To ub
            If arr(rw, j) = val Then
                FindCol = j
                Exit Function
            End If
        Next j
    End If
    
    If silent Then Exit Function
    
    If g_NLSData.isSet Then
        Call ShowMessage("System", "notFound", smWarning, CStr(val), where, "FindCol")
    Else
        MsgBox "Value <" & val & "> not found in <" & where & "> .Function: <FindCol>", vbOKOnly + vbCritical, "Error"
    End If
    
    Exit Function
Err_Exit:

    If g_NLSData.isSet Then
        Call ShowMessage("System", "opsFail", smSystemError, "FindCol", Err.Description)
    Else
        MsgBox "Error during operation <FindCol>. Hint <" & Err.Description & ">.", vbOKOnly + vbCritical, "Error"
    End If
    
End Function
Function GetArrayFromRangeName(rangeName As String, Optional shName As String, Optional bkName As String, Optional quiet As Boolean = False) As Variant()
    'Returns an arry of values(s) - even from single cell
    Dim bk As Workbook
    Dim rg As Range
    Dim sh As Worksheet
    Dim arr(1 To 1, 1 To 1) As Variant
    
    On Error GoTo ErrorExit
    If Len(bkName) = 0 Then Set bk = ActiveWorkbook Else Set bk = Workbooks(bkName)
    If Len(shName) = 0 Then
        Set sh = GetObjectOfType(rangeName, objtRange, , bk).Parent
    Else
        Set sh = bk.Worksheets(shName)
    End If
    Set rg = sh.Range(rangeName)
    If rg.Cells.CountLarge > 1 Then
        GetArrayFromRangeName = rg.Value2
    Else
        arr(1, 1) = rg.Value
        GetArrayFromRangeName = arr
    End If
    Exit Function
ErrorExit:
    If Not quiet Then Call ShowMessage("System", "MissingRange", smSystemError, rangeName)
End Function
Function GetArrayFromRange(rg As Range) As Variant
    'Returns a 2dimensional array even if single cell
    Dim arr(1 To 1, 1 To 1) As Variant
    
    If rg.Cells.CountLarge > 1 Then
        GetArrayFromRange = rg.Value
    Else
        arr(1, 1) = rg.Value
        GetArrayFromRange = arr
    End If
End Function
Function GetUsedRange(sh As Worksheet, Optional origin As Boolean = True, Optional checkTables As Boolean = True, Optional CheckPivots As Boolean = True) As Range
'***** Returns the used range in a sheet that contains data
'   Optional parameters:
'       origin(true)        starts at sheet(1,1) - else at usedrange(1,1)
'       checkTables(true)   include empty tables/listobject rows
'       checkPivots(true)   include empty Pivot columns / Rows
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim temp As Worksheet
    Dim sheetChanged As Boolean, modeChanged As Boolean
    Dim cnt As Long, i As Long
    Dim uRg As Range
    Dim lo As ListObject
    Dim pt As PivotTable
    
    If sh Is Nothing Then Exit Function
 
    With sh.UsedRange
        If WorksheetFunction.CountA(.Cells) > 0 Then
                cnt = .rows.count
                i = 0
                Do Until WorksheetFunction.CountA(.rows(cnt - i)) > 0
                    i = i + 1
                Loop
                lastRow = cnt - i
                cnt = .Columns.count
                i = 0
                Do Until WorksheetFunction.CountA(.Columns(cnt - i)) > 0
                    i = i + 1
                Loop
                lastColumn = cnt - i
            Else
                lastRow = 1
                lastColumn = 1
        End If
    End With
    
    lastRow = lastRow + sh.UsedRange.row - 1
    lastColumn = lastColumn + sh.UsedRange.column - 1
    
    If checkTables Then
        For Each lo In sh.ListObjects                           'Check for Tables with empty rows
            With lo.DataBodyRange
                If .Cells(.rows.count, 1).row > lastRow Then lastRow = .Cells(.rows.count, 1).row
                If .Cells(1, .Columns.count).column > lastColumn Then lastColumn = .Cells(1, .Columns.count).column
            End With
        Next lo
    End If
        
    If CheckPivots Then
        For Each pt In sh.PivotTables                            'Check for Pivots with empty columns
            With pt.TableRange1
                If .Cells(.rows.count, 1).row > lastRow Then lastRow = .Cells(.rows.count, 1).row
                If .Cells(1, .Columns.count).column > lastColumn Then lastColumn = .Cells(1, .Columns.count).column
            End With
        Next pt
    End If
    
    If origin Then
        Set GetUsedRange = Range(sh.Cells(1, 1), sh.Cells(lastRow, lastColumn))
    Else
        Set GetUsedRange = Range(sh.UsedRange.Cells(1, 1), sh.Cells(lastRow, lastColumn))
    End If
    
End Function

Public Function CombineTwoDArrays(Arr1 As Variant, _
    Arr2 As Variant) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TwoArraysToOneArray
' This takes two 2-dimensional arrays, Arr1 and Arr2, and
' returns an array combining the two. The number of Rows
' in the result is NumRows(Arr1) + NumRows(Arr2). Arr1 and
' Arr2 must have the same number of columns, and the result
' array will have that many columns. All the LBounds must
' be the same. E.g.,
' The following arrays are legal:
'        Dim Arr1(0 To 4, 0 To 10)
'        Dim Arr2(0 To 3, 0 To 10)
'
' The following arrays are illegal
'        Dim Arr1(0 To 4, 1 To 10)
'        Dim Arr2(0 To 3, 0 To 10)
'
' The returned result array is Arr1 with additional rows
' appended from Arr2. For example, the arrays
'    a    b        and     e    f
'    c    d                g    h
' become
'    a    b
'    c    d
'    e    f
'    g    h
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''
    ' Upper and lower bounds of Arr1.
    '''''''''''''''''''''''''''''''''
    Dim LBoundRow1 As Long
    Dim UBoundRow1 As Long
    Dim LBoundCol1 As Long
    Dim UBoundCol1 As Long
    
    '''''''''''''''''''''''''''''''''
    ' Upper and lower bounds of Arr2.
    '''''''''''''''''''''''''''''''''
    Dim LBoundRow2 As Long
    Dim UBoundRow2 As Long
    Dim LBoundCol2 As Long
    Dim UBoundCol2 As Long
    
    '''''''''''''''''''''''''''''''''''
    ' Upper and lower bounds of Result.
    '''''''''''''''''''''''''''''''''''
    Dim LBoundRowResult As Long
    Dim UBoundRowResult As Long
    Dim LBoundColResult As Long
    Dim UBoundColResult As Long
    
    '''''''''''''''''
    ' Index Variables
    '''''''''''''''''
    Dim RowNdx1 As Long
    Dim ColNdx1 As Long
    Dim RowNdx2 As Long
    Dim ColNdx2 As Long
    Dim RowNdxResult As Long
    Dim ColNdxResult As Long
    
    
    '''''''''''''
    ' Array Sizes
    '''''''''''''
    Dim NumRows1 As Long
    Dim NumCols1 As Long
    
    Dim NumRows2 As Long
    Dim NumCols2 As Long
    
    Dim NumRowsResult As Long
    Dim NumColsResult As Long
    
    Dim Done As Boolean
    Dim result() As Variant
    Dim ResultTrans() As Variant
    
    Dim V As Variant
    
    
    '''''''''''''''''''''''''''''''
    ' Ensure that Arr1 and Arr2 are
    ' arrays.
    ''''''''''''''''''''''''''''''
    If (IsArray(Arr1) = False) Or (IsArray(Arr2) = False) Then
        CombineTwoDArrays = Null
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''
    ' Ensure both arrays are allocated
    ' two dimensional arrays.
    ''''''''''''''''''''''''''''''''''
    If (NumberOfArrayDimensions(Arr1) <> 2) Or (NumberOfArrayDimensions(Arr2) <> 2) Then
        CombineTwoDArrays = Null
        Exit Function
    End If
       
    '''''''''''''''''''''''''''''''''''''''
    ' Ensure that the LBound and UBounds
    ' of the second dimension are the
    ' same for both Arr1 and Arr2.
    '''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''
    ' Get the existing bounds.
    ''''''''''''''''''''''''''
    LBoundRow1 = LBound(Arr1, 1)
    UBoundRow1 = UBound(Arr1, 1)
    
    LBoundCol1 = LBound(Arr1, 2)
    UBoundCol1 = UBound(Arr1, 2)
    
    LBoundRow2 = LBound(Arr2, 1)
    UBoundRow2 = UBound(Arr2, 1)
    
    LBoundCol2 = LBound(Arr2, 2)
    UBoundCol2 = UBound(Arr2, 2)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get the total number of rows for the result
    ' array.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    NumRows1 = UBoundRow1 - LBoundRow1 + 1
    NumCols1 = UBoundCol1 - LBoundCol1 + 1
    NumRows2 = UBoundRow2 - LBoundRow2 + 1
    NumCols2 = UBoundCol2 - LBoundCol2 + 1
    
    '''''''''''''''''''''''''''''''''''''''''
    ' Ensure the number of columns are equal.
    '''''''''''''''''''''''''''''''''''''''''
    If NumCols1 <> NumCols2 Then
        CombineTwoDArrays = Null
        Exit Function
    End If
    
    NumRowsResult = NumRows1 + NumRows2
    
    '''''''''''''''''''''''''''''''''''''''
    ' Ensure that ALL the LBounds are equal.
    ''''''''''''''''''''''''''''''''''''''''
    If (LBoundRow1 <> LBoundRow2) Or _
        (LBoundRow1 <> LBoundCol1) Or _
        (LBoundRow1 <> LBoundCol2) Then
        CombineTwoDArrays = Null
        Exit Function
    End If
    '''''''''''''''''''''''''''''''
    ' Get the LBound of the columns
    ' of the result array.
    '''''''''''''''''''''''''''''''
    LBoundColResult = LBoundRow1
    '''''''''''''''''''''''''''''''
    ' Get the UBound of the columns
    ' of the result array.
    '''''''''''''''''''''''''''''''
    UBoundColResult = UBoundCol1
    
    UBoundRowResult = LBound(Arr1, 1) + NumRows1 + NumRows2 - 1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Redim the Result array to have number of rows equal to
    ' number-of-rows(Arr1) + number-of-rows(Arr2)
    ' and number-of-columns equal to number-of-columns(Arr1)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim result(LBoundRow1 To UBoundRowResult, LBoundColResult To UBoundColResult)
    
    RowNdxResult = LBound(result, 1) - 1
    
    Done = False
    Do Until Done
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Copy elements of Arr1 to Result
        ''''''''''''''''''''''''''''''''''''''''''''
        For RowNdx1 = LBound(Arr1, 1) To UBound(Arr1, 1)
            RowNdxResult = RowNdxResult + 1
            For ColNdx1 = LBound(Arr1, 2) To UBound(Arr1, 2)
                V = Arr1(RowNdx1, ColNdx1)
                result(RowNdxResult, ColNdx1) = V
            Next ColNdx1
        Next RowNdx1
    
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Copy elements of Arr2 to Result
        '''''''''''''''''''''''''''''''''''''''''''''
        For RowNdx2 = LBound(Arr2, 1) To UBound(Arr2, 1)
            RowNdxResult = RowNdxResult + 1
            For ColNdx2 = LBound(Arr2, 2) To UBound(Arr2, 2)
                V = Arr2(RowNdx2, ColNdx2)
                result(RowNdxResult, ColNdx2) = V
            Next ColNdx2
        Next RowNdx2
       
        If RowNdxResult >= UBound(result, 1) + (LBoundColResult = 1) Then
            Done = True
        End If
    '''''''''''''
    ' End Of Loop
    '''''''''''''
    Loop
    '''''''''''''''''''''''''
    ' Return the Result
    '''''''''''''''''''''''''
    CombineTwoDArrays = result

End Function

Sub PutTextToClipboard(text As String)
'Puts text into clipboard
    Dim oClipboard As MSForms.DataObject
    
    Set oClipboard = New MSForms.DataObject
    oClipboard.SetText text
    oClipboard.PutInClipboard

End Sub
Function GetTextFromClipboard() As String
'Puts text into clipboard
    Dim oClipboard As MSForms.DataObject
    
    Set oClipboard = New MSForms.DataObject
    oClipboard.GetFromClipboard
    
    GetTextFromClipboard = oClipboard.GetText()

End Function
Sub ClearClipBoard()                         'safety
    Application.CutCopyMode = False
End Sub
Function decodeInParentheses(inp As String, Optional delimiter As String = ",", Optional paranthesesChars As String = "()", Optional removeParameterNames As Boolean = True) As String()
'Decode the first occurrence enclosed in parantheses (or end of string), delimited by delimiter.
'All in parameter quoted or in-parantheses parts are returned as-is.
'Ubound = 0 .... indicates an invalid input string
    Dim i As Long, pLevel As Long, cnt As Long, pos As Long
    Dim quoted As Boolean
    Dim quote As String, char As String, parm As String
    Dim parms() As String
    
    quote = Chr(34)
    
    If Len(paranthesesChars) <> 2 Then paranthesesChars = "()"
    
    ReDim parms(0 To 0)

    For i = 1 To Len(inp)
    
        char = Mid(inp, i, 1)
        
        If char = quote Then
            quoted = Not quoted
            parm = parm & char
            GoTo Iterate
        End If
        
        If quoted Then
            If pLevel > 0 Then parm = parm & char
            GoTo Iterate
        End If
            
        If char = Left(paranthesesChars, 1) Then
            pLevel = pLevel + 1
            If pLevel = 1 Then
                parm = ""
                GoTo Iterate
            End If
        End If
            
        If char = Right(paranthesesChars, 1) Then
            pLevel = pLevel - 1
            If pLevel = 0 Then GoTo WriteParm
        End If
        
        If pLevel = 0 Then GoTo Iterate
        
        If quoted Or pLevel > 1 Then
            parm = parm & char
            GoTo Iterate
        Else
            If char <> "," Then
                parm = parm & char
                GoTo Iterate
            End If
        End If
WriteParm:
        cnt = cnt + 1
        ReDim Preserve parms(0 To cnt)
        parms(cnt) = trim(parm)
        parm = ""
        If pLevel = 0 Then Exit For
        
Iterate:
    
    Next i
    
    If Len(parm) > 0 Then GoTo WriteParm
    
    If removeParameterNames Then
        For i = 0 To cnt
            pos = InStr(parms(i), ":=")
            If pos > 0 Then parms(i) = Mid(parms(i), pos + 2)
        Next i
    End If
    
    decodeInParentheses = parms

End Function
Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
   Dim i As Long
   Dim j As Long
   
   If Not UCase(Mid(Application.OperatingSystem, 1, 3)) = "WIN" Then Exit Function
 
   i = GetForegroundWindow
 
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
 
      i = j
   End If
 
   GetActiveWindowTitle = GetWindowTitle(i)
End Function
Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim s As String
    
   If Not UCase(Mid(Application.OperatingSystem, 1, 3)) = "WIN" Then Exit Function
 
   l = GetWindowTextLength(hwnd)
   s = Space$(l + 1)
 
   GetWindowText hwnd, s, l + 1
 
   GetWindowTitle = Left$(s, l)
End Function

Function VbeHasFocus() As Boolean
'Check if VBE has focus
    
    VbeHasFocus = InStr(GetActiveWindowTitle(False), "Visual Basic") > 0
    
End Function
Sub SwitchToExcel()
        If VbeHasFocus Then Call SendKeys("%{F11}", True)
End Sub
Sub SwitchToVBE()
        If Not VbeHasFocus Then Call SendKeys("%{F11}", True)
End Sub
Function getActiveVBECodeLine(ByRef codeLine As String, ByRef position As Long, ByRef length As Long) As Boolean
    Dim activePane As VBIDE.codePane
    Dim codeModule As VBIDE.codeModule
    Dim startLine As Long, endLine As Long, endColumn As Long
    
    On Error GoTo EarlyExit
    
    Set activePane = Application.VBE.ActiveCodePane
    
    Set codeModule = activePane.codeModule
    
    activePane.GetSelection startLine, position, endLine, endColumn
    
    length = endColumn - position
    
    codeLine = codeModule.lines(startLine, 1)

    getActiveVBECodeLine = True

EarlyExit:

End Function

Function GetFormula(cl As Range) As String
'Get formula even when cl is on protected sheet and formulahidden is true, which would cause an error

    If cl.Parent.ProtectContents And cl.FormulaHidden Then
        cl.FormulaHidden = False
        GetFormula = cl.Formula
        cl.FormulaHidden = True
    Else
        GetFormula = cl.Formula
    End If

End Function

Public Function CLetter(colNum As Long) As String
'Converts an Excel Column Number to a Column Letter
        If colNum > 26 Then
            CLetter = Chr(Int((colNum - 1) / 26) + 64) & Chr(((colNum - 1) Mod 26) + 65)
        Else
            CLetter = Chr(colNum + 64)
    End If
End Function
Public Function NumberOfArrayDimensions(arr As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Integer
Dim res As Integer
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    res = UBound(arr, Ndx)
Loop Until Err.Number <> 0

NumberOfArrayDimensions = Ndx - 1

End Function

Function getCallParametersFromString(stringToTest As String, functionIdentifiers As Variant, Optional ignoreCase As Boolean = True, Optional maxFinds As Long) As parameterInfo()
'Tests a string for the existance of functionIdentifiers (string or Array of strings) and returns an 1-based string array of all parameters for each find as well as the position of the find
'Returns a 0-based array in case of no find or error
    Dim regex As New RegExp
    Dim res As MatchCollection
    Dim mtch As Match
    Dim inputString As String, pattern As String, line As String
    Dim sp As Variant
    Dim result() As parameterInfo
    Dim i As Long, j As Long, start As Long
    
    If Not IsArray(functionIdentifiers) Then                                'Create pattern for string or array of strings with or function
        pattern = functionIdentifiers & "\("
    Else
        pattern = Join(functionIdentifiers, "\(|") & "\("
    End If
    
    With regex                                                              'find occurrebces of functionidentifier(s) with adjacent (
        .ignoreCase = True
        .Global = True
        .pattern = pattern
        Set res = .Execute(stringToTest)
    End With
        
    If res.count = 0 Then GoTo ErrorExit                                    'No find (or any error) - return Array with Ubound = 0
    
    With res
        If maxFinds = 0 Then                                                'Set max number of finds to evaluate
            maxFinds = .count
        Else
            maxFinds = Application.Min(maxFinds, .count)
        End If
        
        ReDim result(1 To maxFinds)
        
        For i = 1 To maxFinds                                              'For each match set parameterInfo into Array
            Set mtch = .Item(i - 1)
            With mtch
                result(i).position = .FirstIndex + 1                       'Starting position
                line = Mid(stringToTest, .FirstIndex + 1)
                start = InStr(line, "(")
                result(i).funCall = Left(line, start - 1)                  'Call that was actually matched
                line = Mid(line, InStr(line, "("))
                result(i).parms = decodeInParentheses(line)                'all parameters as returned by function decodeInParentheses
            End With
        Next i
    End With
    
    getCallParametersFromString = result
    
    Exit Function
            
ErrorExit:
    ReDim getCallParametersFromString(0)
    
End Function
Function IsLiteral(ByVal str As String, Optional doTrim As Boolean = True) As Boolean
    If doTrim Then str = trim(str)
    If Left(str, 1) = Chr(34) And Right(str, 1) = Chr(34) Then IsLiteral = True
End Function
Function TrimAll(ByVal str As String, Optional chrs As String = " ", Optional LTrim As Boolean = True, Optional RTrim As Boolean = True) As String
' Trims string for all charcters in chrs, optional lefttrim and/or righttrim
    If LTrim Then
        Do Until InStr(chrs, Left(str, 1)) = 0 Or Len(str) = 0
            str = Mid(str, 2, Len(str))
        Loop
    End If
    If RTrim Then
        Do Until InStr(chrs, Right(str, 1)) = 0 Or Len(str) = 0
            str = Left(str, Len(str) - 1)
        Loop
    End If
    TrimAll = str
End Function
Function ArrayToRange(anchor As Range, arr As Variant, Optional rws As Long = 0, Optional cols As Long = 0, Optional noWrapText As Boolean) As Range
'Places values of an array into a range at anchor (optional only rws (number of rows) and cols (number of columns) and returns the range of the placed array
'Works with one and two-dimensional arrays
    Dim sh As Worksheet
    Dim rg As Range
    Dim filterArray()
    Dim currentFiltRange As String
    Dim modeChanged As Boolean
    Dim rowCount As Long, colCount As Long
    Dim dummy As Variant
    
    If IsArray(arr) Then

        On Error Resume Next
        
        dummy = UBound(arr, 2)
            
        If Err.Number = 0 Then                                          'Array has (at least) 2 dimensions, function will fail (do nothing) with more than two dimensions
            
            On Error GoTo ErrorExit
            
            rowCount = UBound(arr, 1) - LBound(arr, 1) + 1
            If rws = 0 Then
                rws = rowCount
            Else
                rws = Application.Min(rws, rowCount)
            End If
            
            colCount = UBound(arr, 2) - LBound(arr, 2) + 1
            If cols = 0 Then
                cols = colCount
            Else
                cols = Application.Min(cols, colCount)
            End If
            
        Else                                                           'Array has just one dimension - will be placed horizontally (1 row)
        
            rws = 1
            
            colCount = UBound(arr) - LBound(arr) + 1
            If cols = 0 Then
                cols = colCount
            Else
                cols = Application.Min(cols, colCount)
            End If
        
        End If

    Else                                                            'not an array - asssume its just a value - else function will fail (do nothing)
    
        rows = 1
        cols = 1
    
    End If
    
    On Error GoTo ErrorExit
    
    Set rg = anchor.Resize(rws, cols)
    
    With rg
        If .Parent.FilterMode Then                                          'Sheet is filtered
            If .rows.count = CountVisibleRows(rg) Then                              'but the target range is not affected
                 rg = arr
            Else
            'Remove/restore Autofilter before executing the ArraytoRange
'                modeChanged = setFastMode
                
                Set sh = rg.Parent
                Call SaveAutoFilter(sh, currentFiltRange, filterArray)
                sh.AutoFilterMode = False
                rg = arr
                Call RestoreAutofilter(sh, currentFiltRange, filterArray)
                
'                If modeChanged Then Call ResetFastMode
            End If
        Else                                                                'No active filter - just put the array
            rg = arr
        End If
    End With
    
    If noWrapText Then rg.WrapText = False
    
    Set ArrayToRange = rg
ErrorExit:

End Function

Public Function IsVector(arg) As Boolean
'"vector" = array of one and only one dimension
    If IsObject(arg) Then
        Exit Function
    End If

    If Not IsArray(arg) Then
        Exit Function
    End If

    If ArrayHasDimension(arg, 1) Then
        IsVector = Not ArrayHasDimension(arg, 2)
    End If
End Function
Sub SaveAutoFilter(ws As Worksheet, ByRef currentFiltRange As String, ByRef filterArray)
    Dim col As Long

    'Capture AutoFilter settings
    If ws.AutoFilterMode = True Then
        With ws.AutoFilter
            currentFiltRange = .Range.Address
            With .Filters
                ReDim filterArray(1 To .count, 1 To 3)
                For col = 1 To .count
                    With .Item(col)
                        If .On Then
                            filterArray(col, 1) = .Criteria1
                            If .Operator Then
                                filterArray(col, 2) = .Operator
                                If .Operator = xlAnd Or .Operator = xlOr Then
                                    filterArray(col, 3) = .Criteria2
                                End If
                            End If
                        End If
                    End With
                Next col
            End With
        End With
    End If
End Sub

Sub RestoreAutofilter(ws As Worksheet, currentFiltRange As String, filterArray)
    Dim col As Long

    'Restore Filter settings
    If Not currentFiltRange = "" Then
        ws.Range(currentFiltRange).AutoFilter
        For col = 1 To UBound(filterArray, 1)
            If Not IsEmpty(filterArray(col, 1)) Then
                If filterArray(col, 2) Then
                    'check if Criteria2 exists and needs to be populated
                    If filterArray(col, 2) = xlAnd Or filterArray(col, 2) = xlOr Then
                        ws.Range(currentFiltRange).AutoFilter Field:=col, _
                        Criteria1:=filterArray(col, 1), _
                        Operator:=filterArray(col, 2), _
                        Criteria2:=filterArray(col, 3)
                    Else
                        ws.Range(currentFiltRange).AutoFilter Field:=col, _
                        Criteria1:=filterArray(col, 1), _
                        Operator:=filterArray(col, 2)
                    End If
                Else
                    ws.Range(currentFiltRange).AutoFilter Field:=col, _
                    Criteria1:=filterArray(col, 1)
                End If
            End If
        Next col
    End If
End Sub


Public Function IsMatrix(arg) As Boolean
'"matrix" = array of two and only two dimensions
    If IsObject(arg) Then
        Exit Function
    End If

    If Not IsArray(arg) Then
        Exit Function
    End If

    If ArrayHasDimension(arg, 2) Then
        IsMatrix = Not ArrayHasDimension(arg, 3)
    End If
End Function
Public Function IsInteger(vl As String) As Boolean
    On Error GoTo EarlyExit
    IsInteger = CStr(CInt(vl)) = trim(vl)
EarlyExit:
End Function

Function CountVisibleRows(InpRange As Range) As Long
'Counts visible rows in a range containing Hidden Rows & Columns; 0 if no cells are visible
    Dim rg As Range
    On Error GoTo EarlyExit
    Set rg = InpRange.SpecialCells(xlCellTypeVisible)
    If rg Is Nothing Then GoTo EarlyExit
    CountVisibleRows = InpRange.Columns(rg.column - InpRange.column + 1).SpecialCells(xlCellTypeVisible).count
EarlyExit:
End Function

Public Function ArrayHasDimension(arr, dimNum As Long) As Boolean
'Tests an array to see if it extends to a given dimension
    Debug.Assert IsArray(arr)
    Debug.Assert dimNum > 0

    'Note that it is possible for a VBA array to have no dimensions (i.e.
    ''LBound' raises an error even on the first dimension). This happens
    'with "unallocated" (borrowing Chip Pearson's terminology; see
    'http://www.cpearson.com/excel/VBAArrays.htm) dynamic arrays -
    'essentially arrays that have been declared with 'Dim arr()' but never
    'sized with 'ReDim', or arrays that have been deallocated with 'Erase'.

    On Error Resume Next
        Dim lb As Long
        lb = LBound(arr, dimNum)

        'No error (0) - array has given dimension
        'Subscript out of range (9) - array doesn't have given dimension
        ArrayHasDimension = (Err.Number = 0)

        'Debug.Assert (err.number = ERR_VBA_NONE Or err.number = ERR_VBA_SUBSCRIPT_OUT_OF_RANGE)
    On Error GoTo 0
End Function

Function spliceMatrix(arr As Variant, rows As Variant, cols As Variant) As Variant
'Accepts a 2d array and 2 1d arrays of row/ col nbrs; returns the requested rows/columns
    Dim temp As Variant
    Dim i As Long, j As Long, numRowsX As Long, numRowsY As Long, lbX As Long, lbY As Long, lbR As Long, ubR As Long, lbC As Long, ubC As Long
    
    Err.Clear
    On Error GoTo ErrorExit
    
    lbX = LBound(arr, 1)
    lbY = LBound(arr, 2)
    lbR = LBound(rows)
    ubR = UBound(rows)
    lbC = LBound(cols)
    ubC = UBound(cols)
    numRowsX = UBound(rows) - LBound(rows) + 1
    numRowsY = UBound(cols) - LBound(cols) + 1
    
    ReDim temp(lbX To lbX + numRowsX - 1, lbY To lbY + numRowsY - 1)
    
    For i = lbR To ubR
        For j = lbC To ubC
            temp(i + lbX - lbR, j + lbY - lbC) = arr(rows(i), cols(j))
        Next j
    Next i
    
    spliceMatrix = temp
    
    Exit Function
    
ErrorExit:
    
End Function
Function TwoDtoOneD(arr, Optional delim As String = "°", _
        Optional SkipBlankRows As Boolean = False) As Variant
'Transforms/Flatten a 2 dimensional array and transform it to one dimension row 1 col 1 - x , row 2 col 1-x etc
    On Error GoTo EarlyExit
    TwoDtoOneD = Split(Join2d(arr, delim, delim), delim)
EarlyExit:
    End Function
Function Join2d(ByRef InputArray As Variant, _
                       Optional RowDelimiter As String = vbCr, _
                       Optional FieldDelimiter = vbTab, _
                       Optional SkipBlankRows As Boolean = False _
                       ) As String

' Join up a 2-dimensional array into a string. Works like the standard
'  VBA.Strings.Join, for a 2-dimensional array.
' Note that the default delimiters are those inserted into the string
'  returned by ADODB.Recordset.GetString

On Error Resume Next

' Coding note: we're not doing any string-handling in VBA.Strings -
' allocating, deallocating and (especially!) concatenating are SLOW.
' We're using the VBA Join & Split functions ONLY. The VBA Join,
' Split, & Replace functions are linked directly to fast (by VBA
' standards) functions in the native Windows code. Feel free to
' optimise further by declaring and using the Kernel string functions
' if you want to.

' ** THIS CODE IS IN THE PUBLIC DOMAIN **
'   Nigel Heffernan   Excellerando.Blogspot.com

Dim i As Long
Dim j As Long

Dim i_lBound As Long
Dim i_uBound As Long
Dim j_lBound As Long
Dim j_uBound As Long

Dim arrTemp1() As String
Dim arrTemp2() As String

Dim strBlankRow As String

i_lBound = LBound(InputArray, 1)
i_uBound = UBound(InputArray, 1)

j_lBound = LBound(InputArray, 2)
j_uBound = UBound(InputArray, 2)

ReDim arrTemp1(i_lBound To i_uBound)
ReDim arrTemp2(j_lBound To j_uBound)

For i = i_lBound To i_uBound

    For j = j_lBound To j_uBound
        arrTemp2(j) = InputArray(i, j)
    Next j

    arrTemp1(i) = Join(arrTemp2, FieldDelimiter)

Next i

If SkipBlankRows Then

    If Len(FieldDelimiter) = 1 Then
        strBlankRow = String(j_uBound - j_lBound, FieldDelimiter)
    Else
        For j = j_lBound To j_uBound
            strBlankRow = strBlankRow & FieldDelimiter
        Next j
    End If

    Join2d = Replace(Join(arrTemp1, RowDelimiter), strBlankRow, RowDelimiter, "")
    i = Len(strBlankRow & RowDelimiter)

    If Left(Join2d, i) = strBlankRow & RowDelimiter Then
        Mid$(Join2d, 1, i) = ""
    End If

Else

    Join2d = Join(arrTemp1, RowDelimiter)

End If

Erase arrTemp1

End Function


Function SortArray(arr, Optional col As Integer = 1, Optional direction As Long = xlAscending, Optional lower As Boolean = False, Optional trim As Boolean = False, _
    Optional quiet As Boolean = False, Optional RemoveDuplicates As Boolean = False) As Variant
    'Sorts an 2-dimensional array by a key column
    'Currently supports just one sort column
    'parm lower indicates to change all key values to lower case - recommended for using it for binary searches
    'parm trim indicates to trim all empty rows at the end
    'Uses a (deep-hidden) sheet %work%
    Dim sh As Worksheet
    Dim shName As String
    Dim rg As Range, cl As Range
    Dim bookWasProtected As Boolean
    Dim tempArr(1 To 1, 1 To 1)
    Dim modeChanged As Boolean
    Dim i As Long
    Dim cbSaved As Boolean
    
    On Error GoTo ErrorExit
    Set cl = ActiveCell
      
    If Not IsArray(arr) Then
        If Not quiet Then Call MsgBox("Parameter for function sortArray is not an array", vbOKOnly + vbCritical, "Error")
        SortArray = arr
        Exit Function
    End If
    
'    cbSaved = SaveClipBoard
    
'    modeChanged = setFastMode
    
    Set sh = ActiveWorkbook.Worksheets(c_sortSheetname)
    
    sh.Activate
    sh.Cells.Clear
    
    If lower Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            arr(i, col) = LCase(arr(i, col))
        Next i
    End If
    
    sh.Range(Cells(1, 1), Cells(UBound(arr, 1) - LBound(arr, 1) + 1, UBound(arr, 2) - LBound(arr, 2) + 1)).NumberFormat = "@"
    Set rg = ArrayToRange(Cells(1, 1), arr)
    
    If RemoveDuplicates Then
        rg.RemoveDuplicates Columns:=Array(1), header:=xlNo
    End If
    
    rg.Sort key1:=rg.Cells(1, col), order1:=direction, header:=xlNo, MatchCase:=False, DataOption1:=xlSortNormal
    
    If trim Or RemoveDuplicates Then
        Set rg = sh.UsedRange
    End If
    
    On Error Resume Next
    If rg.Cells.CountLarge > 1 Then
        SortArray = rg.Value
    Else
        tempArr(1, 1) = rg.Value
        SortArray = tempArr
    End If
    
    sh.Cells.Clear
              
    cl.Parent.Activate
    cl.Select
    
'    If cbSaved Then Call RestoreClipBoard
    
'    If modeChanged Then Call ResetFastMode
    
    Exit Function

ErrorExit:
    If Not quiet Then Call MsgBox("Error in executing function sortArray", vbOKOnly + vbCritical, "Error")
    SortArray = arr
    cl.Parent.Activate
    cl.Select
    
'    If modeChanged Then Call ResetFastMode
End Function
Sub ToggleAddIn(Optional dummy As Boolean)
'Dim Only for members of the platform development team ... toggle Addin state of xlam

    If IsPlatformDeveloper Then
        ThisWorkbook.IsAddin = Not ThisWorkbook.IsAddin
    Else
        Call ShowMessage("system", "notAuthorized", smWarning)
    End If

End Sub
Function setFastMode(Optional changeCalculation As Boolean) As Boolean
'Changes enableEvents and screenUpdating. Calculation setting only on request

    If Application.EnableEvents Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        If changeCalculation Then Application.Calculation = xlCalculationManual
        setFastMode = True
    End If
    
End Function
Sub resetFastMode(Optional changeCalculation As Boolean)
'Resets enableEvents and screenUpdating. Calculation setting only on request
        
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If changeCalculation Then Application.Calculation = xlCalculationAutomatic

End Sub
Function AskForPassWord(Optional confirm As Boolean = True)
    Dim pwd1 As String, pwd2 As String
    
    pwd1 = InputBox(Prompt:=GetNlsText("AskPW", "enterpw"), Title:=GetNlsText("AskPW", "caption"))
    If pwd1 = "" Then Exit Function
    
    If confirm Then
        pwd2 = InputBox(Prompt:=GetNlsText("AskPW", "reenterpw"), Title:=GetNlsText("AskPW", "caption"))
        
         'Check if both the passwords are identical
        If pwd1 <> pwd2 Then
            MsgBox Prompt:=GetNlsText("AskPW", "mismatch"), Title:=GetNlsText("askPW", "caption")
            Exit Function
        End If
    End If
        
     AskForPassWord = pwd1
    
End Function
