Attribute VB_Name = "xlmod_ArrayProcessing_v_1_9"
'---Moved from TRepFuncPack
'Conventions:
'udf* - UserDefinedFunction
'arr* - Array Variable
'*buf* - buffer part for resulting variant
'
' ---Array Processing Functions----
'udf_AsideTwoArrays
'udf_AppendTwoArrays
'udf_JoinArrayByDelim
'udf_SplitArray
'udf_1dArrayTo2d
'udf_ArrayRemDuplicates
'udf_UnlofdFoldedArray
'udf_AddColumnsToArray
'udf_AddRowsToArray
'udf_GetArrayColumns
'udf_ExcludeArrayRows
'udf_ExcludeArrayColumns
'udf_ExcludeArrayElem1d
'udf_MatchInArray
'udf_SumCountIf_Array
'udf_FillArrayWith
'udf_Filter2dArrayBy1Column
'udf_Sort_1d_Array_as_text
'udf_Rebase_Array
'udf_CreateIndexColumn
'udf_CountCasesInArr


' ----Array Content Processing----
'udf_Trim_All_Items

' ----Array Interaction Functions----
'udf_IsInArray
'udf_GetDimension
'udf_Range_To_1D_Array
'udf_ColsRangesTo2dArray
'uds_Array_to_Range
'udf_Clipboard_to_Array
'udf_FileToArray_easy
'udf_RemCommasForCSV
'uds_ArrayToCSV

'------------------------------------
'--release history--
'1.5 - 22.03.2018:
'added udf_Filter2dArrayBy1Column
'1.6 - 25.06.2018:
'added udf_AddRowsToArray
'1.7 - 14.08.2018:
'added: udf_ExcludeArrayColumns
'added: udf_ExcludeArrayElem1d
'extended: udf_FillArrayWith - added replacement option
'1.8 - 28.08.2018
'changed: udf_AppendTwoArrays: added 1d version
'1.9 - 21.10.2018
'added: udf_CreateIndexColumn
'added: udf_CountCasesInArr

Public Enum BorderEnd
    ToBegining
    ToEnd
End Enum
Public Enum Direction
    Rows
    Columns
End Enum
Public Enum en_Position
    First
    Last
End Enum

'-----------------------------------
' --- Array Processing Functions----
'-----------------------------------
Function udf_AsideTwoArrays(arrArray1 As Variant, arrArray2 As Variant, Optional farg_ColBase = 1) As Variant

    If udf_GetDimension(arrArray1) = 1 Then arrArray1 = udf_1dArrayTo2d(arrArray1)
    If udf_GetDimension(arrArray2) = 1 Then arrArray2 = udf_1dArrayTo2d(arrArray2)

    If LBound(arrArray1, 1) <> LBound(arrArray2, 1) Or UBound(arrArray1, 1) <> UBound(arrArray2, 1) Then MsgBox ("Arrays bounds not match! Consider rebasement"): End
    
    lng_Dim2_Sum = (UBound(arrArray1, 2) - LBound(arrArray1, 2) + 1) + (UBound(arrArray2, 2) - LBound(arrArray2, 2) + 1)
    ReDim arrFinJoinedArray(LBound(arrArray1, 1) To UBound(arrArray1, 1), farg_ColBase To farg_ColBase + lng_Dim2_Sum - 1)

    For i = LBound(arrArray1, 1) To UBound(arrArray1, 1)
    x = farg_ColBase
        For j = LBound(arrArray1, 2) To UBound(arrArray1, 2)
            arrFinJoinedArray(i, x) = arrArray1(i, j)
            x = x + 1
        Next
        For j = LBound(arrArray2, 2) To UBound(arrArray2, 2)
            arrFinJoinedArray(i, x) = arrArray2(i, j)
            x = x + 1
        Next
    Next i

udf_AsideTwoArrays = arrFinJoinedArray
End Function
Function udf_AppendTwoArrays(arr1, arr2)

check_dim1 = udf_GetDimension(arr1)
check_dim2 = udf_GetDimension(arr2)

If check_dim1 <> check_dim2 Then MsgBox ("Arrays dimentions do not match!"): End

'1d array section
    If check_dim1 = 1 Then
    
    lng_EndUBound = LBound(arr1) + (UBound(arr1) - LBound(arr1)) + (UBound(arr2) - LBound(arr2) + 1)
    
    ReDim arr_OutputArray1d(LBound(arr1) To lng_EndUBound)
    
        For i = LBound(arr1) To UBound(arr1)
            arr_OutputArray1d(i) = arr1(i)
        Next
        For k = LBound(arr2) To UBound(arr2)
            arr_OutputArray1d(i) = arr2(k)
            i = i + 1
        Next
        
    udf_AppendTwoArrays = arr_OutputArray1d
    End If
    
'2d array section
    If check_dim1 = 2 Then
    
        If LBound(arr1, 2) <> LBound(arr2, 2) Or UBound(arr1, 2) <> UBound(arr2, 2) Then MsgBox ("Arrays bounds do not match! Consider rebasement"): End
    
        lng_EndUBound = LBound(arr1, 1) + (UBound(arr1, 1) - LBound(arr1, 1)) + (UBound(arr2, 1) - LBound(arr2, 1) + 1)
        ReDim arr_OutputArray2d(LBound(arr1, 1) To lng_EndUBound, LBound(arr1, 2) To UBound(arr1, 2))
                          
        For i = LBound(arr1, 1) To UBound(arr1, 1)
            For j = LBound(arr1, 2) To UBound(arr1, 2)
                arr_OutputArray2d(i, j) = arr1(i, j)
            Next
        Next
        For k = LBound(arr2, 1) To UBound(arr2, 1)
            For j = LBound(arr2, 2) To UBound(arr2, 2)
                arr_OutputArray2d(i, j) = arr2(k, j)
            Next
            i = i + 1
        Next
        
    udf_AppendTwoArrays = arr_OutputArray2d
    End If

End Function

Function udf_JoinArrayByDelim(farg_SourceArray, farg_Delimeter As String) As Variant
    
    ReDim arr_OutputArray(LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1))
    ind_ColsBase = LBound(farg_SourceArray, 2)
    For i = ind_ColsBase To UBound(farg_SourceArray, 1)
        arr_OutputArray(i) = farg_SourceArray(i, ind_ColsBase)
        For j = LBound(farg_SourceArray, 2) + 1 To UBound(farg_SourceArray, 2)
            arr_OutputArray(i) = arr_OutputArray(i) & farg_Delimeter & farg_SourceArray(i, j)
        Next
    Next
    
udf_JoinArrayByDelim = arr_OutputArray
End Function

Function udf_SplitArray(arrArray As Variant, strDelimiter As String, farg_ColsBase As Long) As Variant

ind_ArrDim = udf_GetDimension(arrArray)

If ind_ArrDim = 1 Then
    ReDim arrFinSplitArray(LBound(arrArray) To UBound(arrArray), _
                     farg_ColsBase + LBound(Split(arrArray(1), strDelimiter)) _
                     To farg_ColsBase + UBound(Split(arrArray(1), strDelimiter))) As Variant
    
    For i = LBound(arrArray) To UBound(arrArray)
        For j = farg_ColsBase + LBound(Split(arrArray(1), strDelimiter)) To farg_ColsBase + UBound(Split(arrArray(1), strDelimiter))
            arrFinSplitArray(i, j) = Split(arrArray(i), strDelimiter)(j - farg_ColsBase)
        Next j
    Next i
    
ElseIf ind_ArrDim = 2 Then
    ReDim arrFinSplitArray(LBound(arrArray) To UBound(arrArray), _
                     farg_ColsBase + LBound(Split(arrArray(1, 1), strDelimiter)) _
                     To farg_ColsBase + UBound(Split(arrArray(1, 1), strDelimiter))) As Variant
    
    For i = LBound(arrArray, 1) To UBound(arrArray, 1)
        For j = farg_ColsBase + LBound(Split(arrArray(1, 1), strDelimiter)) To farg_ColsBase + UBound(Split(arrArray(1, 1), strDelimiter))
            arrFinSplitArray(i, j) = Split(arrArray(i, 1), strDelimiter)(j - farg_ColsBase)
        Next j
    Next i
End If

udf_SplitArray = arrFinSplitArray
End Function

Function udf_1dArrayTo2d(farg_SourceArray) As Variant
    
    ReDim arr_OutputArray(LBound(farg_SourceArray) To UBound(farg_SourceArray), 1 To 1)
    For i = LBound(farg_SourceArray) To UBound(farg_SourceArray)
        arr_OutputArray(i, 1) = farg_SourceArray(i)
    Next
    
udf_1dArrayTo2d = arr_OutputArray
End Function

Function udf_ArrayRemDuplicates(arrInitArray As Variant) As Variant
   
    InitArrStart = LBound(arrInitArray)
    j = LBound(arrInitArray)
    
    ReDim arrFinArray(InitArrStart To j)
    
    arrFinArray(j) = arrInitArray(j)
    
    For i = LBound(arrInitArray) To UBound(arrInitArray)
        If Not udf_IsInArray(arrInitArray(i), arrFinArray) Then j = j + 1: ReDim Preserve arrFinArray(InitArrStart To j) As Variant: arrFinArray(j) = arrInitArray(i)
    Next

udf_ArrayRemDuplicates = arrFinArray
End Function

Function udf_UnlofdFoldedArray(arr_FoldedArray As Variant) As Variant

ReDim arr_FinArray(LBound(arr_FoldedArray) To UBound(arr_FoldedArray), LBound(arr_FoldedArray(LBound(arr_FoldedArray))) To UBound(arr_FoldedArray(LBound(arr_FoldedArray))))

For i = LBound(arr_FoldedArray) To UBound(arr_FoldedArray)

    For j = LBound(arr_FoldedArray(i)) To UBound(arr_FoldedArray(i))
        arr_FinArray(i, j) = arr_FoldedArray(i)(j)
    Next j

Next i

udf_UnlofdFoldedArray = arr_FinArray
End Function

Function udf_AddColumnsToArray(farg_SourceArray, farg_ColQty, farg_BorderToAdd As BorderEnd)
    
    arr_ArrayToProcess = farg_SourceArray
    If udf_GetDimension(arr_ArrayToProcess) = 1 Then arr_ArrayToProcess = udf_1dArrayTo2d(arr_ArrayToProcess)

    ReDim arr_IncreasedArray(LBound(arr_ArrayToProcess, 1) To UBound(arr_ArrayToProcess, 1) _
                                , LBound(arr_ArrayToProcess, 2) To UBound(arr_ArrayToProcess, 2) + farg_ColQty)
    If farg_BorderToAdd = ToBegining Then ind_OffsetInput = farg_ColQty Else farg_BorderToAdd = 0

    For i = LBound(arr_ArrayToProcess, 1) To UBound(arr_ArrayToProcess, 1)
        For j = LBound(arr_ArrayToProcess, 2) To UBound(arr_ArrayToProcess, 2)
            arr_IncreasedArray(i, j + ind_OffsetInput) = arr_ArrayToProcess(i, j)
        Next
    Next

udf_AddColumnsToArray = arr_IncreasedArray
End Function
Function udf_GetArrayColumns(farg_SourceArray, ParamArray farg_ColInds())

    If UBound(farg_ColInds) = 0 Then
        
        ReDim arr_OutputArray1d(LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1))
        For i = LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1)
            arr_OutputArray1d(i) = farg_SourceArray(i, farg_ColInds(0))
        Next
        arr_OutputArray = arr_OutputArray1d
    Else
        
        ReDim arr_OutputArray2d(LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1), 1 To 1 + UBound(farg_ColInds) - LBound(farg_ColInds))
        For i = LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1)
            For j = LBound(farg_ColInds) To UBound(farg_ColInds)
                arr_OutputArray2d(i, 1 + j) = farg_SourceArray(i, farg_ColInds(j))
            Next
        Next
        arr_OutputArray = arr_OutputArray2d
    End If
    
udf_GetArrayColumns = arr_OutputArray
End Function
Function udf_AddRowsToArray(arr_ProcessingArray, farg_NumberOfRows)
'finalization needed

    arr_ProcessingArray = Application.Transpose(arr_ProcessingArray)
    
    ReDim Preserve arr_ProcessingArray(LBound(arr_ProcessingArray, 1) To UBound(arr_ProcessingArray, 1) _
                            , LBound(arr_ProcessingArray, 2) To UBound(arr_ProcessingArray, 2) + farg_NumberOfRows)
                            
    arr_ProcessingArray = Application.Transpose(arr_ProcessingArray)
    
udf_AddRowsToArray = arr_ProcessingArray
End Function
Function udf_ExcludeArrayRows(farg_SourceArray, farg_NewBase, ParamArray farg_RowsToExclude())
    
    lng_ExclRowsCount = UBound(farg_RowsToExclude) - LBound(farg_RowsToExclude) + 1

    ReDim arr_OutputArray(farg_NewBase To farg_NewBase + UBound(farg_SourceArray, 1) - LBound(farg_SourceArray, 1) - lng_ExclRowsCount _
                            , LBound(farg_SourceArray, 2) To UBound(farg_SourceArray, 2))
    x = farg_NewBase
    For i = LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1)
        bln_skip = False
        For n = LBound(farg_RowsToExclude) To UBound(farg_RowsToExclude)
            If farg_RowsToExclude(n) = i Then bln_skip = True: Exit For
        Next
        
        If bln_skip = False Then
            For j = LBound(farg_SourceArray, 2) To UBound(farg_SourceArray, 2)
                arr_OutputArray(x, j) = farg_SourceArray(i, j)
            Next
            x = x + 1
        End If
    Next
    
udf_ExcludeArrayRows = arr_OutputArray
End Function
Function udf_ExcludeArrayColumns(farg_SourceArray, ParamArray farg_ColsToExclude())
    
    lng_ExclColumnCount = UBound(farg_ColsToExclude) - LBound(farg_ColsToExclude) + 1

    ReDim arr_OutputArray(LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1) _
                            , LBound(farg_SourceArray, 2) To UBound(farg_SourceArray, 2) - lng_ExclColumnCount)
    
    y = LBound(farg_SourceArray, 1)
    
    For j = LBound(farg_SourceArray, 2) To UBound(farg_SourceArray, 2)
        bln_skip = False
        For n = LBound(farg_ColsToExclude) To UBound(farg_ColsToExclude)
            If farg_ColsToExclude(n) = j Then bln_skip = True: Exit For
        Next
        
        If bln_skip = False Then
            For i = LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1)
                arr_OutputArray(i, y) = farg_SourceArray(i, j)
            Next
            y = y + 1
        End If
    Next
    
udf_ExcludeArrayColumns = arr_OutputArray
End Function
Function udf_ExcludeArrayElem1d(farg_SourceArray, ParamArray farg_ElemsToExclude())
    
    lng_ExclElemsCount = UBound(farg_ElemsToExclude) - LBound(farg_ElemsToExclude) + 1

    ReDim arr_OutputArray(LBound(farg_SourceArray) To UBound(farg_SourceArray) - lng_ExclElemsCount)
    
    x = LBound(farg_SourceArray)
    
    For i = LBound(farg_SourceArray) To UBound(farg_SourceArray)
        bln_skip = False
        For n = LBound(farg_ElemsToExclude) To UBound(farg_ElemsToExclude)
            If farg_ElemsToExclude(n) = i Then bln_skip = True: Exit For
        Next
        
        If bln_skip = False Then
            arr_OutputArray(x) = farg_SourceArray(i)
            x = x + 1
        End If
    Next
    
udf_ExcludeArrayElem1d = arr_OutputArray
End Function
Function udf_MatchInArray(farg_TargetArray, farg_StringToFind, Optional farg_FieldIndex = 1, Optional farg_LookInDim = 1, Optional farg_ReturnFromFieldID = -1) As Variant

ind_TargArrDim = udf_GetDimension(farg_TargetArray)
ind_MatchResult = Array(False)

If ind_TargArrDim = 1 Then
    
    For i = LBound(farg_TargetArray) To UBound(farg_TargetArray)
        If "" & farg_TargetArray(i) = "" & farg_StringToFind Then ind_MatchResult = i
    Next
    
ElseIf ind_TargArrDim = 2 Then
    
    If farg_LookInDim = 1 Then
        For i = LBound(farg_TargetArray, 1) To UBound(farg_TargetArray, 1)
            If "" & farg_TargetArray(i, farg_FieldIndex) = "" & farg_StringToFind Then ind_MatchResult = Array(i)
        Next
    ElseIf farg_LookInDim = 2 Then
        For i = LBound(farg_TargetArray, 2) To UBound(farg_TargetArray, 2)
            If "" & farg_TargetArray(farg_FieldIndex, i) = "" & farg_StringToFind Then ind_MatchResult = Array(i)
        Next
    End If
End If

If farg_ReturnFromFieldID <> -1 And ind_MatchResult(0) <> False Then
    If farg_LookInDim = 1 Then udf_MatchInArray = Array(ind_MatchResult(0), farg_TargetArray(ind_MatchResult(0), farg_ReturnFromFieldID))
    If farg_LookInDim = 2 Then udf_MatchInArray = Array(ind_MatchResult(0), farg_TargetArray(farg_ReturnFromFieldID, ind_MatchResult(0)))
Else
    udf_MatchInArray = Array(ind_MatchResult(0), False)
End If

End Function
Function udf_SumCountIf_Array(farg_SourceArray, farg_CondCol, Optional farg_DataCol _
                        , Optional farg_DoSum As Boolean = True, Optional farg_ResCol = 0, Optional farg_ResArrayBase = 1 _
                        , Optional farg_CountCases As Boolean = False, Optional fagr_CountOutputCol = 0)
    
    If farg_ResCol = 0 Then
        check_AddSumCol = 1
        ind_SumOutputCol = UBound(farg_SourceArray, 2) + 1
    Else
        check_AddSumCol = 0
        ind_SumOutputCol = farg_ResCol
    End If
    
    If farg_CountCases = True And fagr_CountOutputCol = 0 Then
        check_AddCountCol = 1
        ind_CountOutputCol = UBound(farg_SourceArray, 2) + check_AddSumCol * farg_DoSum * -1 + 1
    Else
        check_AddCountCol = 0
        ind_CountOutputCol = fagr_CountOutputCol
    End If
    
    x = farg_ResArrayBase
    ReDim arr_OutputArray(1 To 1, LBound(farg_SourceArray, 2) To UBound(farg_SourceArray, 2) + check_AddSumCol + check_AddCountCol) As Variant
    ind_ColBase = LBound(arr_OutputArray, 2)
    ind_ColLim = UBound(arr_OutputArray, 2)
    arr_OutputArray = Application.Transpose(arr_OutputArray)
    
    For i = LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1)

        ind_MatchOutArr = udf_MatchInArray(arr_OutputArray, farg_SourceArray(i, farg_CondCol), farg_CondCol, 2)
        If ind_MatchOutArr(0) = False Then
            
            ReDim Preserve arr_OutputArray(ind_ColBase To ind_ColLim, farg_ResArrayBase To x)
            For n = LBound(farg_SourceArray, 2) To UBound(farg_SourceArray, 2)
                arr_OutputArray(n, x) = farg_SourceArray(i, n)
            Next
            If farg_DoSum = True Then arr_OutputArray(ind_SumOutputCol, x) = farg_SourceArray(i, farg_DataCol)
            If farg_CountCases = True Then arr_OutputArray(ind_CountOutputCol, x) = 1
            x = x + 1
        Else
            If farg_DoSum = True Then arr_OutputArray(ind_SumOutputCol, ind_MatchOutArr(0)) = arr_OutputArray(ind_SumOutputCol, ind_MatchOutArr(0)) _
                                                            + farg_SourceArray(i, farg_DataCol)
            If farg_CountCases = True Then arr_OutputArray(ind_CountOutputCol, ind_MatchOutArr(0)) = arr_OutputArray(ind_CountOutputCol, ind_MatchOutArr(0)) + 1
        End If
    Next
    
    arr_OutputArray = Application.Transpose(arr_OutputArray)
   
udf_SumCountIf_Array = arr_OutputArray
End Function
Function udf_FillArrayWith(farg_ValueToFill, Optional farg_StringToReplace As String, Optional farg_Lbound, Optional farg_Ubound, Optional farg_SourceArray, Optional farg_ColumnToFill)

Dim arr_FillingArray() As Variant

If Not IsMissing(farg_SourceArray) Then
    arr_FillingArray = farg_SourceArray
    farg_Lbound = LBound(farg_SourceArray, 1): farg_Ubound = UBound(farg_SourceArray, 1)
Else
    ReDim arr_FillingArray(farg_Lbound, farg_Ubound)
End If

For i = farg_Lbound To farg_Ubound
    If farg_StringToReplace <> "" Then
        If farg_ColumnToFill > -1 Then arr_FillingArray(i, farg_ColumnToFill) = Replace(arr_FillingArray(i, farg_ColumnToFill), farg_StringToReplace, farg_ValueToFill) _
            Else arr_FillingArray(i) = Replace(arr_FillingArray(i), farg_StringToReplace, farg_ValueToFill)
    Else
        If farg_ColumnToFill > -1 Then arr_FillingArray(i, farg_ColumnToFill) = farg_ValueToFill Else arr_FillingArray(i) = farg_ValueToFill
    End If
Next
    
udf_FillArrayWith = arr_FillingArray
End Function
Function udf_Filter2dArrayBy1Column(farg_SourceArray, farg_ColIndex, farg_FilterValue, Optional farg_KeepHeader As Boolean)
    
    ind_ColBase = LBound(farg_SourceArray, 2)
    ind_ColLim = UBound(farg_SourceArray, 2)
    ind_RowBase = LBound(farg_SourceArray, 1)
    x = ind_RowBase 'moving index for rows
    
    ReDim arr_OutputArray(1 To 1, LBound(farg_SourceArray, 2) To UBound(farg_SourceArray, 2)) As Variant
    arr_OutputArray = Application.Transpose(arr_OutputArray)
    
    If farg_KeepHeader = True Then
        For n = ind_ColBase To ind_ColLim
            arr_OutputArray(n, x) = farg_SourceArray(LBound(farg_SourceArray, 1), n)
        Next
        x = x + 1
    End If
    
    For i = LBound(farg_SourceArray, 1) To UBound(farg_SourceArray, 1)
        If farg_SourceArray(i, farg_ColIndex) = farg_FilterValue Then
            ReDim Preserve arr_OutputArray(ind_ColBase To ind_ColLim, ind_RowBase To x)
            For n = ind_ColBase To ind_ColLim
                arr_OutputArray(n, x) = farg_SourceArray(i, n)
            Next
            x = x + 1
        End If
    Next
    arr_OutputArray = Application.Transpose(arr_OutputArray)
    
udf_Filter2dArrayBy1Column = arr_OutputArray
End Function
Function udf_Sort_1d_Array_as_text(farg_ArrayToSort As Variant)
    
    Dim arrOutput As Variant
    
    arrOutput = farg_ArrayToSort
    
    For i = LBound(arrOutput) To UBound(arrOutput)
        
        SwapCarr = i
        Do While SwapCarr > LBound(arrOutput)
        
            If UCase(arrOutput(SwapCarr)) <= UCase(arrOutput(SwapCarr - 1)) Then
            
                tmpHi = arrOutput(SwapCarr)
                tmpLow = arrOutput(SwapCarr - 1)
                arrOutput(SwapCarr - 1) = tmpHi
                arrOutput(SwapCarr) = tmpLow
                
            End If
            SwapCarr = SwapCarr - 1
        Loop
        
    'next i
    Next

udf_Array_1d_Sort_as_text = arrOutput

End Function
Function udf_Rebase_Array(farg_SourceArray, farg_NewRowBase)

lng_BaseDiff = LBound(farg_SourceArray) - farg_NewRowBase
ReDim arr_OutputArray(farg_NewRowBase To UBound(farg_SourceArray) - LBound(farg_SourceArray) + 1)

For i = LBound(farg_SourceArray) To UBound(farg_SourceArray)
    arr_OutputArray(i - lng_BaseDiff) = farg_SourceArray(i)
Next

udf_Rebase_Array = arr_OutputArray
End Function

'--------------------------
' ----Array Content Processing----
'--------------------------

Function udf_Trim_All_Items(farg_SourceArray)

arr_ProcessingArray = farg_SourceArray
lng_ArrDim = udf_GetDimension(arr_ProcessingArray)

If lng_ArrDim = 1 Then
    For i = LBound(arr_ProcessingArray) To UBound(arr_ProcessingArray)
        arr_ProcessingArray(i) = Trim(arr_ProcessingArray(i))
    Next
ElseIf lng_ArrDim = 2 Then
    For i = LBound(arr_ProcessingArray, 1) To UBound(arr_ProcessingArray, 1)
        For j = LBound(arr_ProcessingArray, 2) To UBound(arr_ProcessingArray, 2)
            arr_ProcessingArray(i, j) = Trim(arr_ProcessingArray(i, j))
        Next
    Next
End If

udf_Trim_All_Items = arr_ProcessingArray
End Function


'--------------------------
'----Array Interaction Functions----
'--------------------------

Function udf_IsInArray(stringToBeFound As Variant, arr As Variant) As Boolean
  udf_IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Function udf_GetDimension(farg_ArrayToCheck As Variant) As Long
    On Error GoTo Err
    Dim i As Long
    Dim tmp As Long
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(farg_ArrayToCheck, i)
    Loop
Err:
udf_GetDimension = i - 1

End Function

Function udf_Range_To_1D_Array(rngSourceRange As Range, Optional farg_Base As Long = 0) 'modified in v.3.3

    Dim tmpArray() As Variant
    Dim arrFinOneDimArray As Variant
    
    i = farg_Base
    
    For Each cell In rngSourceRange.Cells
        ReDim Preserve tmpArray(farg_Base To i)
        tmpArray(i) = cell.Value
        i = i + 1
    Next cell
    
udf_Range_To_1D_Array = tmpArray
End Function

Function udf_ColsRangesTo2dArray(farg_SourceRange, farg_RowsBase, farg_ColsBase, ParamArray farg_ColInds() As Variant) As Variant

    Dim arrFinJoinedArray As Variant
    Dim rng_CurrentRange As Range
    
    bufferArr = farg_SourceRange.Columns(farg_ColInds(0)).Value
      
    For i = LBound(farg_ColInds) + 1 To UBound(farg_ColInds)
        bufferArr = udf_AsideTwoArrays(bufferArr, farg_SourceRange.Columns(farg_ColInds(i)).Value, farg_ColsBase)
    Next
    
udf_ColsRangesTo2dArray = bufferArr
End Function
Function udf_CountCasesInArr(farg_TargetArray, farg_StringToFind, Optional farg_FieldIndex = 1, Optional farg_LookInDim = 1)

    ind_TargArrDim = udf_GetDimension(farg_TargetArray)
    lng_ResultCount = 0
    
    If ind_TargArrDim = 1 Then
        
        For i = LBound(farg_TargetArray) To UBound(farg_TargetArray)
            If "" & farg_TargetArray(i) = "" & farg_StringToFind Then ind_MatchResult = i: lng_ResultCount = lng_ResultCount = 1
        Next
        
    ElseIf ind_TargArrDim = 2 Then
        
        If farg_LookInDim = 1 Then
            For i = LBound(farg_TargetArray, 1) To UBound(farg_TargetArray, 1)
                If "" & farg_TargetArray(i, farg_FieldIndex) = "" & farg_StringToFind Then ind_MatchResult = Array(i): lng_ResultCount = lng_ResultCount + 1
            Next
        ElseIf farg_LookInDim = 2 Then
            For i = LBound(farg_TargetArray, 2) To UBound(farg_TargetArray, 2)
                If "" & farg_TargetArray(farg_FieldIndex, i) = "" & farg_StringToFind Then ind_MatchResult = Array(i): lng_ResultCount = lng_ResultCount + 1
            Next
        End If
    End If
   
udf_CountCasesInArr = lng_ResultCount
End Function
Function udf_CreateIndexColumn(farg_InputArray, farg_Position As en_Position, farg_IndDelim, ParamArray farg_ColumnsToIndex())

ReDim arr_ProcessingArray(LBound(farg_InputArray, 1) To UBound(farg_InputArray, 1), LBound(farg_InputArray, 2) To UBound(farg_InputArray, 2) + 1)

If farg_Position = First Then ind_col = 1: ind_shift = 1
If farg_Position = Last Then ind_col = UBound(farg_InputArray, 2) + 1: ind_shift = 0

For i_row = LBound(farg_InputArray, 1) To UBound(farg_InputArray, 1)
    str_index = ""
    For j = LBound(farg_ColumnsToIndex, 1) To UBound(farg_ColumnsToIndex, 1)
        str_index = str_index & farg_IndDelim & farg_InputArray(i_row, farg_ColumnsToIndex(j))
    Next
    
    arr_ProcessingArray(i_row, ind_col) = Right(str_index, Len(str_index) - 1)
    
    For k = LBound(farg_InputArray, 2) To UBound(farg_InputArray, 2)
        arr_ProcessingArray(i_row, k + ind_shift) = farg_InputArray(i_row, k)
    Next
Next

udf_CreateIndexColumn = arr_ProcessingArray
End Function
Sub uds_Array_to_Range(sarg_SourceArray, sarg_TargetCell, Optional sarg_Transpose As Boolean = False)
    
    ind_ArrayDimentions = udf_GetDimension(sarg_SourceArray)
    
    If ind_ArrayDimentions = 1 Then
        'if transpose False parce array to 1col/multi rows
        If sarg_Transpose = False Then
            sarg_TargetCell.Resize(UBound(sarg_SourceArray, 1) - LBound(sarg_SourceArray, 1) + 1 _
            , 1) = sarg_SourceArray
        'if transpose True parce array to 1row/multi col
        Else
            sarg_TargetCell.Resize(1, _
            UBound(sarg_SourceArray, 1) - LBound(sarg_SourceArray, 1) + 1) = sarg_SourceArray
        End If
    
    ElseIf ind_ArrayDimentions = 2 Then
        'if transpose False parce 1 dim to rows, 2 dim to cols
        If sarg_Transpose = False Then
            sarg_TargetCell.Resize(UBound(sarg_SourceArray, 1) - LBound(sarg_SourceArray, 1) + 1 _
            , UBound(sarg_SourceArray, 2) - LBound(sarg_SourceArray, 2) + 1) = sarg_SourceArray
        
        'if transpose True parce 1 dim to cols, 2 dim to rows
        Else
            sarg_TargetCell.Resize(UBound(sarg_SourceArray, 2) - LBound(sarg_SourceArray, 2) + 1 _
            , UBound(sarg_SourceArray, 1) - LBound(sarg_SourceArray, 1) + 1) = sarg_SourceArray
        End If

    End If
    
End Sub

Function udf_Clipboard_to_Array(Optional farg_LineDelimiter As String, Optional farg_TabDelimiter As String) As Variant

Dim DataObj As MSForms.DataObject
Set DataObj = New MSForms.DataObject
Dim arr_BufferByLineAndTabArray As Variant

DataObj.GetFromClipboard

'if No Delimiters presented (so no need in splitting clipboard content)
If farg_LineDelimiter = "" And farg_TabDelimiter = "" Then
    udf_Clipboard_to_Array = DataObj.GetText(1)
    Exit Function
End If

'if only Line Delimiter presented (processing one column array in the clipboard)
If farg_LineDelimiter <> "" And farg_TabDelimiter = "" Then
    Q = Array(1)
    udf_Clipboard_to_Array = Application.Transpose(Split(DataObj.GetText(1), farg_LineDelimiter))
    Exit Function
End If

'if only Tab Delimiter presented (processing one line array in the clipboard)
If farg_LineDelimiter = "" And farg_TabDelimiter <> "" Then
    udf_Clipboard_to_Array = Split(DataObj.GetText(1), farg_TabDelimiter)
    Exit Function
End If

'if both Line Delimiter and Tab Delimeter presented (processing table array in the clipboard)
If farg_LineDelimiter <> "" And farg_TabDelimiter <> "" Then
    
    arr_BufferByLineArray = Split(DataObj.GetText(1), farg_LineDelimiter)
    
    ReDim arr_BufferByLineAndTabArray(0 To UBound(arr_BufferByLineArray) - LBound(arr_BufferByLineArray))
    For i = LBound(arr_BufferByLineArray) To UBound(arr_BufferByLineArray)
        arr_BufferByLineAndTabArray(i) = Split(arr_BufferByLineArray(i), farg_TabDelimiter)
    Next i
    
    udf_Clipboard_to_Array = udf_UnlofdFoldedArray(arr_BufferByLineAndTabArray)

End If

End Function

Function udf_FileToArray_easy(farg_SourceFilePath, Optional farg_SheetIndex = 1)
    
    Set wb_SourceFile = Workbooks.Open(farg_SourceFilePath)
    arr_OutputArray = wb_SourceFile.Sheets(farg_SheetIndex).UsedRange.Value
    wb_SourceFile.Close Savechanges:=False
    
udf_FileToArray_easy = arr_OutputArray
End Function

Function udf_RemCommasForCSV(farg_SourceFilePath)
    For i = LBound(farg_SourceFilePath, 1) To UBound(farg_SourceFilePath, 1)
        
    Next
    For i = LBound(farg_SourceFilePath, 1) To UBound(farg_SourceFilePath, 1)
        For j = LBound(farg_SourceFilePath, 2) To UBound(farg_SourceFilePath, 2)
            farg_SourceFilePath(i, j) = Replace(farg_SourceFilePath(i, j), ",", ".")
        Next
    Next
    
udf_RemCommasForCSV = farg_SourceFilePath
End Function

Sub uds_ArrayToCSV(sarg_SourceArray, sarg_FileName, Optional sarg_Headers = "")

    Open sarg_FileName For Output As #1
    
    If UBound(sarg_SourceArray, 1) = 0 Then Close #1: Exit Sub
    If sarg_Headers <> "" Then Print #1, sarg_Headers
    
    For i = LBound(sarg_SourceArray, 1) To UBound(sarg_SourceArray, 1)
        
        str_CSVLine = sarg_SourceArray(i, LBound(sarg_SourceArray, 2))
        For j = LBound(sarg_SourceArray, 2) + 1 To UBound(sarg_SourceArray, 2)
            str_CSVLine = str_CSVLine & "," & Replace(sarg_SourceArray(i, j), ",", ".")
        Next
        Print #1, str_CSVLine
    
    Next i
    Close #1
    
End Sub
