Attribute VB_Name = "_VBALib_Array"
Option Explicit
Private Const NORMALIZE_LBOUND = 1
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)





Public Function ArrayRank(arr As Variant) As Integer
    ' Returns Rank (#Dimensions) of an Array
    Const VT_BYREF = &H4000& 'real VarType of Arg, like VarType(),...
    Dim ptr As Long, vType As Integer   'but Returns VT_BYREF bit too
    CopyMemory vType, arr, 2   
    If (vType And vbArray) = 0 Then Exit Function 'if not an array
    CopyMemory ptr, ByVal VarPtr(arr) + 8, 4 'get Address of SAFEARRAY Descriptor (stored in 2nd half of Variant Param) that received the Array
    ' see whether the routine was passed a Variant that contains an Array or an Array directly, if the latter we must set ptr to SA structure
    If (vType And VT_BYREF) Then CopyMemory ptr, ByVal ptr, 4 'ptr is a pointer to a pointer.
    ' get Address of SAFEARRAY Structure (stored in the Descriptor) > get 1st Word of SAFEARRAY Structure 
    If ptr Then CopyMemory ArrayRank, ByVal ptr, 2   'which holds # of Dimensions (if saAddr is non-zero)
End Function
Public Function ArrayLen(arr As Variant, Optional nDim As Integer = 1) As Long
    ' Returns #Elements in an Array of given Dimension
    ArrayLen = IIf(IsEmpty(arr), 0, UBound(arr,nDim) - LBound(arr,nDim) + 1)
End Function
Public Function ArrayIndexOf(arr As Variant, val As Variant) As Long
    ' Returns Index of given Value in given Array (or 1-less than the Lower Bound of the Array if the Value is not found in the Array).
    '    @param arr: The array to search through.
    '    @param val: The value to search for.
    Dim i As Long: ArrayIndexOf = LBound(arr) - 1
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then ArrayIndexOf = i: Exit Function
    Next
End Function
Public Function ArrayContains(arr As Variant, val As Variant) As Boolean
    ' Returns whether given Array contains given Value.
    '    @param arr: The array to search through.
    '    @param val: The value to search for.
    ArrayContains = (ArrayIndexOf(arr, val) >= LBound(arr))
End Function





Public Function ArraySubset(arr() As Variant, Optional r1 As Long = -1, Optional r2 As Long = -1, Optional c1 As Long = -1, Optional c2 As Long = -1) As Variant()
    ' Returns Subset of given [1/2]-Dimensional Array specified by the given bounds. Returned Array has Lower Bound of 1.
    '    @param arr: The array to process.
    '    @param r1: The index of the first element to be extracted from the first dimension of the array. If not given, defaults to the lower bound of the first dimension.
    '    @param r2: The index of the last element to be extracted from the first dimension of the array. If not given, defaults to the upper bound of the first dimension.
    '    @param c1: The index of the first element to be extracted from the second dimension of the array. If not given, defaults to the lower bound of the second dimension.
    '    @param c2: The index of the last element to be extracted from the second dimension of the array. If not given, defaults to the upper bound of the second dimension.
    Dim arr2() As Variant, i As Long, j As Long
    If r1 < 0 Then r1 = LBound(arr, 1)
    If r2 < 0 Then r2 = UBound(arr, 1)
    Select Case ArrayRank(arr)
        Case 1
            If c1 >= 0 Or c2 >= 0 Then Err.Raise 32000, Description:="Too many array dimensions passed to ArraySubset."
            ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + r2 - r1)
            For i = r1 To r2
                arr2(i - r1 + NORMALIZE_LBOUND) = arr(i)
            Next
        Case 2
            If c1 < 0 Then c1 = LBound(arr, 2)
            If c2 < 0 Then c2 = UBound(arr, 2)
            ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + r2 - r1, NORMALIZE_LBOUND To NORMALIZE_LBOUND + c2 - c1)
            For i = r1 To r2
            For j = c1 To c2
                arr2(i - r1 + NORMALIZE_LBOUND, j - c1 + NORMALIZE_LBOUND) = arr(i, j)
            Next j, i
        Case Else
            Err.Raise 32000, Description:="Can only take a subset of a 1- or 2-dimensional array."
    End Select
    ArraySubset = arr2
End Function
Public Function NormalizeArray(arr As Variant) As Variant
    ' Returns a single-dimension array with lower bound 1, if given a one-dimensional array with any lower bound or a two-dimensional array with one dimension having only one element.  This function will always return a copy of the given array.
    If ArrayLen(arr) = 0 Then NormalizeArray = Array(): Exit Function
    Dim arr2() As Variant, nItems As Long, i As Long
    Select Case ArrayRank(arr)
        Case 1
            If LBound(arr) = NORMALIZE_LBOUND Then
                NormalizeArray = arr
            Else
                nItems = ArrayLen(arr)
                ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1)
                For i = NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1
                    arr2(i) = arr(i + LBound(arr) - NORMALIZE_LBOUND)
                Next
                NormalizeArray = arr2
            End If
        Case 2
            If LBound(arr, 1) = UBound(arr, 1) Then
                ' Copy values from array's second dimension
                nItems = ArrayLen(arr, 2)
                ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1)
                For i = NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1
                    arr2(i) = arr(LBound(arr, 1), i + LBound(arr, 2) - NORMALIZE_LBOUND)
                Next
                NormalizeArray = arr2
            ElseIf LBound(arr, 2) = UBound(arr, 2) Then
                ' Copy values from array's first dimension
                nItems = ArrayLen(arr, 1)
                ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1)
                For i = NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1
                    arr2(i) = arr(i + LBound(arr, 1) - NORMALIZE_LBOUND, LBound(arr, 2))
                Next
                NormalizeArray = arr2
            Else
                Err.Raise 32000, Description:="Can only normalize a 2-dimensional array if one of " & "the dimensions contains only one element."
            End If
        Case Else
            Err.Raise 32000, Description:="Can only normalize 1- and 2-dimensional arrays."
    End Select
End Function
Public Function GetUniqueItems(arr() As Variant) As Variant()
    ' Returns an array containing each unique item in the given array only once
    If ArrayLen(arr) = 0 Then GetUniqueItems = Array(): Exit Function
    Dim arrSorted() As Variant, i As Long, uniqueItemsList As New VBALib_List
    arrSorted = SortArray(arr): uniqueItemsList.Add arrSorted(LBound(arrSorted))
    For i = LBound(arrSorted) + 1 To UBound(arrSorted)
        If arrSorted(i) <> arrSorted(i - 1) Then uniqueItemsList.Add arrSorted(i)
    Next
    GetUniqueItems = uniqueItemsList.Items
End Function





Private Sub QuickSort(vArray() As Variant, inLow As Long, inHi As Long)
    ' Sorts a section of an array in place
    Dim pivot As Variant, tmpSwap As Variant
    Dim tmpLow As Long, tmpHi As Long
    tmpLow = inLow: tmpHi = inHi
    pivot = vArray((inLow + inHi) \ 2)
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub
Public Function SortArray(arr() As Variant) As Variant()
    ' Returns a sorted copy of the given array
    If ArrayLen(arr) = 0 Then
        SortArray = Array()
    Else
        Dim arr2() As Variant
        arr2 = arr
        SortArrayInPlace arr2
        SortArray = arr2
    End If
End Function
Public Sub SortArrayInPlace(arr() As Variant)
    ' Sorts the given single-dimension array in place
    QuickSort arr, LBound(arr), UBound(arr)
End Sub










Public Function GetItemInArray(Array_src As Variant, idx As Long) As Variant
    GetItemInArray = Array_src(idx)
End Function
Public Function GetItemInStrArray(StrArray_src As String, separator As String, idx As Long) As Variant
    GetItemInStrArray = GetItemInArray(SplitStrIntoArray(StrArray_src, separator), idx)
End Function
Public Function FindItemInArray(Array_src As Variant, item As String) As Long
    Dim i As Long: FindItemInArray = -1                'Find Item in an Array
    For i = LBound(Array_src) To UBound(Array_src)
        If Array_src(i) = item Then FindItemInArray = i: Exit For
    Next i
End Function





Public Function AppendArray(Array_src As Variant, Array_append As Variant) As String
    Dim FailedReason As String, i As Long                  'Append items to an Array
    For i = LBound(Array_append) To UBound(Array_append)
        ReDim Preserve Array_src(LBound(Array_src) To UBound(Array_src) + 1)
        Array_src(UBound(Array_src)) = Array_append(i)
    Next i
Exit_AppendArray:
    AppendArray = FailedReason
    Exit Function
Err_AppendArray:
    FailedReason = Err.Description
    Resume Exit_AppendArray
End Function
Public Sub DeleteArrayItem(arr As Variant, index As Long)
    Dim i As Long          'Delete Item in Array by Index
    For i = index To UBound(arr) - 1: arr(i) = arr(i + 1): Next
    arr(UBound(arr)) = Empty 'VB cvts this to 0 or vbNullString
    ReDim Preserve arr(LBound(arr) To UBound(arr) - 1)
End Sub





