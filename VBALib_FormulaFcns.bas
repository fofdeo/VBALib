Attribute VB_Name = "VBALib_FormulaFunctions"
' Common VBA Library - FormulaFunctions - Provides functions that are useful in Excel formulas.

Option Explicit

Public Function ArrayElement(arr As Variant, i1 As Variant, _
    ' Retrieves the given element of an array.
    Optional i2 As Variant, Optional i3 As Variant, _
    Optional i4 As Variant, Optional i5 As Variant) As Variant
    
    If IsMissing(i2) Then
        If IsObject(arr(i1)) Then
            Set ArrayElement = arr(i1)
        Else
            ArrayElement = arr(i1)
        End If
    ElseIf IsMissing(i3) Then
        If IsObject(arr(i1, i2)) Then
            Set ArrayElement = arr(i1, i2)
        Else
            ArrayElement = arr(i1, i2)
        End If
    ElseIf IsMissing(i4) Then
        If IsObject(arr(i1, i2, i3)) Then
            Set ArrayElement = arr(i1, i2, i3)
        Else
            ArrayElement = arr(i1, i2, i3)
        End If
    ElseIf IsMissing(i5) Then
        If IsObject(arr(i1, i2, i3, i4)) Then
            Set ArrayElement = arr(i1, i2, i3, i4)
        Else
            ArrayElement = arr(i1, i2, i3, i4)
        End If
    Else
        If IsObject(arr(i1, i2, i3, i4, i5)) Then
            Set ArrayElement = arr(i1, i2, i3, i4, i5)
        Else
            ArrayElement = arr(i1, i2, i3, i4, i5)
        End If
    End If
End Function
Public Function StringSplit(s As String, delim As String, Optional limit As Long = -1) As String()
    ' Splits a string into an array, optionally limiting the number of items in the returned array.
    StringSplit = Split(s, delim, limit)
End Function
Public Function StringJoin(arr() As Variant, delim As String) As String
    ' Joins an array into a string by inserting the given delimiter in between items.
    StringJoin = Join(arr, delim)
End Function
Public Function NewLine() As String
    ' Returns a newline (vbLf) character for use in formulas.
    NewLine = vbLf
End Function
Public Function GetArrayForFormula(arr As Variant) As Variant
    ' Returns an array suitable for using in an array formula.  When this function is called from an array formula, it will detect whether or not the array should be transposed to fit into the range.
    If IsObject(Application.Caller) Then
        Dim len1 As Long, len2 As Long
        Select Case ArrayRank(arr)
            Case 0
                GetArrayForFormula = Empty
                Exit Function
            Case 1
                len1 = ArrayLen(arr)
                len2 = 1
            Case 2
                len1 = ArrayLen(arr)
                len2 = ArrayLen(arr, 2)
            Case Else
                Err.Raise 32000, Description:="Invalid number of dimensions (" & ArrayRank(arr) & "; expected 1 or 2)."
        End Select
        
        If Application.Caller.Rows.Count > Application.Caller.Columns.Count And len1 > len2 Then
            GetArrayForFormula = WorksheetFunction.Transpose(arr)
        Else
            GetArrayForFormula = arr
        End If
    Else
        GetArrayForFormula = arr
    End If
End Function
Public Function RangeToArray(r As Range) As Variant()
    ' Converts a range to a normalized array.
    If r.Cells.Count = 1 Then
        RangeToArray = Array(r.Value)
    ElseIf r.Rows.Count = 1 Or r.Columns.Count = 1 Then
        RangeToArray = NormalizeArray(r.Value)
    Else
        RangeToArray = r.Value
    End If
End Function
Public Function ColumnWidth(Optional c As Integer = 0) As Variant
    ' Returns the width of a column on a sheet.  If the column number is not given and this function is used in a formula, then it returns the column width of the cell containing the formula.
    Application.Volatile: Dim s As Worksheet
    Set s = IIf(IsObject(Application.Caller), Application.Caller.Worksheet, ActiveSheet)
    If c <= 0 And IsObject(Application.Caller) Then c = Application.Caller.Column
    ColumnWidth = s.Columns(c).Width
End Function
Public Function RowHeight(Optional r As Integer = 0) As Variant
    ' Returns the height of a row on a sheet.  If the row number is not given and this function is used in a formula, then it returns the row height of the cell containing the formula.
    Application.Volatile: Dim s As Worksheet
    Set s = IIf(IsObject(Application.Caller), Application.Caller.Worksheet, ActiveSheet)
    If r <= 0 And IsObject(Application.Caller) Then r = Application.Caller.Row
    RowHeight = s.Rows(r).Height
End Function
Public Function GetFormula(r As Range, Optional r1c1 As Boolean = False) As Variant
    ' Returns the formula of the given cell or range, optionally in R1C1 style.
    GetFormula = IIf(r1c1, r.FormulaR1C1, r.Formula)
End Function
