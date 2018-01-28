Attribute VB_Name = "_VBALib_General"
Option Explicit
Public NotShowMsgBox As Boolean





'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Sub WrapIfError()
    Dim c as Range
    For Each c in Selection.Cells
        If c.HasFormula And Not c.HasArray Then c.Formula = "=IFERROR(" & Right(c.Formula, Len(c.Formula)-1) & ",0)"
    Next c
End Sub
Public Sub FormulasToValues()
    Selection.Value = Selection.Value
End Sub
Public Sub TextToNumber()
    Dim c as Range
    If Selection.Count > 1 then
        For Each c in Selection
            If IsNumeric(c) And c <> vbNullString Then c.Value = Val(c.Value)
        Next
    Else
        For Each c in ActiveSheet.UsedRange
            If IsNumeric(c) and c <> vbNullString Then c.Value = Val(c.Value)
        Next
    End If
End Sub
Public Sub TogglePageBreaks()
    ActiveSheet.DisplayPageBreaks = Not ActiveSheet.DisplayPageBreaks
End Sub
Public Sub KillStyles()
    'Removes all hidden styles from Workbook
    Dim styT as Style                    'Get Confirmation
    If MsgBox("There are: " & ActiveWorkbook.Styles.Count - 47 & " custom styles." & vbNewLine & vbNewLine & _ 
        "Delete?", vbInformation + vbYesNo) <> vbYes Then Exit Sub 'Get Confirmation
    Application.StatusBar = "Deleting Styles... Beg: " & Time      'Status Bar Update
    Application.Wait Now + (#12:00:01 AM#)                         '1sec Gap to Break
    For Each styT in ActiveWorkbook.styles
        If Not styT.BuiltIn then styT.Delete
    Next styT   'Clear Status Bar
    Application.StatusBar = False
End Sub




'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function Show_MsgBox(str As String) As Boolean
    'Display string in a msgbox depending on the user-defined flag
    If NotShowMsgBox = False Then MsgBox str
    Show_MsgBox = NotShowMsgBox
End Function
Public Function Enable_MsgBox()
    NotShowMsgBox = False 'Enable user-defined MsgBox
End Function
Public Function Disable_MsgBox()
    NotShowMsgBox = True 'Disable user-defined MsgBox
End Function




'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Sub Excel_On()
    With Application
        .DisplayStatusBar = True
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With ' App.
End Sub
Public Sub Excel_Off()
    With Application
        .DisplayStatusBar = False
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With ' App.
End Sub
Public Sub Excel_Toggle()
    With Application
        .DisplayStatusBar = Not .DisplayStatusBar
        .ScreenUpdating = Not .ScreenUpdating
        .EnableEvents = Not .EnableEvents
        .Calculation = xlCalculationManual
    End With ' App.
End Sub
Public Sub PasswordBreaker()
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    On Error Resume Next
    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If ActiveSheet.ProtectContents = False Then
        MsgBox "One usable password is " & Chr(i) & Chr(j) & _
            Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
            Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
         Exit Sub
    End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
End Sub










