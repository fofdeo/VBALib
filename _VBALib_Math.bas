Attribute VB_Name = "_VBALib_Math"
Option Compare Database
Option Explicit
Public Const Pi = 3.14159265359




'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function Min(X As Double, Y As Double) As Double
    Min = IIf(X < Y, X, Y)    
End Function
Public Function Max(X As Double, Y As Double) As Double 
    Max = IIf(X > Y, X, Y)    
End Function
Public Function Log10(X As Double) As Double
    Log10 = Log(X) / Log(10#)
End Function
Public Function Ceil(X As Double) As Double
    Ceiling = Int(X) - (X - Int(X) > 0)
End Function




'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function Ceiling(num As Double, Optional significance As Double = 1, Optional places As Integer = 0) As Double
    ' Returns its argument rounded up to the given significance or the given number of decimal places.
    ' @param significance: The significance, or step size, of the function.  For example, a step size of 0.2 will ensure that the number returned is a multiple of 0.2.
    ' @param places: The number of decimal places to keep.
    ValidateFloorCeilingParams significance, places
    Ceiling = Floor(num, significance)
    If num <> Ceiling Then Ceiling = Ceiling + significance
End Function
Public Function Floor(num As Double, Optional significance As Double = 1, Optional places As Integer = 0) As Double
    ' Returns its argument truncated (rounded down) to the given significance or the given number of decimal places.
    ' @param significance: The significance, or step size, of the function.  For example, a step size of 0.2 will ensure that the number returned is a multiple of 0.2.
    ' @param places: The number of decimal places to keep.
    ValidateFloorCeilingParams significance, places
    Floor = Int(num / significance) * significance
End Function
Private Sub ValidateFloorCeilingParams(ByRef significance As Double, ByRef places As Integer)
    If places <> 0 Then
        If significance <> 1 Then
            Err.Raise 32000, Description:="Pass either a number of decimal places or a significance " & "to Floor() or Ceiling(), not both."
        Else
            significance = 10 ^ -places
        End If
    End If
End Sub




