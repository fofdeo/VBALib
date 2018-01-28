Attribute VB_Name = "_VBALib_RegEx"
Option Explicit




'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function RegExM(ByVal Pattern As String, ByVal SourceString As String, Optional ByVal IgnoreCase As Boolean = False, Optional ByVal Glbl As Boolean = True, Optional ByVal Multiline As Boolean = False) As MatchCollection
    Dim re As Object: Set re = CreateObject("vbscript.regexp")
    With re
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .Global = Glbl
        .Multiline = Multiline
    End With
    Set RegExM = re.Execute(SourceString)
End Function
Public Function RegExR(ByVal Pattern As String, ByVal Replacement As String, ByVal SourceString As String, Optional ByVal IgnoreCase As Boolean = False, Optional ByVal Glbl As Boolean = True, Optional ByVal Multiline As Boolean = False) As String
    Dim re As Object: Set re = CreateObject("vbscript.regexp")
    With re
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .Global = Glbl
        .Multiline = Multiline
    End With
    RegExR = re.Replace(SourceString, Replacement)
End Function
Private Sub RegExDemo()
    Dim Matches As MatchCollection, M As Match 'Demonstration
    Dim foo As String, bar As String: foo = "123 abc 456"
    Set Matches = Regex.RegExM("\d+", foo)
    Debug.Print Matches.Count '--> 2
    For Each M In Matches
        Debug.Print M.FirstIndex
        Debug.Print M.Value
    Next M
    Set Matches = RegEx.RegExM("blahblah", foo)
    Debug.Print Matches.Count '--> 0
    For Each M In Matches
        Debug.Print "There are no matches so you'll never see this"
    Next M
    bar = Regex.RegExR("\d+", "#", foo)
    Debug.Print bar '--> # abc #
    Set Matches = Regex.RegExM("\d+", foo, Glbl:=False)
    Debug.Print Matches.Count
End Sub



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Function stripTags(ByVal txt As String) As String
    regMask = "(<.+?>)" 'Strip all Tags in String
    stripTags = matchExpreg(txt, regMask, "")
End Function
Function FileNameFromFullPath(ByVal txt As String) As String
    regMask = ".+\\(.+)"     'Extract FileName from FullPath
    FileNameFromFullPath = findExpreg(txt, regMask)
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Function matchExpreg(ByVal txt As String, ByVal matchPattern As String, ByVal replacePattern As String) As String
    'Matches Pattern in text given txt and applies Replacement Pattern
    Dim RE As Object, REMatches As Object: Dim reg_exp As New RegExp
    With reg_exp
        .pattern = matchPattern: .IgnoreCase = True: .Global = True
    End With
    txt = reg_exp.Replace(txt, replacePattern)
    matchExpreg = txt
End Function
Function findExpreg(ByVal txt As String, ByVal matchPattern As String) As String
    'Returns 1st-Occurrence of RegEx Pattern found in given expression
    On Error GoTo errorHandler: Dim expReg As New RegExp
    With expReg
        .pattern = matchPattern: .IgnoreCase = True: .Global = True
    End With
    Set res = expReg.Execute(txt)
    txt = res(0).submatches(0)
    findExpreg = txt: Exit Function
errorHandler:
    findExpreg = False
End Function


