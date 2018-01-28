'====================================================================================================
'##  [VBALib] String Utilities
'----------------------------------------------------------------------------------------------------
' Contents:
'
'
'
'----------------------------------------------------------------------------------------------------
' This program is free software; you can redistribute it and/or modify it under the terms of the
' GNU General Public License as published by the Free Software Foundation; version v2 and any later.
'----------------------------------------------------------------------------------------------------
Attribute VB_Name="_VBALib_String"
Option Explicit







'====================================================================================================
'##  [RegEx]
'----------------------------------------------------------------------------------------------------
Public Function RegExM(ByVal Pattern As String, ByVal Source As String, Optional ByVal IgnoreCase As Boolean=False,Optional ByVal MultiLine As Boolean=False,Optional ByVal Globle As Boolean=False) As String
    '------------------------------------------------------------------------------------------------
    ' [http://www.regular-expressions.info/dotnet.html];[http://www.tmehta.com/regexp/];
    ' MsgBox RegExM("qwerty123456uiops123456", "[a-z][A-Z][0-9][0-9][0-9][0-9]", False)
    '------------------------------------------------------------------------------------------------
    On Error GoTo Exit_RegExM
    Dim REMatches As Object
    Dim RE As Object: Set RE = CreateObject("VBScript.RegExp")
    With RE
        .Pattern = Pattern
        .IgnoreCase = Not (IgnoreCase)
        .Multiline = MultiLine
        .Global = Globle
    End With
    Set REMatches=RE.Execute(Source)
    RegExM=IIf(REMatches.Count > 0, REMatches(0), False)
Exit_RegExM:
    Exit Function
Err_RegExM:
    ShowMsgBox (Err.Description)
    Resum Exit_RegExM
End Function
'----------------------------------------------------------------------------------------------------
Public Function RegExR(ByVal Pattern As String,ByVal Replace As String,ByVal Souce As String,Optional ByVal IgnoreCase As Boolean=False,Optional ByVal MultiLine As Boolean=False,Optional ByVal Globle As Boolean=True) As String
    '------------------------------------------------------------------------------------------------
    ' Parameters:             'Perform a regular expression replacement
    '    Pattern:             'The regular expression pattern to search for
    '    Replace:             'The string or pattern to be used as a replacement
    '    Source:              'The string to search within
    '    IgnoreCase (True):   'Whether to be case-insensitive
    '    Globle (True):       '[True,False]:=[ReplaceAll,Replace(1)]
    '    MultiLine (False):   'Carat (^) && Dollar ($) match @ beg & end of each line, 
    '                          rather than the beggining and end of the entire string.
    ' Returns:
    '    String, with `Pattern` replaced with `Replace`, if found in `Source`.
    '    If `Pattern` is not found, returns `Source` as-is.
    '------------------------------------------------------------------------------------------------
    '    Dim Pattern As String: Pattern="(:\$[A-Z]+\$)([0-9]+)" '2 groups, first keep, 2nd replace
    '    Dim Replace As String: Replace="$1" & 10     'keep $1, replace second with max row
    '    Dim TestVal As String: TestVal="$B$2:$C$4"
    '    Debug.Print RegExR(Pattern, Replace, TestVal, True, False, False)
    '------------------------------------------------------------------------------------------------
    On Error GoTo Exit_RegExR
    Dim re As Object: Set re=CreateObject("VBScript.RegExp")
    With re
        .Pattern=Pattern
        .IgnoreCase=IgnoreCase
        .Multiline=Multiline
        .Global=Glbl
    End With
    RegExR=re.Replace(SourceString, Replacement)
Exit_RegExR:
    Exit Function
Err_RegExR:
    ShowMsgBox (Err.Description)
    Resum Exit_RegExR
End Function
'----------------------------------------------------------------------------------------------------





'====================================================================================================
'##  [SplitText]
'----------------------------------------------------------------------------------------------------
Public Function SplitTrim(s As String, delim As String) As String()
    ' Splits a string on a given delimiter, trimming trailing and leading whitespace from each piece of the string.
    Dim arr() As String
    arr=Split(s, delim)
    Dim i As Integer
    For i=0 To UBound(arr)
        arr(i)=Trim(arr(i))
    Next
    SplitTrim=arr
End Function
'----------------------------------------------------------------------------------------------------
Public Function SplitText(InTextLine As String, Delimeter As String) As Variant
    '---------------------------------------------------------------------------------------------------------
    ' SplitText          - Returns a string array of delimited values; removes extra spaces in splits
    '                    - In : InTextLine As String, Delimeter As String
    '                    - Out: SplitText as String()
    '                    - Last Updated: 3/9/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim k As Long, StringCount As Integer
    Dim TempString() As String
    Dim ThisChar As String, LastChar As String
    
    StringCount=1
    ReDim TempString(1 To StringCount)
    LastChar=Delimeter
    
    For k=1 To Len(InTextLine)
        ThisChar=Mid(InTextLine, k, 1)
        If ThisChar=Delimeter Then
            If LastChar <> Delimeter Then
                StringCount=StringCount + 1
                ReDim Preserve TempString(1 To StringCount)
                LastChar=ThisChar
            End If
        Else
            TempString(StringCount)=TempString(StringCount) & ThisChar
            LastChar=ThisChar
        End If
    Next k
    SplitText=TempString
End Function
Public Function SplitTextReturn(InTextLine As String, Delimeter As String, ReturnID As Integer) As String
    '---------------------------------------------------------------------------------------------------------
    ' SplitTextReturn    - Returns a field of a delimited text string
    '                    - In : InTextLine As String, Delimeter As String, ReturnID as Integer
    '                    - Out: SplitTextReturn as String
    '                    - Last Updated: 9/28/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim SplitString As Variant
    On Error GoTo IsErr
    SplitString=SplitText(InTextLine, Delimeter)
    SplitTextReturn=SplitString(ReturnID)
    Exit Function
IsErr:
    SplitTextReturn=""
End Function
Function SplitTextReturnOne(InputString As String, ReturnValue As Integer) As String
    'Parse text string and return desired word
    Dim StringVariant As Variant
    On Error GoTo IsErr:
    
    StringVariant=SplitText(InputString, " ")
    SplitTextReturnOne=StringVariant(ReturnValue)
    Exit Function
IsErr:
    SplitTextReturnOne="SplitTextError"
End Function
'====================================================================================================
'##  [Text] Misc.
'----------------------------------------------------------------------------------------------------
Public Function ReturnTextBetween(SearchText As String, StartField As String, EndField As String) As String
    '---------------------------------------------------------------------------------------------------------
    ' ReturnTextBetween  - Returns string between starting and ending search strings
    '                    - In : SearchText As String, StartField As String, EndField As String
    '                    - Out: ReturnTextBetween as String
    '                    - Last Updated: 3/9/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim CropLeft As String
    If InStr(1, SearchText, EndField, vbTextCompare)=0 Then
        FindTextBetween="ERROR- End field not found (" & """" & EndField & """" & " not not found in " & """" & SearchText & """" & ")"
        MsgBox FindTextBetween
    ElseIf InStr(1, SearchText, StartField, vbTextCompare)=0 Then
        MsgBox FindTextBetween
        FindTextBetween="ERROR- Start field not found (" & """" & StartField & """" & " not not found in " & """" & SearchText & """" & ")"
    Else
        CropLeft=Left(SearchText, InStr(1, SearchText, EndField, vbTextCompare) - 1)
        ReturnTextBetween=Right(CropLeft, Len(CropLeft) - (InStr(1, SearchText, StartField, vbTextCompare) + Len(StartField) - 1))
    End If
End Function
Public Function IsTextFound(ByVal FindText As String, ByVal WithinText As String) As Boolean
    '----------------------------------------------------------------
    ' IsTextFound       - Returns true if text is found, false if otherwise
    '                   - In : ByVal FindText As String, ByVal WithinText As String
    '                   - Out: Boolean true if found, false if not
    '                   - Last Updated: 4/12/11 by AJS
    '----------------------------------------------------------------
    If InStr(1, WithinText, FindText, vbTextCompare) > 0 Then
        IsTextFound=True
    Else
        IsTextFound=False
    End If
End Function
'----------------------------------------------------------------------------------------------------
Function Printf(ByVal FormatWithPercentSign As String, ParamArray InsertArray()) As String
    'http://www.freevbcode.com/ShowCode.asp?ID=9342
    Dim ResultString As String
    Dim Element As Variant
    Dim FormatLocation As Long

    If IsMissingValue(InsertArray()) Then
        'raise an error
    End If
    
    ResultString=FormatWithPercentSign
    For Each Element In InsertArray
        FormatLocation=InStr(ResultString, "%")
        ResultString=Left$(ResultString, FormatLocation - 1) & Element & Right$(ResultString, Len(ResultString) - FormatLocation - 1)
    Next
    Printf=ResultString
End Function
'----------------------------------------------------------------------------------------------------





'====================================================================================================
'##  [Trim] Chars
'----------------------------------------------------------------------------------------------------
Public Function TrimChars(s As String, toTrim As String)
    ' Trims a specified set of characters from the beginning and end of the given string. @param toTrim: The characters to trim.  For example, if ",; " is given, then all spaces, commas, and semicolons will be removed from the beginning and end of the given string.
    TrimChars=TrimTrailingChars(TrimLeadingChars(s, toTrim), toTrim)
End Function
Public Function TrimLeadingChars(s As String, toTrim As String)
    ' Trims a specified set of characters from the beginning of the given string. @param toTrim: The characters to trim.  For example, if ",; " is given, then all spaces, commas, and semicolons will be removed from the beginning of the given string.
    If s= vbNullString Then
        TrimLeadingChars=vbNullString
        Exit Function
    End If
    Dim i As Integer
    i=1
    While InStr(toTrim, Mid(s, i, 1)) > 0 And i <= Len(s)
        i=i + 1
    Wend
    TrimLeadingChars=Mid(s, i)
End Function
Public Function TrimTrailingChars(s As String, toTrim As String)
    ' Trims a specified set of characters from the end of the given string. @param toTrim: The characters to trim.  For example, if ",; " is given, then all spaces, commas, and semicolons will be removed from the end of the given string.
    If s=vbNullString Then
        TrimTrailingChars=vbNullString
        Exit Function
    End If
    Dim i As Integer
    i=Len(s)
    While InStr(toTrim, Mid(s, i, 1)) > 0 And i >= 1
        i=i - 1
    Wend
    TrimTrailingChars=Left(s, i)
End Function
'====================================================================================================
'##  [Trim] Misc.
'----------------------------------------------------------------------------------------------------
Public Function BegWith(s As String, prefix As String, Optional caseSensitive As Boolean=True) As Boolean
    ' Determines whether a string starts with a given prefix.
    If caseSensitive Then
        BegWith=(Left(s, Len(prefix))=prefix)
    Else
        BegWith=(Left(LCase(s), Len(prefix))=LCase(prefix))
    End If
End Function
Public Function EndWith(s As String, suffix As String, Optional caseSensitive As Boolean=True) As Boolean
    ' Determines whether a string ends with a given suffix.
    If caseSensitive Then
        EndWith=(Right(s, Len(suffix))=suffix)
    Else
        EndWith=(Right(LCase(s), Len(suffix))=LCase(suffix))
    End If
End Function
'----------------------------------------------------------------------------------------------------





'====================================================================================================
'##  [Scrub] Stuff
'----------------------------------------------------------------------------------------------------
Function ScrubFileName(stringToScrub As String) As String
    ' remove illegal characters from filenames
    Dim newString As String
    newString=Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(stringToScrub, "|", ""), ">", ""), "<", ""), Chr(34), ""), "?", ""), "*", ""), ":", ""), "/", ""), "\", "")
    ScrubFileName=newString
End Function
Function ScrubCharReturns(ByVal txt As String) As String
    ' removes character Returns... both types
    txt=Replace(txt, Chr(13), "")       '\n
    txt=Replace(txt, Chr(10), "")       '\r
    ScrubCharReturns=txt
End Function
'----------------------------------------------------------------------------------------------------
Function ScrubSqlString(text As String, Optional ByVal isNumber As Boolean) As String
    ' fixes single-quotes & commas to safely* inject SQL into DBs
    If IsMissing(isNumber) Then isNumber=False
    text=Replace(text, "'", "''")
    text=Replace(text, "€", "&euro;")
    If isNumber Then text=Replace(text, ",", ".")
    ScrubSqlString=text
End Function
'----------------------------------------------------------------------------------------------------





'====================================================================================================
'##  [Add] Chars
'----------------------------------------------------------------------------------------------------
Public Function AddNewLine(Optional Repeat As Integer=1) As String
    '----------------------------------------------------
    ' AddNewLine         - Prints a new line Chr(10), can be repeated
    '                    - In : <none>
    '                    - Out: Chr(10)
    '                    - Last Updated: 3/15/11 by AJS
    '----------------------------------------------------
     AddNewLine=WorksheetFunction.Rept(Chr(10), Repeat)
End Function
Public Function AddQuotes(ByVal TextInQuotes As String) As String
    '----------------------------------------------------
    '  AddQuotes         - Surrounds text in quotations
    '                    - In : TextInQuotes as String
    '                    - Out: "TextInQuotes" as String
    '                    - Last Updated: 3/6/11 by AJS
    '----------------------------------------------------
    AddQuotes=Chr(34) & TextInQuotes & Chr(34)
End Function
Public Function AddTab(Optional Repeat As Integer=1) As String
    '----------------------------------------------------
    '  AddTab            - Adds a tab
    '                    - In : Repeat As Integer
    '                    - Out: Tabs in string
    '                    - Last Updated: 6/17/11 by AJS
    '----------------------------------------------------
    AddTab=WorksheetFunction.Rept(Chr(9), Repeat)
End Function
'====================================================================================================
'##  [Add] Misc.
'----------------------------------------------------------------------------------------------------
Public Function AddToArrayIfUnique(ByVal NewString As String, ArrayName() As Variant) As Variant
    '----------------------------------------------------------------
    '    Dim ArrayName() As Variant
    '    Let ArrayName=[{"Andy", "Cara", "Josh"}]
    '    ArrayName()=AddToArrayIfUnique("Bill", ArrayName())
    '    ArrayName()=AddToArrayIfUnique("Andy", ArrayName())
    '----------------------------------------------------------------
    Dim EachValue As Variant
    Dim Duplicate As Boolean
    Dim ArrEmpty As Boolean
    
    ArrEmpty=False
    On Error Resume Next
    If UBound(ArrayName)=0 Then ArrEmpty=True
    On Error GoTo 0
    
    Duplicate=False
    If ArrEmpty=False Then
        For Each EachValue In ArrayName
            If NewString=EachValue Then
                Duplicate=True
                Exit For
            End If
        Next
    End If
    
    If ArrEmpty=True Then
        ReDim ArrayName(1 To 1)
        ArrayName(1)=NewString
    ElseIf Duplicate=False Then
        ReDim Preserve ArrayName(LBound(ArrayName()) To UBound(ArrayName()) + 1)
        ArrayName(UBound(ArrayName))=NewString
    End If
    Let AddToArrayIfUnique=ArrayName
End Function
Public Function SplitStrIntoArray(str As String, separator As String) As Variant
    'Split a string into array by separator
    Dim Arr As Variant, i as Integer
    If Len(str) > 0 Then
        Arr=Split(str, separator)
        For i=LBound(Arr) To UBound(Arr)
            Arr(i)=Trim(Arr(i))
        Next i
    Else
        Arr=Array()
    End If
    SplitStrIntoArray=Arr
End Function
Public Function FindStrInArray(Array_str As Variant, str As String) As Integer
    'Find string in an array
    FindStrInArray=-1
    Dim i As Integer
    For i=LBound(Array_str) To UBound(Array_str)
        If str=Array_str(i) Then 
            FindStrInArray=i
            Exit For
        End If
    Next i
End Function
'----------------------------------------------------------------------------------------------------
Public Function BuildXMLText(ByVal FieldName As String, ByVal Value As String, Optional NumTabs As Integer=0) As String
    '----------------------------------------------------------------
    ' BuildXMLText          - Builds XML text string
    '                       - In : ByVal FieldName As String, ByVal Value As String, Optional NumTabs As Integer=0
    '                       - Out: XML test string for a single line:   <FieldName>Value</Field>
    '                       - Last Updated: 3/23/11 by AJS
    '----------------------------------------------------------------
    BuildXMLText=AddTab(NumTabs) & "<" & FieldName & ">" & Value & "</" & FieldName & ">"
End Function
'----------------------------------------------------------------------------------------------------







