Attribute VB_Name = "_VBALib_Business"
Option Explicit





Public Function GetUsedRangeOnSheet(Optional thisSheet As Worksheet = Nothing) As Range
    'Returns the used range on a sheet by utilizing the two helper functions below.
    'I've found Excel's builtin "usedRange" function to be unreliable at times. This is a highly usable replacement for data-oriented sheets
    '(such as database dumps) where the data begins in cell A1. Beware that if Row 1 or Column A are empty then this will not behave as you'd expect.
    If thisSheet Is Nothing Then Set thisSheet = ActiveSheet
    With thisSheet
        Set GetUsedRangeOnSheet = .Range(.Cells(1, 1), .Cells(getLastUsedRowOnSheet(thisSheet), getLastUsedColumnOnSheet(thisSheet)))
    End With
End Function
Public Function GetLastUsedRowOnSheet(Optional thisSheet As Worksheet = Nothing) As Long
    'Returns the last used row on a sheet as a long by searching backwards from A1
    If thisSheet Is Nothing Then Set thisSheet = ActiveSheet
    With thisSheet
        GetLastUsedRowOnSheet = .Cells.Find("*", [A1], searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    End With
End Function
Public Function GetLastUsedColumnOnSheet(Optional thisSheet As Worksheet = Nothing) As Long
    'Returns the last used column on a sheet as a long by searching backward from A1
    If thisSheet Is Nothing Then Set thisSheet = ActiveSheet
    With thisSheet
        GetLastUsedColumnOnSheet = .Cells.Find("*", [A1], searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
    End With
End Function





Public Function InString(stringToSearch As String, stringToLookFor As String, Optional startingAt As Long = 1, Optional compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    'Helper function for checking if a string is within another string and returns a boolean.
    InString = IIf(InStr(startingAt, stringToSearch, stringToLookFor, compareMethod) > 0, True, False)
End Function
Public Function SheetExists(ByVal sheetName As String) As Boolean
    'Convenience function for checking if a given sheet exists.
    On Error Resume Next
    SheetExists = (Sheets(sheetName).Name <> vbNullString)
    On Error GoTo 0
End Function
Public Function FolderExists(strPath As String) As Boolean
    'Checks for the existance of a folder referenced by strPath
    FolderExists = IIf(Len(Dir(strPath, vbDirectory))=0, False, True)
End Function
Public Function FileExists(FileName As String) As Boolean
    'Checks for the existance of a file referenced by FileName
    FileExists = (Dir(FileName) > vbNullString)
End Function





Public Function GetQuarter(Optional Offset As Long = 0, Optional ByVal DateSeed As Date = vbNull, Optional ErrorLog As BM_ErrorLog = Nothing) As String
    'Returns Fiscal Quarter as String of format "Q1-13", i.e. Quarter number, then Fiscal Year (FY="YY"). This assumes a Fiscal Year ending in September.
    'Returns for current DateTime, but allows Offset to return Previous/Future Quarters from given DateTime. ErrorLog requires a customer logger from GitHub
    If DateSeed = vbNull Then DateSeed = Now
    DateSeed = DateAdd("q", Offset, DateSeed)    'DateSeed + Offset * 90
    Select Case Month(DateSeed) Mod 12
        Case 10, 11, 12, 0    'Zero case handles 12 mod 12
            GetQuarter = "Q1-" & Right(Year(DateSeed) + 1, 2)
        Case 1, 2, 3
            GetQuarter = "Q2-" & Right(Year(DateSeed), 2)
        Case 4, 5, 6
            GetQuarter = "Q3-" & Right(Year(DateSeed), 2)
        Case 7, 8, 9
            GetQuarter = "Q4-" & Right(Year(DateSeed), 2)
        Case Else
            GetQuarter = "INVALID MONTH"
            If Not ErrorLog Is Nothing Then ErrorLog.logError "WARNING: Tried to get quarter for " & DateSeed & ", but the date is invalid."
    End Select
End Function
Public Function GetFiscalYear(Optional Offset As Long = 0, Optional ByVal DateSeed As Date = vbNull) As String
    'Returns Fiscal Year as String of format "FY13". Can be given an Offset. No Date given assumes "Now".
    If DateSeed = vbNull Then DateSeed = Now
    DateSeed = DateAdd("yyyy", Offset, DateSeed)
    Select Case Month(DateSeed)
        Case 10, 11, 12
            GetFiscalYear = "FY" & Right(Year(DateSeed) + 1, 2)
        Case Else
            GetFiscalYear = "FY" & Right(Year(DateSeed), 2)
    End Select
End Function
Public Function GetLastCompleteQuarter(Optional ByVal DateSeed As Date = vbNull) As Long
    'Returns last completed Quarter as Integer (Long) based on given Date. No Date given assumes "Now".
    'Ex: when a Quarter is complete, show "Actuals", but if the Quarter is incomplete, show "Forecast".
    If DateSeed = vbNull Then DateSeed = Now
    Select Case Month(DateSeed)
        Case 10, 11, 12
            GetLastCompleteQuarter = 0
        Case 1, 2, 3
            GetLastCompleteQuarter = 1
        Case 4, 5, 6
            GetLastCompleteQuarter = 2
        Case 7, 8, 9
            GetLastCompleteQuarter = 3
        Case Else
            GetLastCompleteQuarter = -1
    End Select
End Function
Public Function GetFirstMonthOfQuarter(Optional ByVal DateSeed As Date = vbNull) As Long
    'Returns 1st Month of a Quarter as Long of given Date. No Date given assumes "Now".
    'Useful if you want to iterate of each month of a quarter.
    If DateSeed = vbNull Then DateSeed = Now
    Select Case Month(DateSeed)
        Case 10, 11, 12
            GetFirstMonthOfQuarter = 10
        Case 1, 2, 3
            GetFirstMonthOfQuarter = 1
        Case 4, 5, 6
            GetFirstMonthOfQuarter = 4
        Case 7, 8, 9
            GetFirstMonthOfQuarter = 7
        Case Else
            GetFirstMonthOfQuarter = -1
    End Select
End Function
Public Function GetStartOfYearOffset(Optional DateSeed As Date = vbNull) As Long
    'For handling an Offset between actual and fiscal years in some scenarios.
    If DateSeed = vbNull Then DateSeed = Now
    Select Case Month(DateSeed)
        Case 10, 11, 12
            GetStartOfYearOffset = 0
        Case Else
            GetStartOfYearOffset = -1
    End Select
End Function
Public Function GetActualsOfYearOffset(Optional DateSeed As Date = vbNull) As Long
    'For handling an Offset between actual and fiscal years in some scenarios.
    If DateSeed = vbNull Then DateSeed = Now
    Select Case Month(DateSeed)
        Case 10, 11, 12
            GetActualsOfYearOffset = 1
        Case Else
            GetActualsOfYearOffset = 0
    End Select
End Function





Public Function CreatePivotTableOnSheet(ws As Worksheet, dataSheet As Worksheet, Optional atCell As String = "A1") As PivotTable
    'Convenience function for creating a pivot table from a selected data sheet at a specified location
    Dim pt As PivotTable: Set CreatePivotTableOnSheet = ActiveWorkbook.PivotCaches.Create(xlDatabase, dataSheet.UsedRange).CreatePivotTable(ws.Range(atCell), ws.Name)
End Function
Public Function ReadTextFile(Fname As String, Length As Integer) As Variant
    'Reads Length bytes of content from file Fname, and returns the result as a Variant.
    If FileExists(Fname) Then
        Close #1    
        Open Fname For Input As #1
        ReadTextFile = Input(Length, 1)
        Close 1
    Else
        ReadTextFile = False
    End If
End Function
Public Function StripDateFromSheetName(thisSheet As Worksheet) As String
    'Helper function for striping dates from sheets named in the format "SheetName 2-18" where "2-18" represents a date. For reports generated automatically from code.
    StripDateFromSheetName = IIf(InString(thisSheet.Name, "-"), Strings.Left(thisSheet.Name, Strings.InStr(1, thisSheet.Name, "-") -2), thisSheet.Name)
End Function

Private Function PromptForMultipleTextInputs() As Variant
    'Returns an Array of Paths to .txt files for processing. Allows batch importing. Can use w/ openPipeSeparatedUTF8()
    Dim filter As String: filter = "Text Files (*.txt),*.txt"
    Dim title As String: title = "Select multiple txt files to process..."
    With Application
        PromptForMultipleTextInputs = .GetOpenFilename(filter, 1, title, , True)
    End With
End Function
Private Function Read2DExceptionList(fileName) As Dictionary
    ' Returns a Dictionary-of-Dictionaries given txt Exception List w/ syntax "Header: Value"... Ex: "Division: Healthcare" & vbNewLine & "Division: Embedded"
    ' Useful in processing DataBase Dumps for Exceptions... ToDo: adapt to read in 2-Dimensional Data for other, fancier uses.
    Dim exceptionsDict As Scripting.Dictionary, tempDict As Scripting.Dictionary, fHandle As Integer, fLine As String, errorLog As String
    Dim strToKeep As String, delim As String, pos As Integer, resp As Integer, headerMatch As String, valueMatch As String, LineNum As Long
    On Error Resume Next: delim = ":": LineNum = 0
    Set exceptionsDict = New Scripting.Dictionary
    errorLog = vbNullString: fHandle = FreeFile()
    Open fileName For Input As fHandle
    Do While (Not (EOF(fHandle)))
        Line Input #fHandle, fLine
        LineNum = LineNum + 1: fLine = Trim(fLine)
        If fLine <> vbNullString And Strings.Left(fLine, 1) <> "'" And Strings.Left(fLine, 1) <> "#" Then    'comments delimited by ' or #
            pos = InStr(1, fLine, "'")
            If pos = 0 Then pos = InStr(1, fLine, "#")
            strToKeep = IIf(pos = 0, fLine, Trim(Left(fLine, pos - 1)))
            pos = Strings.InStr(1, strToKeep, delim) 'split line into header and value:
            If pos = 0 Then
                errorLog = errorLog & Chr(9) & "Missing ':' separator in line " & LineNum & ":   '" & strToKeep & "'" & Chr(13)
            ElseIf pos = 1 Then
                errorLog = errorLog & Chr(9) & "Column header empty in line " & LineNum & ":   '" & strToKeep & "'" & Chr(13)
            Else
                headerMatch = Strings.Trim(Strings.Left(strToKeep, pos - 1))
                valueMatch = Strings.Trim(Strings.Mid(strToKeep, pos + 1))
                If Not exceptionsDict.exists(headerMatch) Then
                    Set tempDict = New Scripting.Dictionary
                    tempDict.Add Key:=valueMatch, Item:=headerMatch
                    exceptionsDict.Add Key:=headerMatch, Item:=tempDict
                    Set tempDict = Nothing
                Else
                    exceptionsDict(headerMatch).Add Key:=valueMatch, Item:=headerMatch
                End If
            End If
        End If
    Loop
    Close fHandle
    If errorLog <> vbNullString Then
        resp = MsgBox("Errors were found in " & fileName & ":" & Chr(13) & Chr(13) & errorLog & Chr(13) & "Continue anyway?", vbCritical + vbYesNo + vbDefaultButton2, "Error(s) in exception list!")
        If resp = vbNo Then Exit Function
    End If
    Set Read2DExceptionList = exceptionsDict
End Function
Private Function OpenPipeSeparatedUTF8() As Workbook
    'Opens a pipe-separated text file, enforcing UTF8 encoding and US English number separators
    'Returns workbook object representing processed pipe-separated file
    On Error Resume Next: Dim fn As String
    fn = Excel.Application.GetOpenFilename(fileFilter:="Text Files (*.txt), *.txt,All Files (*.*),*.*", title:="Open Pipe-Separated Report...")
    If fn <> "False" Then
        Excel.Workbooks.OpenText fileName:=fn, Origin:=msoEncodingUTF8, DataType:=xlDelimited, TextQualifier:=xlTextQualifierNone, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, other:=True, OtherChar:="|", DecimalSeparator:=".", ThousandsSeparator:=","
        Set OpenPipeSeparatedUTF8 = Excel.ActiveWorkbook
    End If
End Function





Private Sub ReSaveWorkbookAsXlsx(thisWorkbook As Workbook)
    'Re-Saves an opened .txt (or any 3-digit extension) WorkBook as .xlsx
    On Error Resume Next: Dim fileName As String
    fileName = Left(thisWorkbook.FullName, Len(thisWorkbook.FullName) - 3) & "xlsx"
    thisWorkbook.SaveAs fileName, xlOpenXMLWorkbook
    On Error GoTo 0
End Sub
Public Sub FreezePanesOnSheet(sheetToFreeze As Worksheet, atPosition As Range)
    'Helper function to freeze panes on a sheet. Leaves handling of screen updating to the user
    Dim currentlyActiveSheet As Worksheet: Set currentlyActiveSheet = ActiveSheet
    sheetToFreeze.Activate
    Application.GoTo sheetToFreeze.Range("A1"), True
    sheetToFreeze.Range(atPosition.Address).Select
    ActiveWindow.FreezePanes = True
    currentlyActiveSheet.Activate 'reactivate whatever sheet was previously active
End Sub
Public Sub AddConditionalFormattingForUndefinedOnSheet(sheetToFormat As Worksheet)
    ' Will colorize all cells in the range that contain the text "Undefined", useful in select scenarios.
    With sheetToFormat.UsedRange
        .FormatConditions.Add xlCellValue, xlEqual, "=""Undefined"""
        .FormatConditions(1).Font.ThemeColor = xlThemeColorAccent2
        .FormatConditions(1).Font.TintAndShade = -0.249946592608417
    End With
End Sub
Public Sub InitializeColumnHeadersFor(sheetToInitialize As Worksheet, outputDictionary As Dictionary, Optional ByVal headerRow As Long = 1)
    'Parses Header Row to a Dictionary Object for easy access. Any repeated Headings will have count appended (Ex. [Name|Name]>[Name|Name2]
    'ToDo: re-write Sub to Function for clearer semantics... however passing the outputDictionary to the Sub initializes it for the user...    
    Dim lastDataColumn As Long, currentColumn As Long
    Dim currentKey As String, numberOfRepeats As Long
    Set outputDictionary = New Scripting.Dictionary
    lastDataColumn = sheetToInitialize.UsedRange.Columns.Count
    For currentColumn = 1 To lastDataColumn
        currentKey = Trim(sheetToInitialize.Cells(headerRow, currentColumn).Value)
        currentKey = Trim(currentKey): numberOfRepeats = 1
        Do While outputDictionary.exists(currentKey)
            numberOfRepeats = numberOfRepeats + 1
            currentKey = currentKey & " " & numberOfRepeats
        Loop
        If currentKey <> vbNullString Then outputDictionary.Add Key:=currentKey, Item:=currentColumn
    Next currentColumn
End Sub




