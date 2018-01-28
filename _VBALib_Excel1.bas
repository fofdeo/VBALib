Attribute VB_Name = "_VBALib_Excel1"
Option Explicit
Private Declare Function CallNamedPipe Lib "kernel32" Alias "CallNamedPipeA" (ByVal lpNamedPipeName As String, ByVal lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesRead As Long, ByVal nTimeOut As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Enum Corner
    cnrTopLeft
    cnrTopRight
    cnrBottomLeft
    cnrBottomRight
End Enum
Public Enum OverwriteAction
    oaPrompt = 1
    oaOverwrite = 2
    oaSkip = 3
    oaError = 4
    oaCreateDirectory = 8
End Enum




'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function IsWorkbookOpen(wbFilename As String) As Boolean
    ' Determines whether a given workbook has been opened.  Pass this function a filename only, not a full path.
    Dim w As Workbook
    On Error GoTo notOpen
    Set w = Workbooks(wbFilename)
    IsWorkbookOpen = True
    Exit Function
notOpen:
    IsWorkbookOpen = False
End Function
Public Function ChartExists(chartName As String, Optional sheetName As String = vbNullString, Optional wb As Workbook) As Boolean
    ' Determines whether a chart with the given name exists.
    ' @param chartName: The name of the chart to check for.
    ' @param sheetName: The name of the worksheet that contains the given chart (optional; the default is to search all worksheets).
    ' @param wb: The workbook to check for the given chart name (defaults to the active workbook.
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim s As Worksheet, c As ChartObject: ChartExists = False
    If sheetName = vbNullString Then
        For Each s In wb.Sheets
            If ChartExists(chartName, s.Name, wb) Then: ChartExists = True: Exit Function
        Next
    Else
        Set s = wb.Sheets(sheetName)
        On Error GoTo notFound
        Set c = s.ChartObjects(chartName)
        ChartExists = True
notFound:
    End If
End Function
Public Function SheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    ' Determines whether a sheet with the given name exists.
    ' @param wb: The workbook to check for the given sheet name (defaults to the active workbook).    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim s As Object ' Not Worksheet because it could also be a chart
    On Error GoTo notFound
    Set s = wb.Sheets(sheetName)
    SheetExists = True
    Exit Function
notFound:
    SheetExists = False
End Function
Public Function WorkSheetExist(oWb As Workbook, SheetName As String) As Boolean
    'Check if WorkSheet exists in given WorkBook          'ToDo: cvt to 1-Liner
    Dim WS As Worksheet: WorkSheetExist = False
    For Each ws In oWb.Worksheets
        If SheetName = ws.Name Then 
            WorkSheetExist = True
            Exit For
        End If
    Next WS
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Sub DeleteSheetByName(sheetName As String, Optional wb As Workbook)
    ' Deletes the sheet with the given name, without prompting for confirmation.
    ' @param wb: The workbook to check for the given sheet name (defaults to the active workbook).
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If SheetExists(sheetName, wb) Then DeleteSheetOrSheets wb.Sheets(sheetName)
End Sub
Public Sub DeleteSheet(s As Worksheet)
    DeleteSheetOrSheets s 're-Direct
End Sub
Public Sub DeleteSheets(s As Sheets)
    DeleteSheetOrSheets s 're-Direct
End Sub
Private Sub DeleteSheetOrSheets(s As Object)
    Dim prevDisplayAlerts As Boolean
    prevDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    s.Delete
    On Error GoTo 0
    Application.DisplayAlerts = prevDisplayAlerts
End Sub



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function CellReference(ByVal r As Long, ByVal c As Integer, Optional sheet As String = "", Optional absoluteRow As Boolean = False, Optional absoluteCol As Boolean = False) As String
    ' Builds an Excel cell reference.
    Dim ref As String: ref = IIf(absoluteCol, "$", "") & ExcelCol(c) & IIf(absoluteRow, "$", "") & r
    CellReference = IIf(sheet = vbNullString, ref, "'" & Replace(sheet, "'", "''") & "'!" & ref)
End Function
Public Function GetRealUsedRange(s As Worksheet, Optional fromTopLeft As Boolean = True) As Range
    ' Returns the actual used range from a sheet. @param fromTopLeft: If True, returns the used range starting from cell A1, which is different from the way Excel's UsedRange property behaves if the sheet does not use any cells in the top row(s) and/or leftmost column(s).
    Set GetRealUsedRange = IIf(fromTopLeft, s.Range(s.Cells(1,1), s.Cells(s.UsedRange.Rows.Count + s.UsedRange.Row - 1, s.UsedRange.Columns.Count + s.UsedRange.Column - 1)), s.UsedRange)
    'If fromTopLeft Then
    '    Set GetRealUsedRange = s.Range(s.Cells(1, 1), s.Cells( _
    '            s.UsedRange.Rows.Count + s.UsedRange.Row - 1, _
    '            s.UsedRange.Columns.Count + s.UsedRange.Column - 1))
    'Else
    '    Set GetRealUsedRange = s.UsedRange
    'End If
End Function
Public Function SetValueIfNeeded(rng As Range, val As Variant) As Boolean
    ' Sets the value of the given range if it is different than the proposed value. Returns whether the value of the range was changed.
    If rng.Value = val Then
        SetValueIfNeeded = False
    Else
        rng.Value = val
        SetValueIfNeeded = True
    End If
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function ExcelCol(c As Integer) As String
    ' Converts an integer column number to an Excel column string.
    ExcelCol = ExcelCol_ZeroBased(c - 1)
End Function
Public Function ExcelColNum(c As String) As Integer
    ' Converts an Excel column string to an integer column number.
    Dim i As Integer: ExcelColNum = 0
    For i = 1 To Len(c)
        ExcelColNum = (ExcelColNum + Asc(Mid(c, i, 1)) - 64)
        If i < Len(c) Then ExcelColNum = ExcelColNum * 26
    Next
End Function
Private Function ExcelCol_ZeroBased(c As Integer) As String
    Dim c2 As Integer: c2 = c \ 26
    ExcelCol_ZeroBased = IIf(c2 = 0, Chr(65 + c), ExcelCol(c2) & Chr(65 + (c Mod 26)))
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function ExcelErrorType(e As Variant) As String
    ' Returns a string describing the type of an Excel error value ("#DIV/0!", "#N/A", etc.)
    If IsError(e) Then
        Select Case e
            Case CVErr(xlErrDiv0)
                ExcelErrorType = "#DIV/0!"
            Case CVErr(xlErrNA)
                ExcelErrorType = "#N/A"
            Case CVErr(xlErrName)
                ExcelErrorType = "#NAME?"
            Case CVErr(xlErrNull)
                ExcelErrorType = "#NULL!"
            Case CVErr(xlErrNum)
                ExcelErrorType = "#NUM!"
            Case CVErr(xlErrRef)
                ExcelErrorType = "#REF!"
            Case CVErr(xlErrValue)
                ExcelErrorType = "#VALUE!"
            Case Else
                ExcelErrorType = "#UNKNOWN_ERROR"
        End Select
    Else
        ExcelErrorType = "(not an error)"
    End If
End Function
Public Sub ShowStatusMessage(statusMessage As String)
    ' Shows a status message to update the user on the progress of a long-running operation, in a way that can be detected by external applications.
    Application.StatusBar = statusMessage    ' Show the message in the status bar. Set the Excel window title to the updated status message.
    ' To allow external applications to extract just the status message, put the length of the message at the beginning.
    Application.Caption = Len(statusMessage) & ":" & statusMessage ' Window(API) Title: "Status Message - WorkbookFilename.xlsm"
End Sub
Public Sub FlashStatusMessage(statusMessage As String)
    ' Shows a status message for 2-3 seconds then removes it.
    ShowStatusMessage statusMessage
    Application.OnTime Now + TimeValue("0:00:02"), "ClearStatusMessage"
End Sub
Public Sub ClearStatusMessage()
    ' Clears any status message that is currently being displayed by a macro.
    Application.StatusBar = False
    Application.Caption = Empty
End Sub
Public Sub SendMessageToListener(msg As String)
    ' Attempts to send a message to an external program that is running this macro and listening for messages.
    Dim bArray(0 To 0) As Byte, bytesRead As Long
    CallNamedPipe "\\.\pipe\ExcelMacroCommunicationListener." & GetCurrentProcessId, msg, Len(msg), bArray(0), 1, bytesRead, 500
End Sub



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function GetExcelTable(tblName As String, Optional wb As Workbook) As VBALib_ExcelTable
    ' Returns a wrapper object for the table with the given name in the given workbook.
    ' @param wb: The workbook that contains the table (defaults to the active workbook).
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error GoTo notFound: wb.Activate
    Dim wbPrevActive As Workbook: Set wbPrevActive = ActiveWorkbook

    ' We could just do Range(tblName).ListObject, but then this would allow getting a table by any of its cells or columns. 
    ' Instead, do some verification that the string we were passed is actually the name of a table.
    Dim tbl As ListObject: Set tbl = Range(tblName).Parent.ListObjects(tblName)
    wbPrevActive.Activate: Set GetExcelTable = New VBALib_ExcelTable
    GetExcelTable.Initialize tbl
    Exit Function
notFound:
    On Error GoTo 0
    Err.Raise 32000, Description:="Could not find table '" & tblName & "'."
End Function
Public Function GetExcelLink(linkFilename As String, Optional wb As Workbook) As VBALib_ExcelLink
    ' Returns an object representing the link to the Excel workbook with the given filename.
    ' @param linkFilename: The path or filename of the linked Excel workbook.
    ' @param wb: The workbook that contains the link (defaults to the active workbook).
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim matchingLinkName As String: matchingLinkName = GetMatchingLinkName(linkFilename, wb)
    If matchingLinkName = vbNullString Then
        Err.Raise 32000, Description:="No Excel link exists with the given name ('" & linkFilename & "')."
    Else
        Set GetExcelLink = New VBALib_ExcelLink
        GetExcelLink.Initialize wb, matchingLinkName
    End If
End Function
Public Function GetWorkbookFileFormat(fileExtension As String) As XlFileFormat
    ' Returns the Excel workbook format for the given file extension.
    Select Case LCase(Replace(fileExtension, ".", ""))
        Case "xls"
            GetWorkbookFileFormat = xlExcel8
        Case "xla"
            GetWorkbookFileFormat = xlAddIn8
        Case "xlt"
            GetWorkbookFileFormat = xlTemplate8
        Case "csv"
            GetWorkbookFileFormat = xlCSV
        Case "txt"
            GetWorkbookFileFormat = xlCurrentPlatformText
        Case "xlsx"
            GetWorkbookFileFormat = xlOpenXMLWorkbook
        Case "xlsm"
            GetWorkbookFileFormat = xlOpenXMLWorkbookMacroEnabled
        Case "xlsb"
            GetWorkbookFileFormat = xlExcel12
        Case "xlam"
            GetWorkbookFileFormat = xlOpenXMLAddIn
        Case "xltx"
            GetWorkbookFileFormat = xlOpenXMLTemplate
        Case "xltm"
            GetWorkbookFileFormat = xlOpenXMLTemplateMacroEnabled
        Case Else
            Err.Raise 32000, Description:="Unrecognized Excel file extension: '" & fileExtension & "'"
    End Select
End Function
Public Function SaveWorkbookAs(wb As Workbook, newFilename As String, Optional oAction As OverwriteAction = oaPrompt, Optional openReadOnly As Boolean = False) As Boolean
    ' Saves the given workbook as a different filename, with options for handling the case where the file already exists. Returns True if the workbook was saved, or False if it was not saved.
    ' @param oAction: The action that will be taken if the given file exists. This parameter also accepts the oaCreateDirectory flag, which means that the directory hierarchy of the requested filename will be created if it does not already exist. If not given, defaults to oaPrompt.
    If Not FolderExists(GetDirectoryName(newFilename)) Then
        If oAction And oaCreateDirectory Then
            MkDirRecursive GetDirectoryName(newFilename)
        Else
            Err.Raise 32000, Description:="The parent folder of the requested workbook filename " & "does not exist:" & vbLf & vbLf & newFilename
        End If
    End If
    If FileExists(newFilename) Then
        If oAction And oaOverwrite Then
            Kill newFilename 
        ' Proceed to save the file
        ElseIf oAction And oaError Then
            Err.Raise 32000, Description:= "The given filename already exists:" & vbLf & vbLf & newFilename  
        ElseIf oAction And oaPrompt Then
            Dim r As VbMsgBoxResult
            r = MsgBox(Title:="Overwrite Excel file?", Buttons:=vbYesNo + vbExclamation, Prompt:="The following Excel file already exists:" & vbLf & vbLf & newFilename & vbLf & vbLf & "Do you want to overwrite it?")
            If r = vbYes Then
                Kill newFilename
            ' Proceed to save the file
            Else
                SaveWorkbookAs = False
                Exit Function
            End If
        ElseIf oAction And oaSkip Then
            SaveWorkbookAs = False
            Exit Function
        Else
            Err.Raise 32000, Description:="Bad overwrite action value passed to SaveWorkbookAs."
        End If
    End If
    ' wb.SaveCopyAs doesn't take all the fancy arguments that wb.SaveAs does, but it's the only way to save a copy of the current workbook.
    ' This means, among other things, that it is not possible to save the workbook as a different format than the original workbook.
    ' To work around this, call SaveCopyAs with a temporary filename first, then open the temporary file, then call SaveAs with the desired filename and options.
    Dim wbTmp As Workbook, tmpFilename As String: tmpFilename = CombinePaths(GetTempPath, Int(Rnd * 1000000) & "-" & wb.Name)
    wb.SaveCopyAs tmpFilename: Set wbTmp = Workbooks.Open(tmpFilename, UpdateLinks:=False, ReadOnly:=True)
    wbTmp.SaveAs filename:=newFilename, FileFormat:=GetWorkbookFileFormat(GetFileExtension(newFilename)), ReadOnlyRecommended:=openReadOnly
    wbTmp.Close SaveChanges:=False: Kill tmpFilename: SaveWorkbookAs = True
End Function
Public Function SaveActiveWorkbookAs(newFilename As String, Optional oAction As OverwriteAction = oaPrompt, Optional openReadOnly As Boolean = False) As Boolean
    ' Saves the current workbook as a different filename, with options for handling the case where the file already exists. Returns True if the workbook was saved, or False if it was not saved.
    ' @param oAction: The action that will be taken if the given file exists.  This parameter also accepts the oaCreateDirectory flag, which means that the directory hierarchy of the requested filename will be created if it does not already exist.  If not given, defaults to oaPrompt.
    ' @param openReadOnly: True or False to determine whether the created workbook will prompt users to open it as read-only (defaults to False).
    SaveActiveWorkbookAs = SaveWorkbookAs(ActiveWorkbook, newFilename, oAction, openReadOnly)
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function GetCornerCell(r As Range, c As Corner) As Range
    ' Returns the cell in the given corner of the given range.
    Select Case c
        Case cnrTopLeft
            Set GetCornerCell = r.Cells(1, 1)
        Case cnrTopRight
            Set GetCornerCell = r.Cells(1, r.Columns.Count)
        Case cnrBottomLeft
            Set GetCornerCell = r.Cells(r.Rows.Count, 1)
        Case cnrBottomRight
            Set GetCornerCell = r.Cells(r.Rows.Count, r.Columns.Count)
    End Select
End Function
Public Function GetAllExcelLinks(Optional wb As Workbook) As Variant
    ' Returns an array of objects representing the other Excel workbooks that the given workbook links to.
    ' @param wb: The source workbook (defaults to the active workbook).
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim i as Long, linkNames() As Variant, linksArr() As VBALib_ExcelLink
    linkNames = NormalizeArray(ActiveWorkbook.LinkSources(xlExcelLinks))
    If ArrayLen(linkNames) Then
        ReDim linksArr(1 To ArrayLen(linkNames))
        For i = 1 To UBound(linkNames)
            Set linksArr(i) = New VBALib_ExcelLink
            linksArr(i).Initialize wb, CStr(linkNames(i))
        Next
        GetAllExcelLinks = linksArr
    Else
        GetAllExcelLinks = Array()
        Exit Function
    End If
End Function
Private Function GetMatchingLinkName(linkFilename As String, Optional wb As Workbook) As String
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim i as Long, linkNames() As Variant, matchingLinkName As String
    linkNames = NormalizeArray(wb.LinkSources(xlExcelLinks))
    For i = 1 To UBound(linkNames) '1st look for Link w/ exact FullPath given by linkFilename
        If LCase(linkNames(i)) = LCase(linkFilename) Then
            GetMatchingLinkName = linkNames(i): Exit Function
        End If
    Next    ' Next look for a link with the same filename as linkFilename.
    For i = 1 To UBound(linkNames) '2-steps b/c its possible for Excel to link 2x WBs w/ same name in diff folders
        If LCase(GetFilename(linkNames(i))) = LCase(GetFilename(linkFilename)) Then
            GetMatchingLinkName = linkNames(i): Exit Function
        End If
    Next
    GetMatchingLinkName = vbNullString
End Function
Public Function ExcelLinkExists(linkFilename As String, Optional wb As Workbook) As Boolean
    ' Returns whether an Excel link matching the given workbook filename exists. @param wb: The workbook that contains the link (defaults to the active workbook).
    ExcelLinkExists = (GetMatchingLinkName(linkFilename, wb) <> "")
End Function







'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Sub RefreshAccessConnections(Optional wb As Workbook)
    ' Refreshes all Access database connections in the given workbook.
    ' @param wb: The workbook to refresh (defaults to the active workbook).
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim cn As WorkbookConnection, i as Long, numConnnections as Long
    On Error GoTo err_
    Application.Calculation = xlCalculationManual
    For Each cn In wb.Connections
        If cn.Type = xlConnectionTypeOLEDB Then numConnections = numConnections + 1
    Next
    For Each cn In wb.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            i = i + 1
            ShowStatusMessage "Refreshing data connection '" & cn.OLEDBConnection.CommandText & "' (" & i & " of " & numConnections & ")"
            cn.OLEDBConnection.BackgroundQuery = False: cn.Refresh
       End If
    Next
    GoTo done_
err_:
    MsgBox "Error " & Err.Number & ": " & Err.Description
done_:
    ShowStatusMessage "Recalculating"
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ClearStatusMessage
End Sub
Public Function AppendSqlObjToAndActivateWs(oWs As Worksheet, SqlObj_name As String, Optional AddBorder As Boolean = False) As String
    'Append Access SQL Object(Table, Query) to Excel worksheet, and activate it 
    On Error GoTo Err_AppendSqlObjToAndActivateWs: Dim FailedReason As String
    If TableExist(SqlObj_name) = False And QueryExist(SqlObj_name) Then
        FailedReason = SqlObj_name & " does not exist!"
        GoTo Exit_AppendSqlObjToAndActivateWs
    End If
    With oWs 'Have to activate the worksheet for copying query with no error!
        .Activate  'Store the new start row >> append SqlObj_Name to sheet
        Dim RowEnd_old As Long: RowEnd_old = .UsedRange.Rows.Count
        Dim RowBeg_new As Long: RowBeg_new = RowEnd_old + 1          'Create Recordset Object
        Dim rs As DAO.Recordset: Set rs = CurrentDb.OpenRecordset(SqlObj_name, dbOpenSnapshot)
        .Range("A" & CStr(.UsedRange.Rows.count + 1)).CopyFromRecordset rs, 65534
        .Range(.Cells(RowEnd_old, 1), .Cells(RowEnd_old, .UsedRange.Columns.count)).Copy 'Copy format from prev  rows to new rows
        .Range(.Cells(RowBeg_new, 1), .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count)).PasteSpecial Paste:=xlPasteFormats
        If AddBorder = True Then                 'Add border at the last row
            With .UsedRange.Rows(.UsedRange.Rows.count).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlAutomatic
            End With '.UsedRange.Rows(.UsedRange.Rows.count).Borders(xlEdgeBottom)
        End If
    End With 'oWs
Exit_AppendSqlObjToAndActivateWs:
    AppendSqlObjToAndActivateWs = FailedReason
    Exit Function
Err_AppendSqlObjToAndActivateWs:
    FailedReason = Err.Description
    Resume Exit_AppendSqlObjToAndActivateWs
End Function
Public Function ReplaceStrInWsRng(oWsRng As Range, What As Variant, Replacement As Variant, Optional LookAt As Variant, Optional SearchOrder As Variant, Optional MatchCase As Variant, Optional MatchByte As Variant, Optional SearchFormat As Variant, Optional ReplaceFormat As Variant) As String
    'Replace String in a range of a worksheet that enclose any excel error in a function
    On Error GoTo Err_ReplaceStrInWsRng: Dim FailedReason As String
    With oWsRng
        .Application.DisplayAlerts = False
        .Replace What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat, ReplaceFormat
        .Application.DisplayAlerts = True
    End With '.oWsRng
Exit_ReplaceStrInWsRng:
    ReplaceStrInWsRng = FailedReason
    Exit Function
Err_ReplaceStrInWsRng:
    FailedReason = Err.Description
    Resume Exit_ReplaceStrInWsRng
End Function
Public Function LinkToWorksheetInWorkbook(Wb_path As String, ByVal SheetNameList As Variant, Optional ByVal SheetNameLocalList As Variant, Optional ByVal ShtSeriesList As Variant, Optional HasFieldNames As Boolean = True) As String
    On Error GoTo Err_LinkToWorksheetInWorkbook    'Link multiple worksheets in workbooks
    Dim FullNameList() As Variant, SheetNameAndRangeList() As Variant, FailedReason As String
    Dim oExcel As Excel.Application: Set oExcel = CreateObject("Excel.Application")
    If Len(Dir(Wb_path)) = 0 Then: FailedReason = Wb_path: GoTo Exit_LinkToWorksheetInWorkbook
    If VarType(SheetNameLocalList) <> vbArray + vbVariant Then SheetNameLocalList = SheetNameList   
    If UBound(SheetNameList) <> UBound(SheetNameLocalList) Then 'Prepare worksheets to be linked.
        FailedReason = "No. of elements in SheetNameList and SheetNameLocalList are not equal"
        GoTo Exit_LinkToWorksheetInWorkbook
    End If
    With oExcel
        Dim oWb As Workbook: Set oWb = .Workbooks.Open(Filename:=Wb_path, ReadOnly:=True)
        With oWb            'Prepare to link worksheets in series
            If VarType(ShtSeriesList) = vbArray + vbVariant Then
                Dim ShtSeries As Variant, ShtSeries_name As String, ShtSeries_local_name As String
                Dim ShtSeries_beg_idx As Long, ShtSeries_end_idx As Long, WsInS_idx As Long, WsInS_cnt As Long
                For Each ShtSeries In ShtSeriesList
                    ShtSeries_name = ShtSeries(0): ShtSeries_local_name = ShtSeries(1)
                    ShtSeries_beg_idx = ShtSeries(2): ShtSeries_end_idx = ShtSeries(3)
                    If ShtSeries_local_name = vbNullString Then ShtSeries_local_name = ShtSeries_name
                    If ShtSeries_end_idx < ShtSeries_beg_idx Then ShtSeries_end_idx = .Worksheets.count - 1
                    WsInS_cnt = 0
                    For WsInS_idx = ShtSeries_beg_idx To ShtSeries_end_idx
                        If WorkSheetExist(oWb, Replace(ShtSeries_name, "*", WsInS_idx)) = True Then
                            WsInS_cnt = WsInS_cnt + 1
                        Else
                            Exit For
                        End If
                    Next WsInS_idx
                    If WsInS_cnt > 0 Then
                        For WsInS_idx = 0 To WsInS_cnt
                            FailedReason = AppendArray(SheetNameList, Array(Replace(ShtSeries_name, "*", WsInS_idx)))
                            FailedReason = AppendArray(SheetNameLocalList, Array(Replace(ShtSeries_local_name, "*", WsInS_idx)))
                        Next WsInS_idx
                    End If
                Next ShtSeries
            End If
            'Link worksheets
            ReDim FullNameList(0 To UBound(SheetNameList))
            ReDim SheetNameAndRangeList(0 To UBound(SheetNameList))
            Dim SheetNameAndRange As String, SheetName As String, FullName As String
            Dim SheetNameIdx As Integer, ShtColCnt As Long, col_idx As Long
            For SheetNameIdx = 0 To UBound(SheetNameList)
                SheetName = SheetNameList(SheetNameIdx)
                DelTable (SheetNameLocalList(SheetNameIdx))
                On Error Resume Next
                .Worksheets(SheetName).Activate
                On Error GoTo Next_SheetNameIdx_1
                With .ActiveSheet.UsedRange
                    ShtColCnt = .Columns.count
                    If HasFieldNames = True Then
                        For col_idx = 1 To ShtColCnt
                            If IsEmpty(.Cells(1, col_idx)) = True Then
                                ShtColCnt = col_idx - 1
                                Exit For
                            End If
                        Next col_idx
                    End If
                    SheetNameAndRange = SheetName & "!A1:" & ColumnLetter(oWb.ActiveSheet, ShtColCnt) & .Rows.count
                End With '.ActiveSheet.UsedRange
                FullNameList(SheetNameIdx) = .FullName
                SheetNameAndRangeList(SheetNameIdx) = SheetNameAndRange
Next_SheetNameIdx_1:
            Next SheetNameIdx
            .Close False
        End With 'oWb
        .Quit
    End With 'oExcel
    For SheetNameIdx = 0 To UBound(SheetNameList)
        SheetName = IIf(SheetNameLocalList(SheetNameIdx) <> vbNullString, SheetNameLocalList(SheetNameIdx), SheetNameList(SheetNameIdx))
        FullName = FullNameList(SheetNameIdx): SheetNameAndRange = SheetNameAndRangeList(SheetNameIdx): DelTable(SheetName)
        On Error Resume Next
        DoCmd.TransferSpreadsheet acLink, , SheetName, FullName, True, SheetNameAndRange
        On Error GoTo Next_SheetNameIdx_2
Next_SheetNameIdx_2:
    Next SheetNameIdx
    On Error GoTo Err_LinkToWorksheetInWorkbook
Exit_LinkToWorksheetInWorkbook:
    LinkToWorksheetInWorkbook = FailedReason
    Exit Function
Err_LinkToWorksheetInWorkbook:
    FailedReason = Err.Description
    Resume Exit_LinkToWorksheetInWorkbook
End Function
Public Function ExportTblToSht(Wb_path, Tbl_name As String, sht_name As String) As String
    'Export a table to one or more worksheets in case row count over 65535
    On Error GoTo Err_ExportTblToSht: Dim FailedReason As String
    If TableExist(Tbl_name) = False Then
        FailedReason = Tbl_name & " does not exist"
        GoTo Exit_ExportTblToSht
    End If
    If Len(Dir(Wb_path)) = 0 Then
        Dim oExcel As Excel.Application: Set oExcel = CreateObject("Excel.Application")
        With oExcel
            Dim oWb As Workbook: Set oWb = .Workbooks.Add
            With oWb
                .SaveAs Wb_path
                .Close
            End With 'oWb_DailyRpt
            .Quit
        End With 'oExcel
        Set oExcel = Nothing
    End If
    Dim RecordCount As Long: RecordCount = Table_RecordCount(Tbl_name)
    Dim MaxRowPerSht As Long: MaxRowPerSht = 65534
    If RecordCount <= 0 Then
        GoTo Exit_ExportTblToSht
    ElseIf RecordCount <= MaxRowPerSht Then
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, Tbl_name, Wb_path, True, sht_name
    Else 'handle error msg, "File sharing lock count exceeded. Increase MaxLocksPerFile registry entry"
        DAO.DBEngine.SetOption dbMaxLocksPerFile, 40000
        Dim Tbl_COPY_name As String, SQL_cmd As String
        Tbl_COPY_name = Tbl_name & "_COPY": DelTable (Tbl_COPY_name)
        SQL_cmd = "SELECT * " & vbCrLf & "INTO [" & Tbl_COPY_name & "] " & vbCrLf & "FROM [" & Tbl_name & "]" & vbCrLf & ";"
        RunSQL_CmdWithoutWarning (SQL_cmd)
        SQL_cmd = "ALTER TABLE [" & Tbl_COPY_name & "] " & vbCrLf & "ADD record_idx COUNTER " & vbCrLf & ";"
        RunSQL_CmdWithoutWarning (SQL_cmd)
        Dim Sht_part_name As String, Tbl_part_name As String, sht_idx As Long
        Dim ShtCount As Long: ShtCount = Int(RecordCount / MaxRowPerSht)
        For sht_idx = 0 To ShtCount
            Sht_part_name = sht_name
            If sht_idx > 0 Then Sht_part_name = Sht_part_name & "_" & sht_idx
            Tbl_part_name = Tbl_name & "_" & sht_idx: DelTable (Tbl_part_name)
            SQL_cmd = "SELECT * " & vbCrLf & "INTO [" & Tbl_part_name & "] " & vbCrLf & "FROM [" & Tbl_COPY_name & "]" & vbCrLf & "WHERE [record_idx] >= " & sht_idx * MaxRowPerSht + 1 & vbCrLf & "AND [record_idx] <= " & (sht_idx + 1) * MaxRowPerSht & vbCrLf & ";"
            RunSQL_CmdWithoutWarning (SQL_cmd)   'MsgBox (SQL_cmd)
            SQL_cmd = "ALTER TABLE [" & Tbl_part_name & "] " & vbCrLf & "DROP COLUMN [record_idx] " & vbCrLf & ";"
            RunSQL_CmdWithoutWarning (SQL_cmd)
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, Tbl_part_name, Wb_path, True, Sht_part_name
            DelTable (Tbl_part_name)
        Next sht_idx
        DelTable (Tbl_COPY_name)
    End If
Exit_ExportTblToSht:
    ExportTblToSht = FailedReason
    Exit Function
Err_ExportTblToSht:
    FailedReason = Err.Description
    Resume Exit_ExportTblToSht
End Function


