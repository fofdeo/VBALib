Attribute VB_Name = "_VBALib_Misc"


'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function SheetExists(sWSName as String) As Boolean
    Dim iShtCnt as Integer
    For iShtCnt = 1 to ActiveWorkbook.Sheets.Count
        If LCase(Sheets(iShtCnt).Name) = LCase(sWSName) Then
            SheetExists = True: Exit Function
        End If 'ShtName Check
    Next iShtCnt 'Sheet Loop
End Function 'SheetExists
Public Function FileExists(sFileFullPath as String) As Boolean
    Dim objFSO as Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(sFileFullPath) Then FileExists = True
    Set objFSO = Nothing
End Function 'FileExists
Public Function FolderExists(sFolderPath as String) As Boolean
    Dim objFSO as Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(sFolderPath) Then FolderExists = True
    Set objFSO = Nothing
End Function 'FolderExists
Public Function GetFileName(sFilePath as String) As String
    Dim sFileName as String: GetFileName = vbNullString
    If InStr(1, sFilePath, "\") > 0 Then
        sFileName = Split(sFilePath, "\")(UBound(Split(sFilePath,"\")))
        GetFileName = sFileName
    ElseIf InStr(1, sFilePath, "/") > 0 Then
        sFileName = Split(sFilePath, "/")(UBound(Split(sFilePath,"/")))
        GetFileName = sFileName
    End If
End Function 'GetFileName
Public Function IsWbOpen(wbName as String) As Boolean
    Dim iWbCnt as Integer
    For iWbCnt = Workbooks.Count to 1 Step -1
        If Workbooks(iWbCnt).Name = wbName Then
            IsWbOpen = True: Exit Function
        End If
    Next iWbCnt
End Function 'IsWbOpen
Public Function FindString(sFind as String) As Boolean
    Dim RngFound
    Set RngFound = Cells.Find(What:=sFind, After:=ActiveCell, _
        LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False)
    If Not (RngFound is Nothing) Then
        Cells.FindNext(After:=ActiveCell).Select
        FindString = True
    End If
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function Zip(sZipPath as String, sFilePath1 as String, Optional sZipFileName as String) As String
    Dim objApp as Object, iCnt as Integer
    Dim arrFiles, ZipFile
    arrFiles = Array(sFilePath1) 'add sFilePath2 to function & here to append more, or build loop or whatever
    If Right(sZipPath, 1) <> "\" Then sZipPath = sZipPath & "\"
    If sZipFileName <> vbNullString Then
        ZipFile = sZipPath & sZipFileName & ".zip"
    Else
        ZipFile = sZipPath & GetFileName(sFilePath1) & ".zip"
    End If
    If IsArray(arrFiles) Then
        If Len(Dir(ZipFile)) > 0 Then Kill ZipFile
        Open ZipFile For Output as #1
        Print #1, Chr$(80) & Chr$(5) & Chr$(6) & String(18,0)
        Close #1
        Set objApp = CreateObject("Shell.Application")
        For iCnt = LBound(arrFiles) To UBound(arrFiles)
            objApp.Namespace(ZipFile).CopyHere arrFiles(iCnt)
        Next iCnt
    End If
    Set objApp = Nothing
    Zip = ZipFile
End Function
Public Function BrowseFile(Optional sWinMsg as String) As String
    Dim sWinFilter as String
    If sWinMsg = vbNullString Then sWinMsg = "Please select file."
    sWinFilter  = "All Files (*.*),*.*,Excel 2007 Files (*.xlsx),*.xlsx,Excel Files (*.xls),*.xls,Excel Macro Enabled Files (*.xlsm),*.xlsm" 
    BrowseFile = Application.GetOpenFileName(sWinFilter, , sWinMsg, , False)
End Function 'BrowseFile
Public Function BrowseFolder(Optional Hwnd as Long=0,Optional sTitle as String="Please select folder.",Optional BIF_Options as Integer,Optional vRootFolder as Variant) As String
    Dim objShell as Object
    Dim objFolder as Variant
    Dim sFolderFullPath as String
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseFolder(Hwnd,sTitle,BIF_Options,vRootFolder)

    If (Not objFolder Is Nothing) Then 
        On Error Resume Next
        If IsError(objFolder.Items.Item.Path) Then sFolderFullPath = CStr(objFolder): GoTo GotIt
        On Error GoTo 0
        If Len(objFolder.Items.Item.Path) > 3 Then
            sFolderFullPath = objFolder.Items.Item.Path & Application.PathSeparator
        Else
            sFolderFullPath = objFolder.items.Item.Path
        End If
    Else
        GoTo XitProperly
    End If

GoIt:
    BrowseFolder = sFolderFullPath
XitProperly:
    Set objFolder = Nothing
    Set objShell = Nothing
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function GetDate(Optional sDateFormat As String, Optional sDiversionFrom As String, Optional iDiversionValue As Integer) As String
    Dim dDate As Date
    If sDateFormat = vbNullString Then sDateFormat = " mmm-yy"
    If sDiversionFrom = vbNullString Then sDiversionFrom = "m"
    dDate = DateAdd(sDiversionFrom, iDiversionValue, CDate(Now))
    GetDate = Format(dDate, sDateFormat)
End Function
Public Function FileCopy(sSrcFullPath As String, sDestFullPath As String) As String
    Dim objFSO As Object
    If Right(sDestFullPath, 1) <> "\" Then sDestFullPath = sDestFullPath & "\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If GetFileName(sSrcFullPath) <> "" And objFSO.FolderExists(sDestFullPath) Then
        objFSO.CopyFile Source:=sSrcFullPath, Destination:=sDestFullPath, overwritefiles:=True
    End If
    Set objFSO = Nothing
    FileCopy = "CopyFile performed. Source=" & sSrcFullPath & " Destination=" & sDestFullPath
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function PivotCacheRefresh() As Boolean
    Dim pvt 'use in Workbook_Open() to refresh pivot/xml cached lists
    For Each pvt In ActiveSheet.PivotTables
        pvt.PivotCache.MissingItemsLimit = xlMissingItemsNone
        pvt.PivotCache.Refresh
    Next pvt
    ThisWorkbook.RefreshAll
    PivotCacheRefresh = True
End Function
Public Function OpenURL(sURL As String) As String
    Dim objIE As Object
    Set objIE = CreateObject("Internetexplorer.Application")
    objIE.Visible = True
    objIE.Navigate strURL
    Set objIE = Nothing
    OpenURL = sURL
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function NewLog(Optional sShtName As String) As Boolean
    Dim func: ThisWorkbook.Activate
    If sShtName = vbNullString Then sShtName = "Log"
    If SheetExists(sShtName) = False Then
        ThisWorkbook.Worksheets.Add.Name = sShtName
        With ThisWorkbook.Worksheets(sShtName)
            .Cells(1, 1).Value = "Date"
            .Cells(1, 2).Value = "Time"
            .Cells(1, 3).Value = "Log"
        End With
        func = Log("New log named " & sShtName & " has been created.")
		NewLog = True
    End If
End Function
Public Function Log(sLogInfo As String, Optional sShtName As String) As String
    Dim rLastRow
    If sShtName = vbNullString Then sShtName = "Log"
    rLastRow = ThisWorkbook.Worksheets(sShtName).UsedRange.Rows.Count + 1
    With ThisWorkbook.Worksheets(sShtName)
        .Cells(rLastRow, 1).Value = Date
        .Cells(rLastRow, 2).Value = Time
        .Cells(rLastRow, 3).Value = sLogInfo
    End With
    Log = "Date=" & Date & " Time=" & Time & " Message=" & sLogInfo
End Function





'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Sub addLinks()
    For Each cell In Selection
        Call addLink(cell.Value, cell)
    Next
End Sub
Public Sub addLink(ByVal url As String, ByVal cell As Range)
    cell.Worksheet.Hyperlinks.Add Anchor:=cell, Address:=url
End Sub





'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function getAddress(ByVal cell As Range)
    getAddress = cell.Address
End Function
Public Function selectXlDownRange(ByVal cell As Range) As Range
    Application.Volatile ' Force Excel to recalculate on workbook change
    If cell.Offset(1, 0) = vbNullString Then Set selectXlDownRange = cell: Exit Function
    Set selectXlDownRange = Range(cell, cell.End(xlDown))
End Function




'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function ModuleExists(moduleName As String, Optional wb As Workbook) As Boolean
    ' Determines whether a VBA code module with a given name exists. @param wb: The workbook to check for the given module name (defaults to the active workbook).
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim c As Variant ' VBComponent
    On Error GoTo notFound
    Set c = wb.VBProject.VBComponents.Item(moduleName)
    ModuleExists = True
    Exit Function
notFound:
    ModuleExists = False
End Function
Public Sub RemoveModule(moduleName As String, Optional wb As Workbook)
    ' Removes the VBA code module with the given name. @param wb: The workbook to remove the module from (defaults to the active workbook).
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not ModuleExists(moduleName, wb) Then Err.Raise 32000, Description:="Module '" & moduleName & "' not found."
    Dim c As Variant: Set c = wb.VBProject.VBComponents.Item(moduleName) ' VBComponent
    wb.VBProject.VBComponents.Remove c
    ' Sometimes the line above does not remove the module successfully.  When this happens, c.Name does not return an error - otherwise it does.
    On Error GoTo nameError
    Dim n As String: n = c.Name
    On Error GoTo 0
    Err.Raise 32000, Description:="Failed to remove module '" & moduleName & "'.  Try again later."
    ' Everything worked fine (the module was removed)
nameError: 
End Sub
Public Sub ExportModule(moduleName As String, moduleFilename As String, Optional wb As Workbook)
    ' Exports a VBA code module to a text file. @param wb: The workbook that contains the module to export (defaults to the active workbook).
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not ModuleExists(moduleName, wb) Then Err.Raise 32000, Description:="Module '" & moduleName & "' not found."
    wb.VBProject.VBComponents.Item(moduleName).Export moduleFilename
End Sub
Public Function ImportModule(moduleFilename As String, Optional wb As Workbook) As VBComponent
    ' Imports a VBA code module from a text file. @param wb: The workbook that will receive the imported module (defaults to the active workbook).
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Set ImportModule = wb.VBProject.VBComponents.Import(moduleFilename)
End Function
