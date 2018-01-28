


Function SheetExists(strWSName As String) As Boolean
    Dim intCountSheet As Integer
    For intCountSheet = 1 To ActiveWorkbook.Sheets.Count
        If LCase(Sheets(intCountSheet).Name) = LCase(strWSName) Then
            SheetExists = True
            Exit Function
        End If
    Next intCountSheet
End Function
Function FileExists(strFileFullPath As String) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(strFileFullPath) Then FileExists = True
    Set objFSO = Nothing
End Function
Function FolderExists(strFolderPath As String) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(strFolderPath) Then FolderExists = True
    Set objFSO = Nothing
End Function



Function GetFileName(strFilePath As String) As String
    Dim strFileName As String: GetFileName = vbNullString
    If InStr(1, strFilePath, "\") > 0 Then
        strFileName = Split(strFilePath, "\")(UBound(Split(strFilePath, "\")))
        GetFileName = strFileName
    ElseIf InStr(1, strFilePath, "/") > 0 Then
        strFileName = Split(strFilePath, "/")(UBound(Split(strFilePath, "/")))
        GetFileName = strFileName
    End If
End Function
Function FindString(strFind As String) As Boolean
    Dim FoundRange: Set FoundRange = Cells.Find(What:=strFind, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
    If Not (FoundRange Is Nothing) Then: Cells.FindNext(After:=ActiveCell).Select: FindString = True
End Function
Function IsWbOpen(wbName As String) As Boolean
    Dim intCountWb As Integer
    For intCountWb = Workbooks.Count To 1 Step -1
        If Workbooks(intCountWb).Name = wbName Then
            IsWbOpen = True
            Exit Function
        End If
    Next intCountWb
End Function
Function GetDate(Optional strDateFormat As String, Optional strDiversionFrom As String, Optional intDiversionValue As Integer) As String
    Dim dDate As Date
    If strDateFormat = vbNullString Then strDateFormat = " mmm-yy"
    If strDiversionFrom = vbNullString Then strDiversionFrom = "m"
    dDate = DateAdd(strDiversionFrom, intDiversionValue, CDate(Now))
    GetDate = Format(dDate, strDateFormat)
End Function




Function Zip(strZipPath As String, strFilePath1 As String, Optional strZipFileName As String) As String
    Dim objApp As Object, intCount As Integer
    Dim arryFiles, ZipFile
    arryFiles = Array(strFilePath1)
    If Right(strZipPath, 1) <> "\" Then strZipPath = strZipPath & "\"
    If strZipFileName <> "" Then
        ZipFile = strZipPath & strZipFileName & ".zip"
    Else
        ZipFile = strZipPath & GetFileName(strFilePath1) & ".zip"
    End If
    If IsArray(arryFiles) Then
        If Len(Dir(ZipFile)) > 0 Then Kill ZipFile
        Open ZipFile For Output As #1
        Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
        Close #1
        Set objApp = CreateObject("Shell.Application")
        For intCount = LBound(arryFiles) To UBound(arryFiles)
            objApp.Namespace(ZipFile).CopyHere arryFiles(intCount)
        Next intCount
    End If
    Set objApp = Nothing: Zip = ZipFile
End Function
Function Mail( strTo As String, strSubject As String, strBody As String, Optional bSend As Boolean, Optional strCC As String, Optional strSignName As String, Optional strAttachPath1 As String, Optional strAttachPath2 As String, Optional strAttachPath3 As String, Optional strAttachPath4 As String, Optional strAttachPath5 As String)
    Dim objOutApp As Object, objOutMail As Object, objFSO As Object, txtStream As Object, Signature As String, strSignature As String
    Set objOutApp = CreateObject("Outlook.Application")
    Set objOutMail = objOutApp.CreateItem(0)                                              'Get the outlook signature by its default path
    strSignature = "C:\Documents and Settings\" & Environ("username") & "\Application Data\Microsoft\Signatures\" & strSignName & ".htm"
    'strSignature = "C:\Users\" & Environ("username") & "\AppData\Roaming\Microsoft\Signatures\Mysig.htm"
    If Dir(strSignature) <> vbNullString Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set txtStream = objFSO.GetFile(strSignature).OpenAsTextStream(1, -2)
        Signature = txtStream.readall: txtStream.Close
    End If
    On Error Resume Next
    With objOutMail
        .To = strTo: .CC = strCC .Subject = strSubject '.BCC = strBCC
        .HTMLBody = "<style type='text/css'>.style1{font-family:'Futura Bk',Times,serif;font-size:95%;}</style><div class='style1'>" & strBody & "</div>" & Signature
        If strAttachPath1 <> vbNullString Then .Attachments.Add (strAttachPath1) 'You can add files also like this
        If strAttachPath2 <> vbNullString Then .Attachments.Add (strAttachPath2)
        If strAttachPath3 <> vbNullString Then .Attachments.Add (strAttachPath3)
        If strAttachPath4 <> vbNullString Then .Attachments.Add (strAttachPath4)
        If strAttachPath5 <> vbNullString Then .Attachments.Add (strAttachPath5)
        IIf bSend, .Send, .Display
    End With
    On Error GoTo 0: Set objOutMail = Nothing: Set objOutApp = Nothing
End Function





Function BrowseForFile(Optional strWindowMsg As String) As String
    Dim strWindowFilter As String
    If strWindowMsg = vbNullString Then strWindowMsg = "Please select file."
    strWindowFilter = "All Files (*.*),*.*,Excel 2007 Files (*.xlsx),*.xlsx,Excel Files (*.xls),*.xls,Excel Macro Enabled Files (*.xlsm),*.xlsm"
    BrowseForFile = Application.GetOpenFilename(strWindowFilter, , strWindowMsg, , False)
End Function
Public Function BrowseForFolder(Optional Hwnd As Long = 0,Optional sTitle As String = "Please, select a folder", Optional BIF_Options As Integer,Optional vRootFolder As Variant) As String 
    'Optional BIF_Options As Integer = BIF_VALIDATE, _
    Dim objShell As Object,objFolder As Variant, strFolderFullPath As String
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(Hwnd, sTitle, BIF_Options, vRootFolder)
    If (Not objFolder Is Nothing) Then '// NB: If SpecFolder= 0 = Desktop then ....
        On Error Resume Next
        If IsError(objFolder.Items.Item.Path) Then strFolderFullPath = CStr(objFolder): GoTo GotIt
        On Error GoTo 0                '// Is it the Root Dir?...if so change
        strFolderFullPath = objFolder.Items.Item.Path & IIf(Len(objFolder.Items.Item.Path) > 3, Application.PathSeparator, vbNullString)
        'If Len(objFolder.Items.Item.Path) > 3 Then
        '    strFolderFullPath = objFolder.Items.Item.Path & Application.PathSeparator
        'Else
        '    strFolderFullPath = objFolder.Items.Item.Path
        'End If
    Else '// User cancelled
        GoTo XitProperly
    End If
GotIt:
    BrowseForFolder = strFolderFullPath
XitProperly:
    Set objFolder = Nothing
    Set objShell = Nothing
End Function







Function Copy(strSourceFullPath As String, strCopyToDestination As String) As String
    Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Right(strCopyToDestination, 1) <> "\" Then strCopyToDestination = strCopyToDestination & "\"
    If GetFileName(strSourceFullPath) <> vbNullString And objFSO.FolderExists(strCopyToDestination) Then
        objFSO.CopyFile Source:=strSourceFullPath, Destination:=strCopyToDestination, overwritefiles:=True
    End If
    Set objFSO = Nothing
    Copy = "CopyFile performed. Source=" & strSourceFullPath & " Destination=" & strCopyToDestination
End Function

Function PivotCacheRefresh() As Boolean
    Dim pvt
    For Each pvt In ActiveSheet.PivotTables
        pvt.PivotCache.MissingItemsLimit = xlMissingItemsNone
        pvt.PivotCache.Refresh
    Next pvt
    ThisWorkbook.RefreshAll
    PivotCacheRefresh = True
End Function

'@Name: OpenURL
'@Description: Opens an url in the internet explorer. This function could be assigned to button.
'@Version:1.0
'@Autor:
'@Date: 20.12.2011
'Input parameters
    '@Param strURL: String. Url.
'Output Parameters
    'Action: Opens the IExplorer with url address the input argument.
	'@Param OpenURL: String. The opened link.
Function OpenURL(strURL As String) As String
    
    Dim objIE As Object
    Set objIE = CreateObject("Internetexplorer.Application")
    objIE.Visible = True
    objIE.Navigate strURL
    Set objIE = Nothing
    
    OpenURL = strURL
    
End Function

'@Name: NewLog
'@Description: Used to create a sheet that would be used by the function Log() where log information would be stored.
'@Version: 1.0
'@Autor: velin.georgiev@gmail.com
'@Date: 20.12.2011
'Input parameters:
    '@Param strSheetName: Optional String. Name of a sheet where the log information where log information would be stored. Default: Log
'Output patameters:
    '@Param NewLog: Boolean. True is success.
	'@Param: Action. Creates a sheet.
    
Function NewLog(Optional strSheetName As String) As Boolean
    
    Dim func
    ThisWorkbook.Activate
    If strSheetName = "" Then strSheetName = "Log"
    'SheetExists function should be available within this module
    If SheetExists(strSheetName) = False Then
        ThisWorkbook.Worksheets.Add.Name = strSheetName
        With ThisWorkbook.Worksheets(strSheetName)
            .Cells(1, 1).Value = "Date"
            .Cells(1, 2).Value = "Time"
            .Cells(1, 3).Value = "Log"
        End With
        func = Log("New log named " & strSheetName & " has been created.")
		NewLog = True
    End If
    
End Function

'@Name: Log
'@Description: Used to record action notes (from macro execution) in a sheet named 'Log' by default.
' It depends on the programmer what should be logged in the log
' so this function is used depending on the programmer needs.
' It can be applied on every line of the macro if an event information has to be recorded in the Log sheet.
' Before the use of this function a new sheet should be created to store the logs. The NewLog() function could do this for you.
'@Version: 1.0
'@Autor: velin.georgiev@gmail.com
'@Date: 20.12.2011
'Input parameters:
    '@Param strLogInfo: String. Describes error or some information of a taken action or event within the vba module.
    '@Param strSheetName: Optional String. Name of a sheet where the log information would be stored.
'Output patameters:
    '@Param Date: Data Record. Enters the current date in the Log sheet
    '@Param Time: Data Record. Enters the current time in the Log sheet
    '@Param strLogInfo: Data Record. Enters the strLogInfo string as a text in the Log sheet
Function Log(strLogInfo As String, Optional strSheetName As String) As String

    Dim rngLastRow
    If strSheetName = "" Then strSheetName = "Log"
    rngLastRow = ThisWorkbook.Worksheets(strSheetName).UsedRange.Rows.Count + 1
    With ThisWorkbook.Worksheets(strSheetName)
        .Cells(rngLastRow, 1).Value = Date
        .Cells(rngLastRow, 2).Value = Time
        .Cells(rngLastRow, 3).Value = strLogInfo
    End With
    Log = "Date=" & Date & " Time=" & Time & " Message=" & strLogInfo
    
End Function