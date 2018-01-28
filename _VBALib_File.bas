'====================================================================================================
Attribute VB_Name = "_VBALib_File"
Option Explicit
'----------------------------------------------------------------------------------------------------
Public Const ZipTool_local_path = "\7za\7za"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'----------------------------------------------------------------------------------------------------
Enum FileTypes
    'Used by:   SelectNewOrExistingFile, SelectExistingFile, GetCustomFilterList
    AnyExtension = 0: Custom = 99
    ExcelFiles = 1: ExcelFilesTemplate = 2
    WordFiles = 3: WordFilesTemplate = 4
    TextFiles = 5: CSVFiles = 6
End Enum

Enum GetFileInfo
    'Used by: GetFileInfo
    PathOnly = 1: NameAndExtension = 2: NameOnly = 3: 
    ExtensionOnly = 4: ParentFolder = 5
    FileExists = 6: FolderExists = 7
    DateLastMod = 8: FileSizeKB = 9
End Enum
'====================================================================================================





'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function SelectNewOrExistingFile(Optional FileType As FileTypes = 0, Optional MenuTitleName = "Select File", Optional StartingPath = "WBPath", Optional CustomFilter As String = "Any File (*.*), *.*") As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' SelectNewOrExistingFile   - Select a new or an existing file, using custom filters for specific file types if needed
        '                               New Function in Excel 2007; will not work with previous versions of Excel (http://msdn.microsoft.com/en-us/library/bb209903(v=office.12).aspx)
        '                           - In : Optional FileType as FileTypes (defined above, specify file filters, by default any file)
        '                               Optional MenuTitleName = "Select File" (Default)
        '                               Optional Strpath = Workbook Path (Default)
        '                               Optional CustomFilter As String = "Any File (*.*), *.*" (Custom Filter if User-defined)
        '                           - Out: Full Path to selected file, or FALSE if user cancelled
        '                           - Requires: Function ReturnCustomFilterList
        '                           - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim OutputFile As Variant
        On Error GoTo IsError
        CustomFilter = GetCustomFilterList(FileType, CustomFilter)
        If StartingPath = "WBPath" Then StartingPath = CStr(ThisWorkbook.Path)
        Do
                OutputFile = Application.GetSaveAsFilename(StartingPath, CustomFilter, 1, MenuTitleName)
                If GetFileInfo(CStr(OutputFile), FileExists) = False Then
                        Exit Do
                Else
                        If vbYes = MsgBox("File already exists, replace existing file?" & vbNewLine & vbNewLine & OutputFile, vbYesNo, "Replace existing file?") Then Exit Do
                End If
        Loop
        SelectNewOrExistingFile = OutputFile
        Exit Function
IsError:
        SelectNewOrExistingFile = CVErr(xlErrNA)
        Debug.Print "Error in SelectNewOrExistingFile: " & Err.Number & ": " & Err.Description
End Function
Public Function SelectExistingFolder(Optional MenuTitleName As String = "Select Folder", Optional ByVal StartingPath As String = "WBPath") As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' SelectExistingFolder  - Selecting an existing folder
        '                       - In :  Optional MenuTitleName = "Select Folder" (Default)
        '                               Optional Strpath = Workbook Path (Default)
        '                       - Out: Folder Path including final backslash "\"
        '                       - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim fldr As FileDialog
        On Error GoTo IsError
        If StartingPath = "WBPath" Then StartingPath = CStr(ThisWorkbook.Path)
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
        With fldr
                .InitialView = msoFileDialogViewDetails
                .Title = MenuTitleName
                .AllowMultiSelect = False
                .InitialFileName = StartingPath
                If .Show <> -1 Then GoTo UserCancelled
                SelectExistingFolder = .SelectedItems(1) & "\"
        End With
        Exit Function
UserCancelled:
        SelectExistingFolder = False
                Exit Function
IsError:
        SelectExistingFolder = CVErr(xlErrNA)
        Debug.Print "Error in SelectExistingFolder: " & Err.Number & ": " & Err.Description
End Function
Public Function SelectExistingFile(Optional FileType As FileTypes = 0, Optional MenuTitleName = "Select File", Optional StartingPath = "WBPath", Optional CustomFilter As String = "Any File (*.*), *.*") As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' SelectExistingFile    - Selecting an exisiting file, using custom filters for pre-defined file types or create a new custom file type
        '                       - In : Optional FileType as FileTypes (defined above, specify file filters, by default any file)
        '                               Optional MenuTitleName = "Select File" (Default)
        '                               Optional Strpath = Workbook Path (Default)
        '                               Optional CustomFilter As String = "Any File (*.*), *.*" (Custom Filter if User-defined)
        '                       - Out: Full Path to selected file, or FALSE if user cancelled
        '                       - Requires: Function ReturnCustomFilterList
        '                       - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        
        CustomFilter = GetCustomFilterList(FileType, CustomFilter)
        If StartingPath = "WBPath" Then StartingPath = CStr(ThisWorkbook.Path)
        ChDir StartingPath
        SelectExistingFile = Application.GetOpenFilename(FileFilter:=CustomFilter, Title:=MenuTitleName, MultiSelect:=False)
        Exit Function
IsError:
        SelectExistingFile = CVErr(xlErrNA)
        Debug.Print "Error in SelectExistingFile: " & Err.Number & ": " & Err.Description
End Function



'====================================================================================================
' [Get] File Info 
'----------------------------------------------------------------------------------------------------
Private Function GetCustomFilterList(FileTypeNumber As FileTypes, CustomFilter As String) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' ReturnCustomFilterList    - Returns custom filter lists for each specified type of file
        '                           - In : FileTypeNumber FileType as FileTypes (defined above, specify file filters, by default any file)
        '                                    CustomFilter As String = "Any File (*.*), *.*" (only used if custom filetypes is selected)
        '                           - Out: FilterList as string
        '                           - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim ReturnValue As String
        On Error GoTo IsError
        Select Case FileTypeNumber
                Case 0
                        ReturnValue = "Any File (*.*),*.*"
                Case 1
                        ReturnValue = "Excel File (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls"
                Case 2
                        ReturnValue = "Excel File or Excel Template (*.xlsx; *.xlsm; *.xls; *.xlt; *.xltx; *.xltm), *.xlsx; *.xlsm; *.xls; *.xlt; *.xltx; *.xltm"
                Case 3
                        ReturnValue = "Word File (*.docx; *.docm; *.doc), *.docx; *.docm; *.doc"
                Case 4
                        ReturnValue = "Word File or Word Template (*.docx; *.docm; *.doc; *.dotx; *.dotm; *.dot), *.docx; *.docm; *.doc; *.dotx; *.dotm; *.dot"
                Case 5
                        ReturnValue = "Text File (*.txt; *.dat), *.txt; *.dat"
                Case 6
                        ReturnValue = "CSV File (*.csv), *.csv"
                Case 99
                        ReturnValue = CustomFilter
        End Select
        GetCustomFilterList = ReturnValue
        Exit Function
IsError:
        GetCustomFilterList = CVErr(xlErrNA)
        Debug.Print "Error in GetCustomFilterList: " & Err.Number & ": " & Err.Description
End Function
Public Function GetFileInfo(FN As String, FileInfo As GetFileInfo, Optional ShowErrorPopup As Boolean = False) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' GetFileInfo        - Returns key file information for a file or folder passed to the function, uses the enumeration GetFileInfo
        '                            1: PathOnly            (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "C:\USEPA\BMDS212\00Hill.exe")
        '                            2: NameAndExtension    (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "00Hill.exe")
        '                            3: NameOnly            (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "00Hill")
        '                            4: ExtensionOnly       (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "exe")
        '                            5: ParentFolder        (FN = "C:\USEPA\BMDS212\",           Return = "C:\USEPA\")
        '                            6: FileExists          (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = TRUE)
        '                            7: FolderExists        (FN = "C:\USEPA\BMDS212\",           Return = TRUE)
        '                            8: DateLastMod         (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "5/20/2010 1:23:56 AM")
        '                            9: FileSizeKB          (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = 12.8)
        '                       (May also display a popup message if file or folder doesn't exist)
        '                    - In : FN As String, FileInfo As GetFileInfo
        '                    - Out: Depends on the file info type selected, error if error
        '                    - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim fso As Object
        On Error GoTo IsError
        Set fso = CreateObject("Scripting.FileSystemObject")
        Select Case FileInfo
                Case 1
                        GetFileInfo = fso.GetParentFolderName(FN) & "\"
                Case 2
                        GetFileInfo = fso.GetFileName(FN)
                Case 3
                        GetFileInfo = fso.GetBaseName(FN)
                Case 4
                        GetFileInfo = fso.GetExtensionName(FN)
                Case 5
                        GetFileInfo = fso.GetParentFolderName(FN) & "\"
                Case 6
                        GetFileInfo = fso.FileExists(FN)
                        If ShowErrorPopup = True And GetFileInfo = False Then MsgBox "Error- file doesn't exist!" & vbNewLine & vbNewLine & FN, vbCritical, "File does not exist!"
                Case 7
                        GetFileInfo = fso.FolderExists(FN)
                        If ShowErrorPopup = True And GetFileInfo = False Then MsgBox "Error- folder doesn't exist!" & vbNewLine & vbNewLine & FN, vbCritical, "Folder does not exist!"
                Case 8
                        GetFileInfo = CStr(fso.GetFile(FN).DateLastModified)
                Case 9
                        GetFileInfo = FileLen(FN) / 1000
                Case Else
                        GoTo IsError
        End Select
        Exit Function
IsError:
        GetFileInfo = CVErr(xlErrNA)
        Debug.Print "Error in GetFileInfo: " & Err.Number & ": " & Err.Description & vbNewLine & FN
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function MakeDirString(PathString As String) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' MakeDirString      - Adds a parenthesis to the end of a path if it doesn't already exist
        '                    - In : PathString As String
        '                    - Out: MakeDirString as string if valid, error if not valid
        '                    - Last Updated: 7/3/11 by AJS
        '---------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        If Right(PathString, 1) <> "\" Then
                MakeDirString = PathString & "\"
        Else
                MakeDirString = PathString
        End If
        Exit Function
IsError:
        MakeDirString = CVErr(xlErrNA)
        Debug.Print "Error in MakeDirString: " & Err.Number & ": " & Err.Description & vbNewLine & PathString
End Function
Public Function MakeDirFullPath(Path As String) As Boolean
        '-----------------------------------------------------------------------------------------------------------
        ' MakeDirFullPath   - Creates the full path directory if it doesn't already exist, can for example
        '                       create C:\Temp\Temp\Temp if it doesn't alreay dexist
        '                   - In : Path as String
        '                   - Out: TRUE if path exists, FALSE if path doesn't exist
        '                   - Last Updated: 7/2/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim UncreatedPaths As Collection, EachPath As Variant
        Set UncreatedPaths = New Collection
        Dim NewPath As String
        
        On Error GoTo IsError
        NewPath = Path
        Do While GetFileInfo(NewPath, FolderExists) = False
                UncreatedPaths.Add NewPath
                NewPath = GetFileInfo(NewPath, ParentFolder)
        Loop
        Do While UncreatedPaths.Count > 0
                MkDir UncreatedPaths(UncreatedPaths.Count)
                UncreatedPaths.Remove UncreatedPaths.Count
        Loop
        MakeDirFullPath = GetFileInfo(Path, FolderExists)
        Exit Function
IsError:
        MakeDirFullPath = GetFileInfo(Path, FolderExists)
        Debug.Print "Error in MakeDirFullPath: " & Err.Number & ": " & Err.Description & vbNewLine & Path
End Function
Public Function FileListInFolder(ByVal PathName As String, Optional ByVal FileFilter As String = "*.*") As Collection
        '-----------------------------------------------------------------------------------------------------------
        ' FileListInFolder   - Returns a collection of files in a given folder with the specified filter
        '                       Can filter by a certain type of filename, if file filter is set to equal a certain extension
        '                       Replacement for Application.FileSearch, removed from Excel 2007
        '                       Uses MSDOS Dir function: http://www.computerhope.com/dirhlp.htm
        '                    - In : PathName As String, Optional FileFilter As String
        '                    - Out: A string collection of file names in the specified folder
        '                    - Created: Greg Haskins
        '                    - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim sTemp As String, sHldr As String
        Dim RetVal As New Collection
        
        On Error GoTo IsError
        If Right$(PathName, 1) <> "\" Then PathName = PathName & "\"
        sTemp = Dir(PathName & FileFilter)
        If sTemp = "" Then
                Set FileListInFolder = RetVal
                Exit Function
        Else
                RetVal.Add sTemp
        End If
        Do
                sHldr = Dir
                If sHldr = "" Then Exit Do
                'sTemp = sTemp & "|" & sHldr
                RetVal.Add sHldr
        Loop
        'FileList = Split(sTemp, "|")
        Set FileListInFolder = RetVal
        Exit Function
IsError:
        FileListInFolder.Add CVErr(xlErrNA)
        Debug.Print "Error in FileListInFolder: " & Err.Number & ": " & Err.Description & vbNewLine & PathName & FileFilter
End Function




'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function Kill2(ByVal PathName As String) As Boolean
        '----------------------------------------------------------------
        ' Kill2             - Deletes file; continues until succesfullly deleted
        '                   - In : ByVal PathName As String
        '                   - Out: Boolean true if file is succesfully removed
        '                   - Last Updated: 7/3/11 by AJS
        '----------------------------------------------------------------
        Dim TimeOut As String
        TimeOut = Now + TimeValue("00:00:10")
        On Error Resume Next
                Do While GetFileInfo(PathName, FileExists) = True
                        Kill PathName
                        If Now > TimeOut Then
                                MsgBox "Error- File deletion has time out, file cannot be deleted:" & vbNewLine & vbNewLine & PathName, vbCritical, "Error in deleting file"
                                GoTo IsError
                        End If
                Loop
        On Error GoTo 0
        Kill2 = True
        Exit Function
IsError:
        Kill2 = False
        Debug.Print "Error in Kill2: " & Err.Number & ": " & Err.Description & vbNewLine & PathName
End Function
Public Function FileCopy2(ByVal SourceFile As String, ByVal DestinationFile As String) As Boolean
        '----------------------------------------------------------------
        ' FileCopy2             - Revised version of FileCopy that will return TRUE when file is actually copied
        '                       - In : SourceFile As String, DestinationFile As String
        '                       - Out: Boolean true if file is succesfully copied; false otherwise
        '                       - Last Updated: 7/3/11 by AJS
        '----------------------------------------------------------------
        Dim TimeOut As String
        TimeOut = Now + TimeValue("00:00:10")
        If GetFileInfo(SourceFile, FileExists) = False Then
                MsgBox "Error- file does not exist and cannot be copied:" & vbNewLine & vbNewLine & SourceFile, vbCritical, "File cannot be copied"
                GoTo IsError
        End If
        If GetFileInfo(DestinationFile, FileExists) = True Then z_Files.Kill2 DestinationFile
        On Error Resume Next
        Do While GetFileInfo(DestinationFile, FileExists) = False
                FileCopy SourceFile, DestinationFile
                If Now > TimeOut Then
                        MsgBox "Error- File copy has timed out, file was probably not succesfully copied (may be open?):" & vbNewLine & vbNewLine & _
                                        "Source: " & SourceFile & vbNewLine & _
                                        "Destination: " & DestinationFile, vbCritical, "Error in copying file"
                        GoTo IsError
                End If
        Loop
        FileCopy2 = True
        Exit Function
IsError:
        FileCopy2 = False
        Debug.Print "Error in FileCopy2: " & Err.Number & ": " & Err.Description & vbNewLine & SourceFile & vbNewLine & DestinationFile
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Private Function DoesFileExist(FN As String) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' DoesFileExist      - Alternate way to test to see if file exists (instead of GetFileInfo)
        '                    - In : FN as String
        '                    - Out: TRUE/FALSE if filename is valid
        '-----------------------------------------------------------------------------------------------------------
    Dim fso As Object
    On Error GoTo IsError
    Set fso = CreateObject("Scripting.FileSystemObject")
    DoesFileExist = fso.FileExists(FN)
    Exit Function
IsError:
    DoesFileExist = CVErr(xlErrNA)
    Debug.Print "Error in Private Function DoesFileExist: " & Err.Number & ": " & Err.Description
End Function
Public Function IsValidFileName(FN As String) As Boolean
        '-----------------------------------------------------------------------------------------------------------
        ' IsValidFileName    - Returns true if filename is valid using the Win32 naming scheme
        '                    - Adapted from: http://www.bytemycode.com/snippets/snippet/334/
        '                    - In : FN as String
        '                    - Out: TRUE/FALSE if filename is valid
        '                    - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim RE As Object, REMatches As Object
        On Error GoTo IsError
        Set RE = CreateObject("vbscript.regexp")
        
        With RE
                .MultiLine = False
                .Global = False
                .IgnoreCase = True
                .Pattern = "[\\\/\:\*\?\" & Chr(34) & "\<\>\|]" 'If any of the following characters are found: \ / : * ? " < > |
        End With
        Set REMatches = RE.Execute(FN)
        If REMatches.Count > 0 Or FN = "" Then
                MsgBox "Filename not valid: " & vbNewLine & FN, vbCritical, "Filename not valid"
                IsValidFileName = False
        Else
                IsValidFileName = True
        End If
        Exit Function
IsError:
        IsValidFileName = False
        Debug.Print "Error in IsValidFileName: " & Err.Number & ": " & Err.Description & vbNewLine & FN
End Function
Public Function IsFileOpen(FN As String) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' IsFileOpen    - Returns TRUE if file is currently open, FALSE if it's not open, or error if other error occurs
        '               - Adapted from: http://www.vbaexpress.com/kb/getarticle.php?kb_id=468
        '               - In : FN as String
        '               - Out: TRUE if file is currently open, FALSE if it's not open, or error if other error occurs
        '               - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim iErr As Long, iFilenum As Long
        On Error Resume Next
                Err.Clear
                iFilenum = FreeFile()
                Open FN For Input Lock Read As #iFilenum
                Close iFilenum
                iErr = Err
        On Error GoTo 0
        Select Case iErr
                Case 0:    IsFileOpen = False
                Case 70:   IsFileOpen = True
                Case Else: Error iErr
        End Select
End Function







'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function FileExists0(file As String) As Boolean
    Dim objFSO As Object: FileExists = False
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.fileExists(file) Then fileExists = True
    Set objFSO = Nothing
End Function
Public Function FileExists(ByVal sFile As String, Optional bFindFolders As Boolean = False) As Boolean
    'Purpose:       Return True if the File exists (even Hidden)
    'Params:        sFile: FileName to check, either FullName or relative to current Directory
    '               bFindFolders: if sFile is a Folder, FilExists() returns False unless this is True
    'Note:          Does not examine sub-Folders, does include [read-only, hidden, system] Files
    Dim attrs As Long: attrs = (vbReadOnly Or vbHidden Or vbSystem)
    If bFindFolders Then attrs = (attrs Or vbDirectory)                  'Include Folders as well
    FileExists = False: On Error Resume Next        'If Dir() returns something, the File exists
    FileExists = (Dir(TrimTrailingChars(sFile, "/\"), attrs) <> vbNullString)
End Function
Public Function FolderExists(sFolder As String) As Boolean
    ' Determines whether a folder with the given name exists.
    On Error Resume Next
    FolderExists = ((GetAttr(folderName) And vbDirectory) = vbDirectory)
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function GetDirectoryName(ByVal p As String) As String
    'Returns the folder name of a path (removes the last component of the path).
    Dim i As Integer: i = InStrRev(p, "\")
    p = NormalizePath(p)
    GetDirectoryName = IIf(i = 0, vbNullString, Left(p, i - 1))
End Function
Public Function GetFilename(ByVal p As String) As String
    'Returns the filename of a path (the last component of the path).
    Dim i As Integer: i = InStrRev(p, "\")
    p = NormalizePath(p)
    GetFilename = Mid(p, i + 1)
End Function
Public Function GetFileExtension(ByVal p As String) As String
    'Returns the extension of a filename (including the dot).
    Dim i As Integer: i = InStrRev(p, ".")
    GetFileExtension = IIf(i > 0, Mid(p, i), vbNullString)
End Function
Public Function GetFileUpdateTime(ByVal file As String) As Double
    GetFileUpdateTime = FileDateTime(file)
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function ListFiles(filePattern As String) As Variant()
    ' Lists all files matching the given pattern.
    ' @param filePattern: A directory name, or a path with wildcards:
    '  - C:\Path\to\Folder
    '  - C:\Path\to\Folder\ExcelFiles.xl*
    ListFiles = ListFiles_Internal(filePattern, vbReadOnly Or vbHidden Or vbSystem)
End Function
Public Function ListFolders(folderPattern As String) As Variant()
    ' Lists all folders matching the given pattern.
    ' @param folderPattern: A directory name, or a path with wildcards:
    '  - C:\Path\to\Folder
    '  - C:\Path\to\Folder\OtherFolder_*
    ListFolders = ListFiles_Internal(folderPattern, vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)
End Function
Public Function GetTempPath() As String
    ' Returns the path to a folder that can be used to store temporary files.
    Const MAX_PATH = 256
    Dim folderName As String: folderName = String(MAX_PATH, 0)
    Dim ret As Long: ret = GetTempPathA(MAX_PATH, folderName)
    If ret <> 0 Then
        GetTempPath = Left(folderName, InStr(folderName, Chr(0)) - 1)
    Else
        Err.Raise 32000, Description:="Error getting temporary folder."
    End If
End Function
Private Function ListFiles_Internal(filePattern As String, attrs As Long) As Variant()
    Dim folderName As String, currFileName As String, filesList As New VBALib_List
    If FolderExists(filePattern) Then
        filePattern = NormalizePath(filePattern) & "\"
        folderName = filePattern
    Else
        folderName = GetDirectoryName(filePattern) & "\"
    End If
    currFileName = Dir(filePattern, attrs)
    While currFileName <> vbNullString
        If (attrs And vbDirectory) = vbDirectory Then
            If FolderExists(folderName & currFileName) And currFileName <> "." And currFileName <> ".." Then filesList.Add folderName & currFileName
        Else
            filesList.Add folderName & currFilename
        End If
        currFilename = Dir
    Wend
    ListFiles_Internal = filesList.Items
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function CombinePaths(p1 As String, p2 As String) As String
    'Merges two path components into a single path.
    CombinePaths = TrimTrailingChars(p1, "/\") & "\" & TrimLeadingChars(p2, "/\")
End Function
Public Function NormalizePath(ByVal p As String) As String
    ' Fixes slashes within a path:
    '  - Converts all forward slashes to backslashes
    '  - Removes multiple consecutive slashes (except for UNC paths)
    '  - Removes any trailing slashes
    Dim isUNC As Boolean: isUNC = StartsWith(p, "\\")
    p = Replace(p, "/", "\")
    While InStr(p, "\\") > 0
        p = Replace(p, "\\", "\")
    Wend
    If isUNC Then p = "\" & p
    NormalizePath = TrimTrailingChars(p, "\")
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function WriteFile(ByVal File As String, ByVal content As String) As String    
   Open File For Output As #1
   Print #1, content
   Close #1
   WriteFile = "File updated yet"
End Function
Public Function ReadFile(ByVal File As String, Optional createFile As Boolean) As String
    If (IsMissing(createFile)) Then createFile = False
    If (fileExists(File) = False) Then
        If (createFile = True) Then
            temp = writeFile(File, vbNewLine)
        Else
            readFile = "Error : File does not exists"
            Exit Function
        End If
    End If
    Dim MyString, MyNumber
    Open File For Input As #1 ' Open file for input.
    fileContent = vbNewLine
    Do While Not EOF(1) ' Loop until end of file.
        Line Input #1, MyString
        Debug.Print MyString
        fileContent = fileContent & MyString & " "
    Loop
    Close #1 ' Close file.
    ReadFile = fileContent
End Function
Public Function ReadFile_Truncate(ByVal File As String, Optional createFile As Boolean) As String
    If (IsMissing(createFile)) Then createFile = False
    ReadFile_Truncate = Left(readFile(File, createFile), 30000)
End Function
Public Sub CopyFileBypassErr(Src As String, des As String)
    'Copy File without error msg
    Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
    'object.copyfile,source,destination,file overright(True is default)
    objFSO.CopyFile Src, des, True
    Set objFSO = Nothing
End Sub



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Sub MkDirRecursive(folderName As String)
    ' Creates the given directory, including any missing parent folders.
    MkDirRecursiveInternal folderName, folderName
End Sub
Private Sub MkDirRecursiveInternal(folderName As String, originalFolderName As String)
    ' Too many recursive calls to this function (GetDirectoryName will eventually return an empty string)
    If folderName = vbNullString Then Err.Raise 32000, Description:="Failed to create folder: " & originalFolderName
    Dim parentFolderName As String: parentFolderName = GetDirectoryName(folderName)
    If Not FolderExists(parentFolderName) Then MkDirRecursiveInternal parentFolderName, originalFolderName
    If Not FolderExists(folderName) Then MkDir folderName
End Sub



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function ExtractZipInDir(SrcDir As String, DesDir As String, Optional Criteria As String = "", Optional DeleteZipFile As Boolean = False) As String
    'Unzip multiple files in directory
    On Error GoTo Err_ExtractZip
    Dim FailedReason As String, Result As String
    Criteria = SrcDir & Criteria
    Result = Dir(Criteria)
    Do While Len(Result) > 0
        Call ExtractZip(SrcDir & Result, DesDir, DeleteZipFile)
        Result = Dir
    Loop

Exit_ExtractZip:
    ExtractZipInDir = FailedReason
    Exit Function

Err_ExtractZip:
    FailedReason = Err.Description
    Resume Exit_ExtractZip

End Function
Public Function ExtractZip(Src As String, DesDir As String, Optional DeleteZipFile As Boolean = False) As String
    'Unzip a file
    On Error GoTo Err_ExtractZip
    Dim ShellCmd As String, Success As Boolean
    Dim FailedReason As String, ZipTool_path As String
    ZipTool_path = [CurrentProject].[Path] & ZipTool_local_path    
    ShellCmd = ZipTool_path & " x " & Src & " -o" & DesDir & " -ry"
    'MsgBox ShellCmd
    Success = ShellAndWait(ShellCmd, vbHide)
    If Success = True And DeleteZipFile = True Then Kill Src

Exit_ExtractZip:
    ExtractZip = FailedReason
    Exit Function

Err_ExtractZip:
    FailedReason = Err.Description
    Resume Exit_ExtractZip

End Function



'====================================================================================================
' [FTP] 
'----------------------------------------------------------------------------------------------------
Public Function FTPUpload(Site, sUsername, sPassword, sLocalFile, sRemotePath, Optional Delay As Integer = 1000) As String
    'FTP Upload File
    On Error GoTo Err_FTPUpload
    Const OpenAsDefault = -2, FailIfNotExist = 0, ForReading = 1, ForWriting = 2
    Dim FailedReason As String, sResults As String, sFTPTemp As String, sFTPTempFile As String
    Dim fFTPScript As Object, sFTPScript As String, fFTPResults As Object, sFTPResults As String
    Dim oFTPScriptFSO As Object: Set oFTPScriptFSO = CreateObject("Scripting.FileSystemObject")
    Dim oFTPScriptShell As Object: Set oFTPScriptShell = CreateObject("WScript.Shell")
    sRemotePath = Trim(sRemotePath)
    sLocalFile = Trim(sLocalFile)
    
    'Check Path: if it contains spaces, add quotes to ensure it parses correctly.
    If InStr(sRemotePath, " ") > 0 Then
        If Left(sRemotePath, 1) <> """" And Right(sRemotePath, 1) <> """" Then
            sRemotePath = """" & sRemotePath & """"
        End If
    End If
    If InStr(sLocalFile, " ") > 0 Then
        If Left(sLocalFile, 1) <> """" And Right(sLocalFile, 1) <> """" Then
            sLocalFile = """" & sLocalFile & """"
        End If
    End If 
    If Len(sRemotePath) = 0 Then sRemotePath = "\"       'Check that a remote path was passed.
    
    'Check the local path and file to ensure that either the a file that exists was passed or a wildcard was passed.
    If InStr(sLocalFile, "*") Then
        If InStr(sLocalFile, " ") Then
            FailedReason = "Error: Wildcard uploads do not work if the path contains a space." & vbCrLf
            FailedReason = FailedReason & "This is a limitation of the Microsoft FTP client."
            GoTo Exit_FTPUpload
        End If
    ElseIf Len(sLocalFile) = 0 Or Not oFTPScriptFSO.FileExists(sLocalFile) Then
        FailedReason = "Error: File Not Found."             'nothing to upload
        GoTo Exit_FTPUpload
    End If 'Path Checks    
    
    'build input file for ftp command    
    sFTPScript = sFTPScript & "USER " & sUsername & vbCrLf
    sFTPScript = sFTPScript & sPassword & vbCrLf
    sFTPScript = sFTPScript & "cd " & sRemotePath & vbCrLf
    sFTPScript = sFTPScript & "binary" & vbCrLf
    sFTPScript = sFTPScript & "prompt n" & vbCrLf
    sFTPScript = sFTPScript & "put " & sLocalFile & vbCrLf
    sFTPScript = sFTPScript & "quit" & vbCrLf & "quit" & vbCrLf & "quit" & vbCrLf

    sFTPTemp = oFTPScriptShell.ExpandEnvironmentStrings("%TEMP%")
    sFTPTempFile = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
    sFTPResults = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
    
    'Write the input file for the ftp command to a temporary file.
    Set fFTPScript = oFTPScriptFSO.CreateTextFile(sFTPTempFile, True)
    fFTPScript.WriteLine (sFTPScript)
    fFTPScript.Close
    Set fFTPScript = Nothing
    oFTPScriptShell.Run "%comspec% /c FTP -n -s:" & sFTPTempFile & " " & Site & " > " & sFTPResults, 0, True
    Sleep Delay
    
    'Check results of transfer.  
    Set fFTPResults = oFTPScriptFSO.OpenTextFile(sFTPResults, ForReading, FailIfNotExist, OpenAsDefault)
    sResults = fFTPResults.ReadAll
    fFTPResults.Close
    
    If InStr(sResults, "226 Transfer complete.") > 0 Then
        FailedReason = vbNullString
    ElseIf InStr(sResults, "File not found") > 0 Then
        FailedReason = "Error: File Not Found"
    ElseIf InStr(sResults, "cannot log in.") > 0 Then
        FailedReason = "Error: Login Failed."
    Else
        FailedReason = "Error: Unknown."
    End If
    oFTPScriptFSO.DeleteFile (sFTPTempFile)
    oFTPScriptFSO.DeleteFile (sFTPResults)
    Set oFTPScriptFSO = Nothing
    oFTPScriptShell.CurrentDirectory = sOriginalWorkingDirectory
    Set oFTPScriptShell = Nothing
    
Exit_FTPUpload:
    FTPUpload = FailedReason
    Exit Function
    
Err_FTPUpload:
    FailedReason = Err.Description
    Resume Exit_FTPDownload
    
End Function
Function FTPDownload(Site, sUsername, sPassword, sLocalPath, sRemotePath, sRemoteFile, Optional Delay As Integer = 1000) As String
    'Ftp download file
    On Error GoTo Err_FTPDownload
    Dim FailedReason As String
    
    Dim oFTPScriptFSO As Object: Set oFTPScriptFSO = CreateObject("Scripting.FileSystemObject")
    Dim oFTPScriptShell As Object: Set oFTPScriptShell = CreateObject("WScript.Shell") 
    
    sRemotePath = Trim(sRemotePath)
    sLocalPath = Trim(sLocalPath)
    
    '----------Path Checks---------
    If InStr(sRemotePath, " ") > 0 Then
        If Left(sRemotePath, 1) <> """" And Right(sRemotePath, 1) <> """" Then
            sRemotePath = """" & sRemotePath & """"
        End If
    End If
    
    
    If Len(sRemotePath) = 0 Then
        sRemotePath = "\"
    End If
    
    
    'If the local path was blank. Pass the current working direcory.
    If Len(sLocalPath) = 0 Then
        sLocalPath = oFTPScriptShell.CurrentDirectory
    End If
    
    
    If Not oFTPScriptFSO.FolderExists(sLocalPath) Then
        'destination not found
        FailedReason = "Error: Local Folder Not Found."
        GoTo Exit_FTPDownload
    End If
    
    
    Dim sOriginalWorkingDirectory As String
    sOriginalWorkingDirectory = oFTPScriptShell.CurrentDirectory
    oFTPScriptShell.CurrentDirectory = sLocalPath
    '--------END Path Checks---------
    
    'build input file for ftp command
    Dim sFTPScript As String
    sFTPScript = ""
    
    sFTPScript = sFTPScript & "USER " & sUsername & vbCrLf
    sFTPScript = sFTPScript & sPassword & vbCrLf
    sFTPScript = sFTPScript & "cd " & sRemotePath & vbCrLf
    sFTPScript = sFTPScript & "binary" & vbCrLf
    sFTPScript = sFTPScript & "prompt n" & vbCrLf
    sFTPScript = sFTPScript & "mget " & sRemoteFile & vbCrLf
    sFTPScript = sFTPScript & "quit" & vbCrLf & "quit" & vbCrLf & "quit" & vbCrLf
    
    
    Dim sFTPTemp As String
    Dim sFTPTempFile As String
    Dim sFTPResults As String
    
    sFTPTemp = oFTPScriptShell.ExpandEnvironmentStrings("%TEMP%")
    sFTPTempFile = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
    sFTPResults = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
    
    'Write the input file for the ftp command to a temporary file.
    Dim fFTPScript As Object
    Set fFTPScript = oFTPScriptFSO.CreateTextFile(sFTPTempFile, True)
    
    fFTPScript.WriteLine (sFTPScript)
    fFTPScript.Close
    
    Set fFTPScript = Nothing
    
    
    oFTPScriptShell.Run "%comspec% /c FTP -n -s:" & sFTPTempFile & " " & Site & " > " & sFTPResults, 0, True
    Sleep Delay

    
    'Check results of transfer.
    Dim fFTPResults As Object
    Dim sResults As String
    
    Const OpenAsDefault = -2
    Const FailIfNotExist = 0
    Const ForReading = 1
    Const ForWriting = 2
    
    Set fFTPResults = oFTPScriptFSO.OpenTextFile(sFTPResults, ForReading, FailIfNotExist, OpenAsDefault)
    sResults = fFTPResults.ReadAll
    fFTPResults.Close
    
    
    If InStr(sResults, "226 Transfer complete.") > 0 Then
        FailedReason = ""
    ElseIf InStr(sResults, "File not found") > 0 Then
        FailedReason = "Error: File Not Found"
    ElseIf InStr(sResults, "cannot log in.") > 0 Then
        FailedReason = "Error: Login Failed."
    Else
        FailedReason = "Error: Unknown."
    End If
    
    
    oFTPScriptFSO.DeleteFile (sFTPTempFile)
    oFTPScriptFSO.DeleteFile (sFTPResults)
    
    Set oFTPScriptFSO = Nothing
    
    oFTPScriptShell.CurrentDirectory = sOriginalWorkingDirectory
    Set oFTPScriptShell = Nothing
    
    
Exit_FTPDownload:
    FTPDownload = FailedReason
    Exit Function
    
Err_FTPDownload:
    FailedReason = Err.Description
    Resume Exit_FTPDownload
    
End Function



'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function CountRowsInText(sFileName As String) As Long
    'Count Row Number of a text file
    On Error GoTo Err_CountRowsInText
    Dim fso As Object: Set fso = CreatObject("Scripting.FileSystemObject")
    Dim File As Object: Set File = fso.OpenTextFile(sFileName, 1)
    Dim sLine As String, RowCnt As Long: RowCnt = 0
    Do Until File.AtEndOfStream = True
        RowCnt = RowCnt + 1
        sLine = File.ReadLine
    Loop
    File.Close
Exit_CountRowsInText:
    CountRowsInText = RowCnt
    Exit Function
Err_CountRowsInText:
    RowCnt = -1
    Call ShowMsgBox(Err.Description)
    Resume Exit_CountRowsInText
End Function
Public Function SplitTextFile(src As String, Optional des_fmt As String, Optional RowCntPerFile As Long = 65535, Optional file_idx_start As Integer = 0, Optional NumOfHdrRows As Long = 0, Optional DeleteSrc As Boolean = False) As String
    'Split a Text File into multiple text files of specified row count(default: 65535)
    On Error GoTo Err_SplitTextFile
    Dim FailedReason As String, NumOfSplit As Long
    If Len(Dir(src)) = 0 Then FailedReason = src: GoTo Exit_SplitTextFile
    If RowCntPerFile < NumOfHdrRows + 1 Then FailedReason = "RowCntPerFile < NumOfHdrRows + 1": GoTo Exit_SplitTextFile
    Dim RowCnt_src As Long: RowCnt_src = CountRowsInText(src)           'if no need to split, return
    If RowCnt_src <= RowCntPerFile Then GoTo Exit_SplitTextFile
    
    'Check whether there exists files which name is same to the splitted files
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim des_dir As String, des_name As String
    Dim des_ext As String, des_path As String
    des_dir = fso.GetParentFolderName(src)
    des_name = fso.GetFileName(src)
    des_ext = fso.GetExtensionName(src)
    If des_fmt = vbNullString Then des_fmt = Left(des_name, Len(des_name) - Len("." & des_ext)) & "_*"
    NumOfSplit = IIf(RowCnt_src <= RowCntPerFile, 0, Int((RowCnt_src - RowCntPerFile) / (RowCntPerFile + 1 - NumOfHdrRows)) + 1)
    

    Dim file_idx As Integer, file_idx_end As Integer
    file_idx_end = file_idx_start + NumOfSplit 'Int(RowCnt_src / (RowCntPerFile + 1 - NumOfHdrRows))
    For file_idx = file_idx_start To file_idx_end
        des_path = des_dir & "\" & Replace(des_fmt, "*", str(file_idx)) & "." & des_ext
        If Len(Dir(des_path)) > 0 Then Exit For
    Next file_idx
    If Len(Dir(des_path)) > 0 Then FailedReason = des_path: GoTo Exit_SplitTextFile    
    
    'Obtain header rows for later files and create the first splitted file
    Dim FileNum_des As Integer, str_line As String, HdrRows As String
    Dim File_src As Object: Set File_src = fso.OpenTextFile(src, 1)
    des_path = des_dir & "\" & Replace(des_fmt, "*", str(file_idx_start)) & "." & des_ext
    FileNum_des = FreeFile: RowCnt = 0
    Open des_path For Output As #FileNum_des
    Do Until RowCnt >= NumOfHdrRows Or File_src.AtEndOfStream = True
        RowCnt = RowCnt + 1
        str_line = File_src.ReadLine
        Print #FileNum_des, str_line
        HdrRows = HdrRows & str_line
    Loop
    Do Until RowCnt >= RowCntPerFile Or File_src.AtEndOfStream = True
        RowCnt = RowCnt + 1
        Print #FileNum_des, File_src.ReadLine
    Loop
    Close #FileNum_des

    'Start to split
    For file_idx = file_idx_start + 1 To file_idx_end
        If File_src.AtEndOfStream = True Then Exit For
        des_path = des_dir & "\" & Replace(des_fmt, "*", str(file_idx)) & "." & des_ext
        FileNum_des = FreeFile: RowCnt = NumOfHdrRows
        Open des_path For Output As #FileNum_des
        Print #FileNum_des, HdrRows
        Do Until RowCnt >= RowCntPerFile Or File_src.AtEndOfStream = True
            RowCnt = RowCnt + 1
            Print #FileNum_des, File_src.ReadLine
        Loop
        Close #FileNum_des
    Next file_idx
    File_src.Close
    If DeleteSrc = True Then Kill src  
Exit_SplitTextFile:
    SplitTextFile = FailedReason
    Exit Function
Err_SplitTextFile:
    FailedReason = Err.Description
    Resume Exit_SplitTextFile
End Function
Public Function DeleteRowInText(file_name As String, StartRow As Long, EndRow As Long) As String
    'Delete rows in a text file
    On Error GoTo Err_DeleteRowInText
    Dim FailedReason As String, str_line As String, row As Long
    Dim temp_file_name As String, temp_file_PortNum As Integer
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim File As Object: Set File = fso.OpenTextFile(file_name, 1)

    If EndRow < StartRow Then EndRow = StartRow
    temp_file_name = file_name & "_temp"
    On Error Resume Next
    Kill temp_file_name
    On Error GoTo Err_DeleteRowInText
    temp_file_PortNum = FreeFile: row = 0
    Open temp_file_name For Output As #temp_file_PortNum

    Do Until File.AtEndOfStream = True 'EOF(2)
        row = row + 1: str_line = File.ReadLine
        If row >= StartRow And row <= EndRow Then GoTo Loop_DeleteRowInText_1
        Print #temp_file_PortNum, str_line
        
Loop_DeleteRowInText_1:
    Loop
    File.Close
    Close #temp_file_PortNum
    Kill file_name
    Name temp_file_name As file_name
    

Exit_DeleteRowInText:
    DeleteRowInText = FailedReason
    Exit Function

Err_DeleteRowInText:
    FailedReason = Err.Description
    GoTo Exit_DeleteRowInText
End Function



'====================================================================================================
' [FTP] 
'----------------------------------------------------------------------------------------------------
Function ReplaceStrInFolder(folder_name As String, Arr_f As Variant, Arr_r As Variant, Optional StartRow As Long = 0) As String
    'Replace multiple strings in multiple files in a folder
    On Error GoTo Err_ReplaceStrInFolder
    Dim FailedReason As String, file_name As String
    file_name = Dir(folder_name & "\")
        
    Do Until file_name = vbNullString
        file_name = folder_name & "\" & file_name
        Call ReplaceStrInFile(file_name, Arr_f, Arr_r, StartRow)
        file_name = Dir()
    Loop

Exit_ReplaceStrInFolder:
    ReplaceStrInFolder = FailedReason
    Exit Function
    
Err_ReplaceStrInFolder:
    FailedReason = Err.Description
    GoTo Exit_ReplaceStrInFolder
    
End Function
Function ReplaceStrInFile(file_name As String, Arr_f As Variant, Arr_r As Variant, Optional StartRow As Long = 0) As String
    'Replace multiple strings in a file
    On Error GoTo Err_ReplaceStrInFile
    Dim FailedReason As String, iFileNum As String, row As Long
    Dim temp_file_name As String, temp_file_PortNum As Integer
    Dim str_f As String, str_r As String, i As Integer
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim File As Object: Set File = fso.OpenTextFile(file_name, 1)

    temp_file_name = file_name & "_temp"
    On Error Resume Next
    Kill temp_file_name
    On Error GoTo Err_ReplaceStrInFile
    iFileNum = FreeFile(): row = 0
    Open temp_file_name For Output As #iFileNum    
    
    Do Until File.AtEndOfStream = True 'EOF(2)
        row = row + 1: str_line = File.ReadLine
        If row < StartRow Then GoTo Loop_ReplaceStrInFile_1
        For i = 0 To UBound(Arr_f)
            str_f = Arr_f(i)
            str_r = Arr_r(i)
            str_line = Replace(str_line, str_f, str_r)
        Next i

Loop_ReplaceStrInFile_1:
    Print #iFileNum, str_line
    Loop
    File.Close
    Close iFileNum
    Kill file_name
    Name temp_file_name As file_name

Exit_ReplaceStrInFile:
    ReplaceStrInFile = FailedReason
    Exit Function
    
Err_ReplaceStrInFile:
    FailedReason = Err.Description
    GoTo Exit_ReplaceStrInFile
    
End Function




