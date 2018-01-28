Attribute VB_Name = "_VBALIB_Folder"
Option Explicit






Public Sub ListFiles(ByVal strPath As String, ByVal cellDestination As Range)
    'List the files in the folder specified in argument and display the list of the files in the cells below the cell 
    '    given in argument as well as the full path in the right column        (Has to be launched by an other macro)
    'strPath as String : the full path of the folder    &&     &&    'cellDestination as Range : the destination cell
    Dim counter As Integer, File As String              'Dim filesTab
    strPath = checkFolder(strPath)    'Add a trailing slash if needed
    File = Dir$(strPath & Extention)  'Count number of files in folder
    Do While Len(File)
        File = Dir$
        counter = counter + 1
    Loop
    If (counter = 0) Then Exit Sub
    ReDim filesTab(counter - 1)
    counter = 0 'reset Array counter
    File = Dir$(strPath & Extention)
    ' List the files and display them in the cells
    Do While Len(File) And counter <= UBound(filesTab)
        cellDestination.Offset(counter, 0) = File
        cellDestination.Offset(counter, 1) = strPath & File
        File = Dir$
        counter = counter + 1
    Loop
End Sub


Function CheckFolder(ByVal strPath As String) As String
    'Check Folder has a trailing slash and add one if needed. Create the folders if needed. strPath as String : the full path of the folder
    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    strPath = Replace(strPath, "/", "\")
    strPath = Replace(strPath, "\\", "\")
    ' createDirs (strPath)
    CheckFolder = strPath
End Function










Function FolderExists(ByVal fullPath As String) As String
    Dim fs: Set fs = CreateObject("Scripting.FileSystemObject")
    FolderExists = fs.folderExists(fullPath)
End Function
Function CreateDirs(ByVal fullPath As String) As String
    fullPath = checkFolder(fullPath): folderCreated = 0
    paths = Split(fullPath, "\"): currentPath = paths(0) & "\"
    For i = 1 To UBound(paths) - 1
        currentPath = currentPath & paths(i) & "\"
        If folderExists(currentPath) = False Then
            createFolder (currentPath)
            folderCreated = folderCreated + 1
        End If
    Next
    createDirs = folderCreated & " folder(s) has/have been generated"
End Function
Function CreateFolder(ByVal fullPath As String) As Boolean
    If (FolderExists(fullPath) = False) Then MkDir (fullPath)    
End Function



'---------------------------------------------------------
Sub ListFolder(sFolderPath As String, ByVal cellDestination As Range)
    'List Folders in Folder & display the List in Cells below given
    'strPath as String : the full path of the folder
    'cellDestination as Range : the destination cell 
    Dim fs As New FileSystemObject, i As Integer: i = 0
    Dim FSfolder As Folder, subfolder As Folder     
    Set FSfolder = fs.GetFolder(sFolderPath)
    For Each subfolder In FSfolder.SubFolders
        DoEvents: i = i + 1
        cellDestination.Offset(i, 0) = subfolder.Name
    Next subfolder
    Set FSfolder = Nothing     
End Sub

 
Sub TestListFolders()
    Application.ScreenUpdating = False
    Cells.Delete         ' add headers
    With Range("A1")
        .Formula = "Folder contents:"
        .Font.Bold = True
        .Font.Size = 12
    End With
    Range("A3").Formula = "Folder Path:"
    Range("B3").Formula = "Folder Name:"
    Range("C3").Formula = "Size:"
    Range("D3").Formula = "Subfolders:"
    Range("E3").Formula = "Files:"
    Range("F3").Formula = "Short Name:"
    Range("G3").Formula = "Short Path:"
    Range("A3:G3").Font.Bold = True
     'ENTER START FOLDER HERE and include subfolders (true/false)
    listFoldersFullInfo "H:\User\02. Projects\", False
    Application.ScreenUpdating = True
End Sub
 
Sub listFoldersFullInfo(SourceFolderName As String, IncludeSubfolders As Boolean)
    ' lists information about the folders in SourceFolder
    Dim r As Long, FSO As Scripting.FileSystemObject: Set FSO = New Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder, subfolder As Scripting.Folder 
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    On Error Resume Next    'Display Folder properties
    r = Range("A65536").End(xlUp).Row + 1
    cells(r, 1).Formula = SourceFolder.Path
    cells(r, 2).Formula = SourceFolder.Name
    cells(r, 3).Formula = SourceFolder.Size
    cells(r, 4).Formula = SourceFolder.SubFolders.Count
    cells(r, 5).Formula = SourceFolder.Files.Count
    cells(r, 6).Formula = SourceFolder.ShortName
    cells(r, 7).Formula = SourceFolder.ShortPath
    If IncludeSubfolders Then
        For Each subfolder In SourceFolder.SubFolders
            listFolders subfolder.Path, True
        Next subfolder
        Set subfolder = Nothing
    End If
    Columns("A:G").AutoFit
    Set SourceFolder = Nothing
    Set FSO = Nothing
End Sub



Function GetOldestFileInDir(ByVal path As String, ByVal fileNameMask As String) As Date
    Dim FileName As String, FileDir As String, FileSearch As String
    Dim MaxDate As Date, interDate As Date, dteFile As Date
    MaxDate = DateSerial(1900, 1, 1)
    FileDir = path: FileName = fileNameMask
    FileSearch = Dir(FileDir & FileName)
    While Len(FileSearch) > 0
        dteFile = FileDateTime(FileDir & FileSearch)
        If dteFile > MaxDate Then MaxDate = dteFile
        FileSearch = Dir()
    Wend
    GetOldestFileInDir = MaxDate
End Function