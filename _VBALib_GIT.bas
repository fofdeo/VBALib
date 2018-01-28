Attribute VB_Name = "_VBALib_Git"
Option Explicit





Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long





Public Function CommitToGIT(OutputDir As String) As Boolean
    'Commits changes in all files to GIT Repo (Dir must contain GIT Repo). Returns True if successful, false o.w.
    Dim CommitMessage As String, ProcessID As Long: On Error GoTo IsError
        Do
            CommitMessage = InputBox("Enter GIT commit input message: ", "GIT Revisions Message")
        Loop Until CommitMessage <> vbNullString
    Open OutputDir & "GITbat.bat" For Output As #1
        Print #1, "cd " & OutputDir     'change directory to tracking folder
        Print #1, "git add -f . && git commit -a -m " & Chr(34) & CommitMessage & Chr(34)
        'add files to staging area, wait until completes succesfully, commit changes into tool
    Close #1
    ChDir OutputDir
    ProcessID = Shell("GITbat.bat > GITout.txt", vbNormalFocus)
    Do While IsProcessOpen(ProcessID) = True
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    OpenAnyFileUsingDefaultProgram (OutputDir & "GITout.txt")
    CommitToGIT = True: Exit Function
IsError:
    CommitToGIT = False
End Function
Private Sub OpenAnyFileUsingDefaultProgram(FullFileName As String)
    ShellExecute 0, "open", FullFileName, 0, 0, 1
End Sub
Private Function IsProcessOpen(PID As Long) As Boolean
    Dim h As Long: h = OpenProcess(&H1, True, PID)
    If h <> 0 Then
        CloseHandle h
      IsProcessOpen = True
    Else
      IsProcessOpen = False
    End If
End Function