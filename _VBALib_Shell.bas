'====================================================================================================
Attribute VB_Name = "_VBALib_Shell"
Option Compare Database
Option Explicit

Const INFINITE = &HFFFF
Const SYNCHRONIZE = &H100000
'====================================================================================================
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GlobalFindAtomA Lib "kernel32" (ByVal lpstr As String) As Integer
Declare Function GlobalAddAtomA Lib "kernel32" (ByVal lpstr As String) As Integer
Declare Function GlobalFindAtomW Lib "kernel32" (ByVal lpstr As String) As Integer
Declare Function GlobalAddAtomW Lib "kernel32" (ByVal lpstr As String) As Integer
Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'====================================================================================================





'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Function ShellAndWait(ByVal cmd As String, ByVal window_style As VbAppWinStyle) As Boolean
    'Start a Shell command and wait for it to finish, hiding while it is running.
    Dim process_id As Long, process_handle As Long
    On Error GoTo ShellError    'Start the program
    ShellAndWait = False
    process_id = Shell(cmd, window_style)
    On Error GoTo 0  'Wait for Program finish w/ Process handle
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
        ShellAndWait = True
    End If
    Exit Function
ShellError:
    'MsgBox "Error starting task " & txtProgram.text & vbCrLf & Err.Description, vbOKOnly Or vbExclamation, "Error"
    ShellAndWait = False
End Function

Public Function Shell_SendKeysWithTimeout(oshell As Object, CmdTxt As String, Timeout As Integer) As String
    'Send multiples shell commands with timeout
    On Error GoTo Err_Shell_SendKeysWithTimeout: Dim FailedReason As String, cmd_idx As Integer, CmdSet As Variant
    CmdSet = SplitStrIntoArray(CmdTxt, Chr(10))
    For cmd_idx = 0 To UBound(CmdSet)
        If CmdSet(cmd_idx) = vbNullString Then GoTo Next_Shell_SendKeysWithTimeout
        With oshell
            .SendKeys (CmdSet(cmd_idx) & vbCrLf)
            Sleep Timeout
        End With 'oShell
Next_Shell_SendKeysWithTimeout:
    Next cmd_idx
Exit_Shell_SendKeysWithTimeout:
    Shell_SendKeysWithTimeout = FailedReason
    Exit Function
Err_Shell_SendKeysWithTimeout:
    FailedReason = Err.Description
    Resume Exit_Shell_SendKeysWithTimeout
End Function





'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Sub RunAndKill()
    Dim ProcessID As Long, TimeOut As Double, TimeOutSec As Integer
	Dim EXE_Dir as String, EXE_Name as String, RunSuccessful as Boolean
	EXE_Dir = "C:\USEPA\BMDS212\": EXE_Name = "BMDS2.exe"
	TimeOutSec = 30: TimeOut = Now + TimeSerial(0,0,TimeOutSec): ChDir EXE_Dir
    ProcessID = Shell(EXE_Name & " " & EXE_Dir & "\INPUTFILE.INP", vbHide): RunSuccessful=True
    Do While IsProcessOpen(ProcessID) = True
        Application.StatusBar = " Waiting for " & EXE_Name & " to complete... " & Now
        Application.Wait Now + TimeSerial(0, 0, 1) 'Chk if Input File was Deleted, then Chk if Time is exceeded
		If z_Files.GetFileInfo(Range("IEUBK_AutoItFN"), FileExists) = False Then: RunSuccessful = True: Exit Do
		If Now > TimeOut Then: RunSuccessful= False: Shell "TASKKILL /F /PID " & ProcessID, vbHide: Exit Do
    Loop
	Application.ScreenUpdating = False
	If RunSuccessful = TRUE Then
		Debug.Print EXE_Name & " completed succesfully"
	Else
		Debug.Print EXE_Name & " timed out"
	End If
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




'====================================================================================================
' [] 
'----------------------------------------------------------------------------------------------------
Public Sub OpenAnyFile(FullFileName as String)
	ShellExecute 0, "open", FullFileName, 0, 0, 1
End Sub