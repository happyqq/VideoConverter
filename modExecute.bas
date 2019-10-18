Attribute VB_Name = "modExecute"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CreateProssessStdOut Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const INFINITE = -1&

Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100&

Public Enum Execute_WindowMode
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_FORCEMINIMIZE = 11
End Enum

Public Enum Execute_ProcessPriority
   IDLE_PRIORITY_CLASS = &H40
   BELOW_NORMAL_PRIORITY_CLASS = &H4000
   NORMAL_PRIORITY_CLASS = &H20
   ABOVE_NORMAL_PRIORITY_CLASS = &H8000
   HIGH_PRIORITY_CLASS = &H80
   REALTIME_PRIORITY_CLASS = &H100
End Enum

Public Function Execute(ByVal CommandLine As String, _
                               Optional ByVal WaitMS As Long = INFINITE, _
                               Optional ByVal ProcessPriority As Execute_ProcessPriority = NORMAL_PRIORITY_CLASS, _
                               Optional ByVal WindowMode As Execute_WindowMode = SW_HIDE, _
                               Optional ByVal Directory As String = vbNullString) As Boolean
On Local Error GoTo EXECUTE_ERROR
Dim SI As STARTUPINFO
Dim PI As PROCESS_INFORMATION

Dim res As Long

    SI.cb = Len(SI)
    SI.wShowWindow = WindowMode
    SI.dwFlags = STARTF_USESHOWWINDOW

    'create the process. returns nonzero on success
    res = CreateProcess(vbNullString, CommandLine, 0&, 0&, 1&, ProcessPriority, 0&, Directory, SI, PI)
    If (WaitMS <> 0 And res <> 0) Then
        WaitForSingleObject PI.hProcess, WaitMS
    End If
    
    If (res = 0) Then
        'failed to create process
        Exit Function
    End If

    CloseHandle PI.hProcess
    CloseHandle PI.hThread
    Execute = True
    Exit Function
    
EXECUTE_ERROR:
End Function

Public Function ExecuteGetStdOut(ByVal CommandLine As String, _
                             TargetFilePath As String, _
                             Optional ByVal ProcessPriority As Execute_ProcessPriority = NORMAL_PRIORITY_CLASS, _
                             Optional ByVal WindowMode As Execute_WindowMode = SW_HIDE, _
                             Optional ByVal Directory As String = vbNullString) As Boolean
On Local Error GoTo EXECUTE_ERROR
Dim SI As STARTUPINFO
Dim SA As SECURITY_ATTRIBUTES
Dim PI As PROCESS_INFORMATION

Dim hReadPipe As Long               'Read Pipe handle created by CreatePipe
Dim hWritePipe As Long              'Write Pite handle created by CreatePipe
Dim lngBytesread As Long            'Amount of byte read from the Read Pipe handle
Dim strBuff As String * 256         'String buffer reading the Pipe
Dim hFile As Long

Dim res As Long

    'create the Pipe
    SA.nLength = Len(SA)
    SA.bInheritHandle = 1&
    SA.lpSecurityDescriptor = 0&
    res = CreatePipe(hReadPipe, hWritePipe, SA, 0)
    
    If (res = 0) Then
        Exit Function
    End If

    SI.cb = Len(SI)
    SI.wShowWindow = WindowMode
    SI.hStdOutput = hWritePipe
    SI.hStdError = hWritePipe
    SI.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW

    'create the process. returns nonzero on success
    res = CreateProssessStdOut(vbNullString, CommandLine, SA, SA, 1&, ProcessPriority, 0&, Directory, SI, PI)
    If (res = 0) Then
        'failed to create process
        Exit Function
    End If
    
    'close the hWritePipe
    res = CloseHandle(hWritePipe)
   
    'write StdOut to file
    hFile = FreeFile
    Open TargetFilePath For Output As #hFile
    Do
        res = ReadFile(hReadPipe, strBuff, 256, lngBytesread, 0&)
        If (res <> 0) Then
            Print #hFile, Left(strBuff, lngBytesread);
        End If
    Loop While (res <> 0)

    Close #hFile

    CloseHandle PI.hProcess
    CloseHandle PI.hThread
    CloseHandle hReadPipe
        
    ExecuteGetStdOut = True
    Exit Function
    
EXECUTE_ERROR:
End Function



