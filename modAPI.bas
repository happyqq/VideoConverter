Attribute VB_Name = "modAPI"
''Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal ncmdShow As Long) As Long
''Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMillSecnonds As Long) As Long
''Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAcess As Long, ByVal bInheri As Long, ByVal dwProcessID As Long) As Long
''Global Const INFINITE = -1
''Global Const SYNCHRONIZE = &H100000
''Global iTask As Long, ret As Long, pHandle As Long
'
'
'
' Option Explicit
'
'
'  Type PROCESS_INFORMATION
'    hProcess As Long
'    hThread As Long
'    dwProcessID As Long
'    dwThreadID As Long
'  End Type
'
'
'
'
'  Type SECURITY_ATTRIBUTES
'    nLength   As Long
'    lpSecurityDescriptor   As Long
'    bInheritHandle   As Long
'  End Type
'
'
' Type STARTUPINFO
'    cb   As Long
'    lpReserved   As Long
'    lpDesktop   As Long
'    lpTitle   As Long
'    dwX   As Long
'    dwY   As Long
'    dwXSize   As Long
'    dwYSize   As Long
'    dwXCountChars   As Long
'    dwYCountChars   As Long
'    dwFillAttribute   As Long
'    dwFlags   As Long
'    wShowWindow   As Integer
'    cbReserved2   As Integer
'    lpReserved2   As Byte
'    hStdInput   As Long
'    hStdOutput   As Long
'    hStdError   As Long
'End Type
'
'Type OVERLAPPED
'        ternal   As Long
'        ternalHigh   As Long
'        offset   As Long
'        OffsetHigh   As Long
'        hEvent   As Long
'End Type
'
'Global Const STARTF_USESHOWWINDOW = &H1
'Global Const STARTF_USESTDHANDLES = &H100
'Global Const SW_HIDE = 0
'Global Const EM_SETSEL = &HB1
'Global Const EM_REPLACESEL = &HC2
'
'
'  Global Const NORMAL_PRIORITY_CLASS = &H20&
'  Global Const INFINITE = -1&
'
'  Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean
'  Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'  Declare Function CreateProcessA Lib "kernel32" ( _
'  ByVal lpApplicationName As Long, _
'  ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal _
'  lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal _
'  dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal _
'  lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, _
'  lpProcessInformation As PROCESS_INFORMATION) As Long
'
'
'  Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
'
'
'  '==============================================================
'
''Redirects   output   from   console   program   to   textbox.
''Requires   two   textboxes   and   one   command   button.
''Set   MultiLine   property   of   Text2   to   true.
''
''Original   bcx   version   of   this   program   was   made   by
''   dl   <dl@tks.cjb.net>
''VB   port   was   made   by   Jernej   Simoncic   <jernej@isg.si>
''Visit   Jernejs   site   at   http://www2.arnes.si/~sopjsimo/
''
''Note:   don 't   run   plain   DOS   programs   with   this   example
''under   Windows   95,98   and   ME,   as   the   program   freezes   when
''execution   of   program   is   finnished.
'
'
'Private Declare Function CreatePipe Lib "kernel32 " (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
'Private Declare Sub GetStartupInfo Lib "kernel32 " Alias "GetStartupInfoA " (lpStartupInfo As STARTUPINFO)
'Private Declare Function CreateProcess Lib "kernel32 " Alias "CreateProcessA " (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'Private Declare Function SetWindowText Lib "user32 " Alias "SetWindowTextA " (ByVal hWnd As Long, ByVal lpString As String) As Long
'Private Declare Function ReadFile Lib "kernel32 " (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
'Private Declare Function SendMessage Lib "user32 " Alias "SendMessageA " (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'
'
'
'
'Declare Function GetProcessVersion Lib "kernel32" (ByVal ProcessId As Long) As Long
'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
'Public Function ShellEx(ByVal FileName As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal DelayTime As Long = -1) As Long
'    '��SHELL����һ���Ĳ���,����������ִ��.(ͬ��)
'    'FileName - Ŀ���ļ���
'    'WindowStyle - ��������ʱ���ڵ���ʽ
'    'DelayTime - �ȴ���ʱ��,��λΪms
'    '��ע:
'    '       DelayTime����Ϊ-1ʱ��ʾһֱ�ȴ�,ֱ��Ŀ��������н���
'    Dim i As Long, J As Long
'
'    i = Shell(FileName, WindowStyle)
'    ShellEx = i
'    Do
'        If GetProcessVersion(i) = 0 Then Exit Do            'Ŀ������˳�ʱ����
'        Sleep 10
'        J = J + 1
'        DoEvents
'        If DelayTime <> -1 And J > DelayTime \ 10 Then Exit Do  '����Զ�ȴ�+�ȴ�ʱ��ﵽʱ����
'    Loop
'End Function
'
'
'
'
'
'
'
'
'
'Public Sub Redirect(cmdLine As String, objTarget As Object)
'    Dim i%, t$
'    Dim pa     As SECURITY_ATTRIBUTES
'    Dim pra     As SECURITY_ATTRIBUTES
'    Dim tra     As SECURITY_ATTRIBUTES
'    Dim PI     As PROCESS_INFORMATION
'    Dim sui     As STARTUPINFO
'    Dim hRead     As Long
'    Dim hWrite     As Long
'    Dim bRead     As Long
'    Dim lpBuffer(1024)     As Byte
'    pa.nLength = Len(pa)
'    pa.lpSecurityDescriptor = 0
'    pa.bInheritHandle = True
'
'    pra.nLength = Len(pra)
'    tra.nLength = Len(tra)
'
'    If CreatePipe(hRead, hWrite, pa, 0) <> 0 Then
'        sui.cb = Len(sui)
'        GetStartupInfo sui
'        sui.hStdOutput = hWrite
'        sui.hStdError = hWrite
'        sui.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
'        sui.wShowWindow = SW_HIDE
'        If CreateProcess(vbNullString, cmdLine, pra, tra, True, 0, Null, vbNullString, sui, PI) <> 0 Then
'            SetWindowText objTarget.hWnd, " "
'            Do
'                Erase lpBuffer()
'                If ReadFile(hRead, lpBuffer(0), 1023, bRead, ByVal 0&) Then
'                    SendMessage objTarget.hWnd, EM_SETSEL, -1, 0
'                    SendMessage objTarget.hWnd, EM_REPLACESEL, False, lpBuffer(0)
'                    DoEvents
'                Else
'                    CloseHandle PI.hThread
'                    CloseHandle PI.hProcess
'                    Exit Do
'                End If
'                CloseHandle hWrite
'            Loop
'            CloseHandle hRead
'        End If
'    End If
'End Sub
'
'
'
'Public Function IsRunning(ByVal ProgramID) As Boolean ' ������̱�ʶID
'  Dim hProgram As Long '�����ĳ�����̾��
'  hProgram = OpenProcess(0, False, ProgramID)
'  If Not hProgram = 0 Then
'  IsRunning = True
'  Else
'  IsRunning = False
'  End If
'  CloseHandle hProgram
'End Function
'
'
'Public Sub ShellAndWait(cmdLine$)
'
'  Dim NameOfProc As PROCESS_INFORMATION
'  Dim NameStart As STARTUPINFO
'  Dim x As Long
'
'  NameStart.cb = Len(NameStart)
'  NameStart.wShowWindow = SW_HIDE
'  x = CreateProcessA(0&, cmdLine$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, _
'  0&, 0&, NameStart, NameOfProc)
'  x = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
'  x = CloseHandle(NameOfProc.hProcess)
'
'End Sub
''
''  ����һ�����壬����һ�����ť(Command1)�����ϡ��� Command1_Click �¼��������������ݣ�
''  Private Sub Command1_Click()
''  Dim AppToLaunch As String
''
''  AppToLaunch = "c:\win95\notepad.exe"
''  ShellAndWait AppToLaunch
''  End Sub
''
'
