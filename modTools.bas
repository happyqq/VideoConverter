Attribute VB_Name = "modTools"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias _
                   "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) _
                   As Long
                   
Private Declare Function GetCurrentProcess Lib "kernel32" () _
        As Long
        
Private Declare Function OpenProcessToken Lib "advapi32" ( _
        ByVal ProcessHandle As Long, ByVal DesiredAccess As _
        Long, TokenHandle As Long) As Long
                                                       
Private Declare Function ExitWindowsEx Lib "user32" (ByVal _
        dwOptions As Long, ByVal dwReserved As Long) As Long
        
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
        Alias "LookupPrivilegeValueA" (ByVal lpSystemName As _
        String, ByVal lpName As String, lpLuid As LUID) As Long
        
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
        (ByVal TokenHandle As Long, ByVal DisableAllPrivileges _
        As Long, NewState As TOKEN_PRIVILEGES, ByVal _
        BufferLength As Long, PreviousState As _
        TOKEN_PRIVILEGES, ReturnLength As Long) As Long
        
Private Declare Function CloseHandle Lib "kernel32" (ByVal _
        hObject As Long) As Long
        
        
Public Enum ShutDownActionsEnum
  [saShutdown] = 1&
  [saReboot] = 2&
  [saLogOff] = 4&
  [saPowerOff] = 8&
  [saForceIfHung] = 16&
End Enum

Private Type OSVERSIONINFO
             dwOSVersionInfoSize As Long
             dwMajorVersion As Long
             dwMinorVersion As Long
             dwBuildNumber As Long
             dwPlatformId As Long
             szCSDVersion As String * 128
End Type

Private Type LUID
  UsedPart As Long
  IgnoredForNowHigh32BitPart As Long
End Type


Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  TheLuid As LUID
  Attributes As Long
End Type

Public Const VER_PLATFORM_WIN32s As Long = 0
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2
Public gbExplicitEnd As Boolean

Public gsWinVersion As String


Private Function IsWinNT() As Boolean
    
    Dim tOSVERSIONINFO As OSVERSIONINFO
    
    tOSVERSIONINFO.dwOSVersionInfoSize = Len(tOSVERSIONINFO)
    
    Call GetVersionEx(tOSVERSIONINFO)
    
    IsWinNT = CBool(tOSVERSIONINFO.dwPlatformId And VER_PLATFORM_WIN32_NT)
    
End Function

Public Sub ShutDownWin()

    Dim ShutdownFlags As Long
    
    gbExplicitEnd = True
    
    ShutdownFlags = [saShutdown] Or [saPowerOff]
    
    If IsWinNT() Then    'Priviligien setzen
        Call SetShutdownPrivilege
    End If
    
    If gsWinVersion = "Windows 2000" Or gsWinVersion = "Windows XP" Then
        ShutdownFlags = ShutdownFlags Or [saForceIfHung]
    Else
        ShutdownFlags = ShutdownFlags Or [saLogOff]
    End If
    
    Call ExitWindowsEx(ShutdownFlags, &HFFFF)
    
End Sub

Private Sub SetShutdownPrivilege()

    Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
    Const TOKEN_QUERY As Long = &H8
    Const SE_PRIVILEGE_ENABLED As Long = &H2
    
    Dim hProcessHandle As Long
    Dim hTokenHandle As Long
    Dim PrivLUID As LUID
    Dim TokenPriv As TOKEN_PRIVILEGES
    Dim tkpDummy As TOKEN_PRIVILEGES
    Dim lDummy As Long
    
    'Ermittlung eines Prozess-Handles dieser Anwendung
    hProcessHandle = GetCurrentProcess()
    
    'Für unseren Prozess soll ein Token geändert werden.
    OpenProcessToken hProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
        TOKEN_QUERY), hTokenHandle
    
    'Die repräsentierende LUID für das "SeShutdownPrivilege" ermitteln
    Call LookupPrivilegeValue("", "SeShutdownPrivilege", PrivLUID)
    
    'Vorbereitungen auf das Ändern des Tokens
    With TokenPriv
        'Anzahl der Privilegien
        .PrivilegeCount = 1
        
        'LUID-Struktur für das Privileg
        .TheLuid = PrivLUID
        
        'Das Privileg soll gesetzt werden
        .Attributes = SE_PRIVILEGE_ENABLED
    End With
    
    'Jetzt wird das Token für diesen Prozess gesetzt, um
    'unserem Prozess das Recht für ein Herunterfahren / einen
    'Neustart zuzuteilen:
    Call AdjustTokenPrivileges(hTokenHandle, False, TokenPriv, _
        Len(tkpDummy), tkpDummy, lDummy)
    
    'Handle auf das geoeffnete Token freigeben
    Call CloseHandle(hTokenHandle)
    
End Sub


Public Function ReadFromFile(ByVal sFileName As String, Optional ByVal bBinary As Boolean = False) As String
    
    On Error Resume Next
    Dim iFileNr As Integer
    Dim sTmp As String
    Dim b() As Byte
    
    If Dir(sFileName) = "" Then Exit Function
    
    iFileNr = FreeFile()
    Open sFileName For Binary Access Read As #iFileNr
        If bBinary Then
            ReDim b(0 To FileLen(sFileName) - 1) As Byte
            
            Get #iFileNr, , b()
            ReadFromFile = b()
        Else
            sTmp = Space$(FileLen(sFileName))
            Get #iFileNr, , sTmp
            ReadFromFile = sTmp
        End If

errExit:
    Close #iFileNr
    
End Function
