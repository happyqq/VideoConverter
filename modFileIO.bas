Attribute VB_Name = "modFileIO"
Option Explicit

Private Declare Function GetTempPathA Lib "kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileNameA Lib "kernel32.dll" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean
Private Declare Sub CopyMem Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
      
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH As Long = 260
Private Const UNIQUE_NAME = &H0
Private Const STD_INPUT_HANDLE = -10&

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_INTERNET = &H1
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_MYDOCUMENTS = &HC
Public Const CSIDL_MYMUSIC = &HD
Public Const CSIDL_MYVIDEO = &HE
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_PRINTHOOD = &H1B
Public Const CSIDL_LOCAL_APPDATA = &H1C
Public Const CSIDL_ALTSTARTUP = &H1D
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22
Public Const CSIDL_COMMON_APPDATA = &H23
Public Const CSIDL_WINDOWS = &H24
Public Const CSIDL_SYSTEM = &H25
Public Const CSIDL_PROGRAM_FILES = &H26
Public Const CSIDL_MYPICTURES = &H27
Public Const CSIDL_PROFILE = &H28
Public Const CSIDL_SYSTEMX86 = &H29
Public Const CSIDL_PROGRAM_FILESX86 = &H2A
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B
Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
Public Const CSIDL_COMMON_TEMPLATES = &H2D
Public Const CSIDL_COMMON_DOCUMENTS = &H2E
Public Const CSIDL_COMMON_ADMINTOOLS = &H2F
Public Const CSIDL_ADMINTOOLS = &H30
Public Const CSIDL_CONNECTIONS = &H31
Public Const CSIDL_COMMON_MUSIC = &H35
Public Const CSIDL_COMMON_PICTURES = &H36
Public Const CSIDL_COMMON_VIDEO = &H37
Public Const CSIDL_RESOURCES = &H38
Public Const CSIDL_RESOURCES_LOCALIZED = &H39
Public Const CSIDL_COMMON_OEM_LINKS = &H3A
Public Const CSIDL_CDBURN_AREA = &H3B
Public Const CSIDL_COMPUTERSNEARME = &H3D

Private Type VS_FIXEDFILEINFO
   Signature As Long
   StrucVersionl As Integer     '  e.g. = &h0000 = 0
   StrucVersionh As Integer     '  e.g. = &h0042 = .42
   FileVersionMSl As Integer    '  e.g. = &h0003 = 3
   FileVersionMSh As Integer    '  e.g. = &h0075 = .75
   FileVersionLSl As Integer    '  e.g. = &h0000 = 0
   FileVersionLSh As Integer    '  e.g. = &h0031 = .31
   ProductVersionMSl As Integer '  e.g. = &h0003 = 3
   ProductVersionMSh As Integer '  e.g. = &h0010 = .1
   ProductVersionLSl As Integer '  e.g. = &h0000 = 0
   ProductVersionLSh As Integer '  e.g. = &h0031 = .31
   FileFlagsMask As Long        '  = &h3F for version "0.42"
   FileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   FileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   FileType As Long             '  e.g. VFT_DRIVER
   FileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   FileDateMS As Long           '  e.g. 0
   FileDateLS As Long           '  e.g. 0
End Type

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Function CorrectFilename(ByVal InFile As String) As String
Dim TempFilename As String

    TempFilename = InFile
    If (Len(InFile) > 0) Then
        TempFilename = Replace(TempFilename, "\", "_")
        TempFilename = Replace(TempFilename, "/", "_")
        TempFilename = Replace(TempFilename, ":", "-")
        TempFilename = Replace(TempFilename, "*", "-")
        TempFilename = Replace(TempFilename, "?", "-")
        TempFilename = Replace(TempFilename, """", "'")
        TempFilename = Replace(TempFilename, "<", "-")
        TempFilename = Replace(TempFilename, ">", "-")
        TempFilename = Replace(TempFilename, "|", "-")
        If (LCase(right(TempFilename, 4)) <> ".pdf") Then
            TempFilename = TempFilename & ".pdf"
        End If
    End If
    CorrectFilename = TempFilename
End Function

Public Function StripPath(ByVal vData As String) As String
On Error Resume Next
Dim DotPos As Long
    
    DotPos = InStrRev(vData, "\")
    If DotPos > 0 Then
        StripPath = right(vData, Len(vData) - DotPos)
    Else
        StripPath = vData
    End If
End Function

Public Function StripFilename(ByVal vData As String) As String
On Error Resume Next
Dim DotPos As Long
    
    DotPos = InStrRev(vData, "\")
    If DotPos > 0 Then
        StripFilename = left(vData, DotPos)
    Else
        StripFilename = Empty
    End If
End Function

Public Function StripExtension(ByVal vData As String) As String
On Error Resume Next
Dim DotPos As Long

    DotPos = InStrRev(vData, ".")
    If DotPos > 0 Then
        StripExtension = left(vData, DotPos - 1)
    Else
        StripExtension = vData
    End If
End Function

Public Function GetExtension(ByVal vData As String) As String
On Error Resume Next
Dim DotPos As Long

    DotPos = InStrRev(vData, ".")
    If DotPos > 0 Then
        GetExtension = right(vData, Len(vData) - DotPos)
    End If
End Function

Public Function AddBackslash(ByVal S As String) As String
    If Len(S) > 0 Then
        If right$(S, 1) <> "\" Then
            AddBackslash = S + "\"
        Else
            AddBackslash = S
        End If
    Else
        AddBackslash = Empty
    End If
End Function

Public Function RemoveBackslash(ByVal S As String) As String
    If Len(S) > 0 Then
        If right$(S, 1) = "\" Then
            RemoveBackslash = left(S, Len(S) - 1)
        Else
            RemoveBackslash = S
        End If
    Else
        RemoveBackslash = Empty
    End If
End Function

Public Function VBPathExists(ByVal FolderPath As String) As Boolean
On Error GoTo NoPath:
    Dir FolderPath
    VBPathExists = True
    Exit Function
NoPath:
    VBPathExists = False
End Function

Public Function PathExists(ByVal FolderPath As String) As Boolean
On Local Error Resume Next
Dim FSO     'As Scripting.FileSystemObject

    Set FSO = CreateObject("Scripting.FileSystemObject")
    PathExists = FSO.FolderExists(FolderPath)
    Set FSO = Nothing
End Function

Public Function CreateDirectory(ByVal FolderPath As String) As Boolean
On Local Error Resume Next
Dim FSO     'As Scripting.FileSystemObject
    
    FolderPath = RemoveBackslash(FolderPath)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CreateFolder FolderPath
    CreateDirectory = FSO.FolderExists(FolderPath)
    Set FSO = Nothing
End Function

Public Function DeleteDirectory(ByVal FolderPath As String) As Boolean
On Local Error Resume Next
Dim FSO     'As Scripting.FileSystemObject
    
    FolderPath = RemoveBackslash(FolderPath)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFolder FolderPath, True
    DeleteDirectory = Not FSO.FolderExists(FolderPath)
    Set FSO = Nothing
End Function

Public Function VBFileExists(ByVal FilePath As String) As Boolean
Dim WFD As WIN32_FIND_DATA
Dim hFile As Long
  
    hFile = FindFirstFile(FilePath, WFD)
    VBFileExists = hFile <> INVALID_HANDLE_VALUE
    FindClose (hFile)
End Function

Public Function FileExists(ByVal FilePath As String) As Boolean
On Local Error Resume Next
Dim FSO     'As Scripting.FileSystemObject
  
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FileExists = FSO.FileExists(FilePath)
    Set FSO = Nothing
End Function

Public Function FileSize(FilePath As String) As Long
On Local Error Resume Next
Dim FSO         'As Scripting.FileSystemObject
Dim FSO_File    'As Scripting.file

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FilePath) Then
        Set FSO_File = FSO.GetFile(FilePath)
        FileSize = FSO_File.Size
        Set FSO_File = Nothing
    End If
    Set FSO = Nothing
End Function

Public Function CopyFile(SourcePath As String, TargetPath As String, Optional Overwrite As Boolean = True) As Boolean
On Local Error Resume Next
Dim FSO     'As Scripting.FileSystemObject

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If (FSO.FileExists(TargetPath)) Then
        If (Overwrite = True) Then
            If (DeleteFile(TargetPath)) Then
                'Successfully deleted
                FSO.CopyFile SourcePath, TargetPath
                CopyFile = FSO.FileExists(TargetPath)
            Else
                'Unable to delete file
                CopyFile = False
            End If
        Else
            'File exists and will not be overwritten
            CopyFile = False
        End If
    Else
        'file doesn't exist, just move it
        FSO.CopyFile SourcePath, TargetPath
        CopyFile = FSO.FileExists(TargetPath)
    End If
    Set FSO = Nothing
End Function

Public Function MoveFile(SourcePath As String, TargetPath As String, Optional Overwrite As Boolean = True) As Boolean
On Local Error Resume Next
Dim FSO     'As Scripting.FileSystemObject
  
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If (FSO.FileExists(TargetPath)) Then
        If (Overwrite = True) Then
'            If (FSO.DeleteFile(TargetPath)) Then '这一句老是判断不成功
                'Successfully deleted
                FSO.DeleteFile TargetPath, True
                FSO.MoveFile SourcePath, TargetPath
                MoveFile = FSO.FileExists(TargetPath) '先强制删除，再移入
'            Else
'                'Unable to delete file
'                MoveFile = False
'            End If
        Else
            'File exists and will not be overwritten
            MoveFile = False
        End If
    Else
        'file doesn't exist, just move it
        FSO.MoveFile SourcePath, TargetPath
        MoveFile = FSO.FileExists(TargetPath)
    End If
    Set FSO = Nothing
End Function

Public Function DeleteFile(FilePath As String) As Boolean
On Local Error Resume Next
Dim FSO     'As Scripting.FileSystemObject
  
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFile (FilePath)
    DeleteFile = Not FSO.FileExists(FilePath)
    Set FSO = Nothing
End Function

Public Function GetTempFileName() As String
On Local Error Resume Next
Dim sTmp As String
Dim sTmp2 As String
    
    sTmp2 = GetTempPath
    sTmp = Space(Len(sTmp2) + 256)
    Call GetTempFileNameA(sTmp2, App.EXEName, UNIQUE_NAME, sTmp)
    'Hier wird immer der kurze Dateiname zur點kgegeben!!
    GetTempFileName = GetShortName(left$(sTmp, InStr(sTmp, Chr$(0)) - 1))
    DeleteFile GetTempFileName
End Function

Public Function GetTempPath() As String
On Local Error Resume Next
Dim sTmp As String
Dim i As Integer

    i = GetTempPathA(0, "")
    sTmp = Space(i)
    Call GetTempPathA(i, sTmp)
    GetTempPath = AddBackslash(left$(sTmp, i - 1))
End Function

Public Function GetSystemDir() As String
On Local Error Resume Next
Dim temp As String * 256
Dim X As Integer
    
    X = GetSystemDirectory(temp, Len(temp))
    GetSystemDir = left$(temp, X)
End Function

Public Function GetWinDir() As String
On Local Error Resume Next
Dim temp As String * 256
Dim X As Integer
    
    X = GetWindowsDirectory(temp, Len(temp))
    GetWinDir = left$(temp, X)
End Function

Public Function GetAllUsersDocumentsDir() As String
On Local Error Resume Next
Dim blnReturn As Long
Dim strBuffer As String
   
   strBuffer = Space(255)
   blnReturn = SHGetSpecialFolderPath(0, strBuffer, CSIDL_COMMON_DOCUMENTS, False)
   GetAllUsersDocumentsDir = left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
End Function

Public Function GetSpecialFolder(FolderID As Long) As String
On Local Error Resume Next
Dim blnReturn As Long
Dim strBuffer As String
   
   strBuffer = Space(255)
   blnReturn = SHGetSpecialFolderPath(0, strBuffer, FolderID, False)
   GetSpecialFolder = left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
End Function

Public Function IsFileOpen(Filename As String)
On Local Error Resume Next
Dim filenum As Integer
Dim errnum As Integer
 
    filenum = FreeFile()
    Open Filename For Input Lock Read As #filenum
    Close filenum
    errnum = Err
    Select Case errnum
        Case 0
            IsFileOpen = False
        Case Else
            IsFileOpen = True
    End Select
End Function

Public Function FileCompare(ByVal FilePath1 As String, ByVal FilePath2 As String) As Boolean
On Error GoTo ErrorHandler
Dim lLen1 As Long, lLen2 As Long
Dim iFileNum1 As Integer
Dim iFileNum2 As Integer
Dim bytArr1() As Byte, bytArr2() As Byte
Dim lCtr As Long, lStart As Long
Dim bAns As Boolean
    
    If Dir(FilePath1) = "" Then Exit Function
    If Dir(FilePath2) = "" Then Exit Function
    lLen1 = FileLen(FilePath1)
    lLen2 = FileLen(FilePath2)
    If lLen1 <> lLen2 Then
        Exit Function
    Else
        iFileNum1 = FreeFile
        Open FilePath1 For Binary Access Read As #iFileNum1
        iFileNum2 = FreeFile
        Open FilePath2 For Binary Access Read As #iFileNum2
        bytArr1() = InputB(LOF(iFileNum1), #iFileNum1)
        bytArr2() = InputB(LOF(iFileNum2), #iFileNum2)
        lLen1 = UBound(bytArr1)
        lStart = LBound(bytArr1)
        bAns = True
        For lCtr = lStart To lLen1
            If bytArr1(lCtr) <> bytArr2(lCtr) Then
                bAns = False
                Exit For
            End If
        Next
        FileCompare = bAns
    End If
ErrorHandler:
    If iFileNum1 > 0 Then Close #iFileNum1
    If iFileNum2 > 0 Then Close #iFileNum2
End Function

Public Function LoadTextFile(ByVal FilePath As String, ByRef Text As String) As Boolean
On Error GoTo ErrorHandler:
Dim iFile As Integer

    If FileExists(FilePath) = False Then
        LoadTextFile = False
        Exit Function
    End If
    iFile = FreeFile
    Open FilePath For Input As #iFile
    Text = Input(LOF(iFile), #iFile)
    Close #iFile
    LoadTextFile = True
    Exit Function
   
ErrorHandler:
    If iFile > 0 Then Close #iFile
    LoadTextFile = False
End Function

Public Function LoadTextFileTail(ByVal FilePath As String, ByRef Text As String, Size As Long) As Boolean
On Error GoTo ErrorHandler:
Dim iFile As Integer

    If FileExists(FilePath) = False Then
        LoadTextFileTail = False
        Exit Function
    End If
    iFile = FreeFile
    Open FilePath For Input As #iFile
       
    If (Size >= LOF(iFile)) Then
        Text = Input(LOF(iFile), #iFile)
    Else
        Seek #iFile, LOF(iFile) - Size
        Text = Input(Size - 1, #iFile)
    End If
    
    Close #iFile
    LoadTextFileTail = True
    Exit Function
   
ErrorHandler:
    If iFile > 0 Then Close #iFile
    LoadTextFileTail = False
End Function

Public Sub WriteLogFile(ByVal myFileName As String, ByVal LogString As String)
On Local Error GoTo ErrorHandler
Dim oFile As Integer
    
    oFile = FreeFile
    If FileExists(myFileName) Then
        Open myFileName For Append As #oFile
    Else
        Open myFileName For Output As #oFile
    End If
    Print #oFile, LogString
    Close #oFile
    Exit Sub
ErrorHandler:
    If oFile > 0 Then Close #oFile
End Sub

Public Function WriteFile(FilePath As String, Content As String) As Boolean
On Local Error GoTo ErrorHandler
Dim FSO As Object               'Scripting.FileSystemObject
Dim FSO_TextStream As Object    'Scripting.TextStream
   
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FSO_TextStream = FSO.CreateTextFile(FilePath, True)
    If Not FSO_TextStream Is Nothing Then
        FSO_TextStream.Write Content
        FSO_TextStream.Close
        If FSO.FileExists(FilePath) Then
            WriteFile = True
        End If
    End If
    Exit Function
ErrorHandler:
    WriteFile = False
End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal As Long
Dim sShortPathName As String
Dim iLen As Integer

    'Set up buffer area for API function call return
    iLen = GetShortPathName(sLongFileName, sShortPathName, iLen)
    sShortPathName = Space(iLen)
    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = left(sShortPathName, lRetVal)
End Function

Public Function ReadStdInToFile(Filename As String) As Boolean
Dim FSO As Object               'Scripting.FileSystemObject
Dim FSO_TextStream As Object    'Scripting.TextStream
Dim hStdIn As Long
Dim Buffer As String * 2048
Dim bytes As Long
     
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FSO_TextStream = FSO.CreateTextFile(Filename, True)

'Dim test As String
'Dim mobjSTDIN As Object         'Scripting.TextStream
'    Set mobjSTDIN = FSO.GetStandardStream(StdIn)
'    test = mobjSTDIN.ReadAll
'    FSO_TextStream.Write test
    
    If Not FSO_TextStream Is Nothing Then
        hStdIn = GetStdHandle(STD_INPUT_HANDLE)
        Do
            ReadFile hStdIn, Buffer, Len(Buffer), bytes, 0&
            FSO_TextStream.Write left(Buffer, bytes)
        Loop Until bytes = 0
        FSO_TextStream.Close
        If FSO.FileExists(Filename) Then
            ReadStdInToFile = True
        End If
    End If
End Function

'Public Function FileVersionInfo(FilePath As String) As String
'On Local Error Resume Next
'Dim FSO     'As Scripting.FileSystemObject
'
'    Set FSO = CreateObject("Scripting.FileSystemObject")
'    FileVersionInfo = FSO.GetFileVersion(FilePath)
'    Set FSO = Nothing
'End Function

Public Function FileVersionInfo(sFileName As String) As String
Dim lFileHwnd As Long, lRet As Long, lBufferLen As Long, lplpBuffer As Long, lpuLen As Long
Dim abytBuffer() As Byte
Dim tVerInfo As VS_FIXEDFILEINFO
Dim sBlock As String, sStrucVer As String

    'Get the size File version info structure
    lBufferLen = GetFileVersionInfoSize(sFileName, lFileHwnd)
    If lBufferLen = 0 Then
       Exit Function
    End If

    'Create byte array buffer, then copy memory into structure
    ReDim abytBuffer(lBufferLen)
    Call GetFileVersionInfo(sFileName, 0&, lBufferLen, abytBuffer(0))
    Call VerQueryValue(abytBuffer(0), "\", lplpBuffer, lpuLen)
    Call CopyMem(tVerInfo, ByVal lplpBuffer, Len(tVerInfo))

    'Determine structure version number (For info only)
    sStrucVer = Format$(tVerInfo.StrucVersionh) & "." & Format$(tVerInfo.StrucVersionl)

    'Concatenate file version number details into a result string
    FileVersionInfo = tVerInfo.FileVersionMSh & "." & tVerInfo.FileVersionMSl & "." & tVerInfo.FileVersionLSh & "." & tVerInfo.FileVersionLSl
End Function

Public Function GetIncrementalFilename(FilePath As String) As String
Dim FolderPath As String
Dim NewFileName As String
Dim NewFilePath As String

Dim Filename As String
Dim ExtensionName As String
Dim BaseName As String

Dim i As Long

    FolderPath = StripFilename(FilePath)
    If PathExists(FolderPath) Then
        'folder exists, that's good
        If Not FileExists(FilePath) Then
            'if given filename doesn't exist, everything's ok
            GetIncrementalFilename = FilePath
        Else
            Filename = StripPath(FilePath)
            ExtensionName = GetExtension(Filename)
            BaseName = GetIncrementalBaseName(StripExtension(Filename))
            i = 1
            Do
                NewFilePath = AddBackslash(FolderPath) & BaseName & "(" & i & ")." & ExtensionName
                i = i + 1
            Loop Until (Not FileExists(NewFilePath))
            GetIncrementalFilename = NewFilePath
        End If
    End If
End Function

Public Function GetIncrementalBaseName(BaseName As String) As String
Dim RegEx As Object             'RegExp
Dim Matches As Object           'MatchCollection
Dim Match As Object             'Match
Dim Submatches As Object        'Submatches

    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.MultiLine = True
    RegEx.Global = True
    
    RegEx.Pattern = "^(.*) \([0-9]+\)$"
    Set Matches = RegEx.Execute(BaseName)
    If Matches.Count > 0 Then
        For Each Match In Matches
            Set Submatches = Match.Submatches
            If (Submatches.Count > 0) Then
                GetIncrementalBaseName = Submatches(0)
            End If
            Exit For
        Next
    Else
        GetIncrementalBaseName = BaseName
    End If
    
    Set RegEx = Nothing
    Set Matches = Nothing
    Set Match = Nothing
    Set Submatches = Nothing
End Function

