Attribute VB_Name = "modFile"

Option Explicit
' Author: Unknown.
'   - Found posted on the Internet.
'
' modFiles.bas
'-------------------------------------------------------------
' Summary of contained methods:
'-------------------------------------------------------------
'   GetWindowsPath()
'   GetSystemPath()
'   FixPath()
'   GetTempFilename()
'   FileExtension()
'   FilePath()
'   FileName()
'   ShortPath()
'   DriveExists()
'   FileExists()
'   FileList()
'   CopyFiles()
'   MoveFiles()
'   RenameFiles()
'   DeleteFiles()

'-------------------------------------------------------------
' API Constant Declarations
'-------------------------------------------------------------
Private Const MAX_PATH = 260

Private Const INVALID_HANDLE_VALUE = -1

Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4

Private Const FOF_SILENT = &H4
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_ALLOWUNDO = &H40

Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

'-------------------------------------------------------------
' API Structure Declarations
'-------------------------------------------------------------
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

Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAborted As Boolean
  hNameMaps As Long
  sProgress As String
End Type

'-------------------------------------------------------------
' API Function Declarations
'-------------------------------------------------------------
Private Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function FindFirstFileA Lib "kernel32" _
    (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFileA Lib "kernel32" _
    (ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" _
    () As Long
Private Declare Function GetShortPathNameA Lib "kernel32" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Private Declare Function GetTempFileNameA Lib "kernel32" _
    (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPathA Lib "kernel32" _
    (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectoryA Lib "kernel32" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetFileAttributes Lib "kernel32.dll" _
  Alias "GetFileAttributesA" ( _
  ByVal lpFileName As String _
  ) As Long
  
Private Declare Function ShellExecute Lib "shell32.dll " Alias "ShellExecuteA " (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
  
  

Const MODULE_NAME = "modFile"

Public Sub Opendir(ByVal DirPath As String, ByVal hwnd As Long)
  ShellExecute hwnd, "open ", DirPath, vbNullString, vbNullString, 5
End Sub
'
'
Public Function GetWindowsPath() As String
  '-------------------------------------------------------------
  '   Returns the path to the Windows folder.
  '-------------------------------------------------------------
  '   API Declarations:
  '-------------------------------------------------------------
  '   MAX_PATH = 260
  '
  '   Private Declare Function GetWindowsDirectoryA Lib "kernel32" _
      '       (ByVal lpBuffer As String, ByVal nSize As Long) As Long
  '-------------------------------------------------------------
  Dim sBuffer As String
  Dim r As Long
  '
  sBuffer = Space$(MAX_PATH)
  r = GetWindowsDirectoryA(sBuffer, MAX_PATH)

  If r Then
    GetWindowsPath = left$(sBuffer, r - 1&)
  End If
  '
End Function


Private Function GetSystemPath() As String
  '-------------------------------------------------------------
  '   Returns the path to the Windows System folder.
  '-------------------------------------------------------------
  '   API Declarations:
  '-------------------------------------------------------------
  '   MAX_PATH = 260
  '
  '   Private Declare Function GetSystemDirectoryA Lib "kernel32" _
      '       (ByVal lpBuffer As String, ByVal nSize As Long) As Long
  '-------------------------------------------------------------
  Dim sBuffer As String
  Dim r As Long

  sBuffer = Space$(MAX_PATH)
  r = GetSystemDirectoryA(sBuffer, MAX_PATH)

  If r Then
    GetSystemPath = left$(sBuffer, r - 1&)
  End If
End Function


Public Function FixPath(ByVal sPath As String) As String
  '-------------------------------------------------------------
  '   Assures that a path ends with "\".
  '-------------------------------------------------------------

  If right$(sPath, 1) <> "\" Then
    FixPath = sPath & "\"
  Else
    FixPath = sPath
  End If
End Function


Public Function GetTempFileName(Optional sPrefix As String) As String
  '-------------------------------------------------------------
  '   Creates a temporary file with a unique name in the
  '   Windows temporary folder.
  '-------------------------------------------------------------
  '   API Declarations:
  '-------------------------------------------------------------
  '   MAX_PATH = 260
  '
  '    Private Declare Function GetTempFileNameA Lib "kernel32" _
       '        (ByVal lpszPath As String, _
       '        ByVal lpPrefixString As String, _
       '        ByVal wUnique As Long, _
       '        ByVal lpTempFileName As String) As Long
  '    Private Declare Function GetTempPathA Lib "kernel32" _
       '        (ByVal nBufferLength As Long, _
       '        ByVal lpBuffer As String) As Long
  '-------------------------------------------------------------

  Dim r As Long
  Dim sTempFile As String
  Dim sTempDir As String

  sTempDir = Space$(MAX_PATH)
  r = GetTempPathA(MAX_PATH, sTempDir)

  If r Then
    sTempDir = left$(sTempDir, r - 1&)

    If Len(sPrefix) = 0 Then
      sPrefix = "tmp"
    End If

    sTempFile = Space$(MAX_PATH)
    r = GetTempFileNameA(sTempDir, sPrefix, 0&, sTempFile)

    If r Then
      GetTempFileName = left$(sTempFile, r - 1&)
    End If
  End If
End Function

Public Function FileExtension(ByVal sFile As String) As String
  '-------------------------------------------------------------
  ' Extracts the extension from a filename is present.
  '-------------------------------------------------------------
  Dim i As Integer

  i = InStrRev(sFile, ".")

  If i > 0 Then
    FileExtension = Mid$(sFile, i + 1)
  End If

End Function


Public Function FilePath(ByVal sFile As String) As String
  '-------------------------------------------------------------
  ' Extracts the path from a fully-qualified filename.
  '-------------------------------------------------------------
  Dim i As Integer

  i = InStrRev(sFile, "\")

  If i > 0 Then
    FilePath = left$(sFile, i)
  End If
End Function


Public Function Filename(ByVal sFullPath As String) As String
  '-------------------------------------------------------------
  ' Extracts the filename from a fully-qualified path.
  '-------------------------------------------------------------
  Dim i As Integer

  i = InStrRev(sFullPath, "\")

  If i > 0 Then
    Filename = Mid$(sFullPath, i + 1)
  Else
    Filename = sFullPath
  End If
End Function


Public Function ShortPath(ByVal LongName As String) As String
  '-------------------------------------------------------------
  '   Returns the short (8.3) filename for a given file or path.
  '-------------------------------------------------------------
  '   API Declarations:
  '-------------------------------------------------------------
  '   MAX_PATH = 260
  '
  '   Private Declare Function GetShortPathNameA Lib "kernel32" _
      '       (ByVal lpszLongPath As String, _
      '       ByVal lpszShortPath As String, _
      '       ByVal cchBuffer As Long) As Long
  '-------------------------------------------------------------
  Dim r As Long
  Dim sBuffer As String

  sBuffer = Space$(MAX_PATH)
  r = GetShortPathNameA(LongName, sBuffer, MAX_PATH)

  If r Then
    ShortPath = left$(sBuffer, r - 1&)
  End If
End Function


Public Function DriveExists(ByVal sDriveLetter As String) As Boolean
  '-------------------------------------------------------------
  '   Determines whether a drive exists.
  '-------------------------------------------------------------
  '   API Declarations:
  '-------------------------------------------------------------
  '   Private Declare Function GetLogicalDrives Lib "kernel32" _
      '       () As Long
  '-------------------------------------------------------------
  Dim dwDrives As Long
  Dim Mask As Long

  dwDrives = GetLogicalDrives()
  Mask = 2 ^ (Asc(UCase$(sDriveLetter)) - 65)

  DriveExists = ((dwDrives And Mask) = Mask)
End Function


Public Function FileList(ByVal sPath As String, _
      ByRef saFileList() As String, _
      Optional ByVal sFileSpec As String = "*") As Long
  '-------------------------------------------------------------
  '   Files saFileList() with a list of all files and folders
  '   in a given path. Returns the number of files.
  '-------------------------------------------------------------
  '   API Declarations:
  '-------------------------------------------------------------
  '   Private Const MAX_PATH = 260
  '   Private Const INVALID_HANDLE_VALUE = -1
  '
  '   Private Type FILETIME
  '        dwLowDateTime As Long
  '        dwHighDateTime As Long
  '    End Type

  '    Private Type WIN32_FIND_DATA
  '        dwFileAttributes As Long
  '        ftCreationTime As FILETIME
  '        ftLastAccessTime As FILETIME
  '        ftLastWriteTime As FILETIME
  '        nFileSizeHigh As Long
  '        nFileSizeLow As Long
  '        dwReserved0 As Long
  '        dwReserved1 As Long
  '        cFileName As String * MAX_PATH
  '        cAlternate As String * 14
  '    End Type
  '
  '    Private Declare Function FindNextFileA Lib "kernel32" _
       '        (ByVal hFindFile As Long, _
       '        lpFindFileData As WIN32_FIND_DATA) As Long
  '
  '   Private Declare Function FindFirstFileA Lib "kernel32" _
      '       (ByVal lpFileName As String, _
      '       lpFindFileData As WIN32_FIND_DATA) As Long
  '
  '   Private Declare Function FindClose Lib "kernel32" _
      '       (ByVal hFindFile As Long) As Long
  '-------------------------------------------------------------

  Dim iCnt As Long
  Dim iMax As Long
  Dim uFIND_DATA As WIN32_FIND_DATA
  Dim r As Long
  Dim hFind As Long
  Dim sName As String

  If right$(sPath, 1) <> "\" Then
    sPath = sPath & "\"
  End If

  sPath = sPath & sFileSpec

  iMax = 49
  ReDim saFileList(iMax)


  hFind = FindFirstFileA(sPath, uFIND_DATA)

  If Not hFind = INVALID_HANDLE_VALUE Then
    sName = uFIND_DATA.cFileName
    If InStr(sName, Chr$(0)) Then
      sName = left$(sName, InStr(sName, Chr$(0)) - 1&)
      If Not sName = "." Then
        If Not sName = ".." Then
          saFileList(0) = sName
          iCnt = 1&
        End If
      End If
    End If

    r = FindNextFileA(hFind, uFIND_DATA)

    Do Until r = 0&
      sName = uFIND_DATA.cFileName

      If InStr(sName, Chr$(0)) Then
        sName = left$(sName, InStr(sName, Chr$(0)) - 1&)
      End If

      If Not sName = "." Then
        If Not sName = ".." Then
          iCnt = iCnt + 1&

          If iCnt >= iMax Then
            iMax = iMax + 50
            ReDim Preserve saFileList(iMax)
          End If

          saFileList(iCnt - 1&) = sName
        End If
      End If

      r = FindNextFileA(hFind, uFIND_DATA)
    Loop

    r = FindClose(hFind)
  End If

  If iCnt = 0& Then
    Erase saFileList()
  Else
    ReDim Preserve saFileList(iCnt - 1&)
  End If

  FileList = iCnt
End Function


'Public Function FileExists(ByVal sFile As String) As Boolean
'  '-------------------------------------------------------------
'  '   Determines whether a file or path exists.
'  '-------------------------------------------------------------
'  '   API Declarations:
'  '-------------------------------------------------------------
'  '   Private Const MAX_PATH = 260
'  '   Private Const INVALID_HANDLE_VALUE = -1
'  '
'  '   Private Type FILETIME
'  '        dwLowDateTime As Long
'  '        dwHighDateTime As Long
'  '    End Type
'
'  '    Private Type WIN32_FIND_DATA
'  '        dwFileAttributes As Long
'  '        ftCreationTime As FILETIME
'  '        ftLastAccessTime As FILETIME
'  '        ftLastWriteTime As FILETIME
'  '        nFileSizeHigh As Long
'  '        nFileSizeLow As Long
'  '        dwReserved0 As Long
'  '        dwReserved1 As Long
'  '        cFileName As String * MAX_PATH
'  '        cAlternate As String * 14
'  '    End Type
'  '
'  '   Private Declare Function FindFirstFileA Lib "kernel32" _
'      '       (ByVal lpFileName As String, _
'      '       lpFindFileData As WIN32_FIND_DATA) As Long
'  '
'  '   Private Declare Function FindClose Lib "kernel32" _
'      '       (ByVal hFindFile As Long) As Long
'  '-------------------------------------------------------------
'  Dim r As Long
'  Dim uFIND_DATA As WIN32_FIND_DATA
'
'  r = FindFirstFileA(sFile, uFIND_DATA)
'  If r = INVALID_HANDLE_VALUE Then
'    FileExists = False
'  Else
'    FileExists = True
'    Call FindClose(r)
'  End If
'
'End Function

Public Function CopyFiles( _
      FileNames As Variant, MoveTo As String, _
      Optional ShowConfirmation As Boolean = True, _
      Optional HideProgress As Boolean = False, _
      Optional RenameOnCollision As Boolean = False) As Boolean
  '-------------------------------------------------------------
  ' Copies the specifed files to a new location. The user can
  ' pass in either a single filename or an array of filenames.
  '-------------------------------------------------------------
  '   Depends on:
  '       CopyMoveFiles()
  '-------------------------------------------------------------

  CopyFiles = CopyMoveFiles(FO_COPY, FileNames, MoveTo, _
      ShowConfirmation, HideProgress, RenameOnCollision)
End Function


Public Function RenameFiles( _
      SrcFile As String, DestFile As String, _
      Optional ShowConfirmation As Boolean = True, _
      Optional RenameOnCollision As Boolean = False) As Boolean
  '-------------------------------------------------------------
  ' Renames the specifed file.
  '-------------------------------------------------------------
  '   Depends on:
  '       CopyMoveFiles()
  '-------------------------------------------------------------

  RenameFiles = CopyMoveFiles(FO_RENAME, SrcFile, DestFile, _
      ShowConfirmation, False, RenameOnCollision)
End Function


Public Function MoveFiles(FileNames As Variant, _
      MoveTo As String, _
      Optional ShowConfirmation As Boolean = True, _
      Optional HideProgress As Boolean = False, _
      Optional RenameOnCollision As Boolean = False) As Boolean
  '-------------------------------------------------------------
  ' Moves the specifed files to a new location. The user can
  ' pass in either a single filename or an array of filenames.
  '-------------------------------------------------------------
  '   Depends on:
  '       CopyMoveFiles()
  '-------------------------------------------------------------

  MoveFiles = CopyMoveFiles(FO_MOVE, FileNames, MoveTo, _
      ShowConfirmation, HideProgress, RenameOnCollision)
End Function


Private Function CopyMoveFiles( _
      Operation As Integer, _
      FileNames As Variant, _
      MoveTo As String, _
      ShowConfirmation As Boolean, _
      HideProgress As Boolean, _
      RenameOnCollision As Boolean) As Boolean
  '-------------------------------------------------------------
  ' Copies/moves/renames the specifed files. The user can
  ' pass in either a single filename or an array of filenames.
  '-------------------------------------------------------------
  '   API Declarations:
  '-------------------------------------------------------------
  '   Private Const FOF_ALLOWUNDO = &H40
  '   Private Const FOF_NOCONFIRMATION = &H10
  '   Private Const FOF_SILENT = &H4
  '
  '   Private Const FO_DELETE = &H3
  '
  '   Private Type SHFILEOPSTRUCT
  '       hwnd As Long
  '       wFunc As Long
  '       pFrom As String
  '       pTo As String
  '       fFlags As Integer
  '       fAborted As Boolean
  '       hNameMaps As Long
  '       sProgress As String
  '   End Type
  '
  '   Private Declare Function SHFileOperation Lib _
      '   "shell32.dll" Alias "SHFileOperationA" _
      '   (lpFileOp As SHFILEOPSTRUCT) As Long
  '-------------------------------------------------------------
  Dim r As Long
  Dim i As Integer
  Dim sOrig As String
  Dim SHFileOp As SHFILEOPSTRUCT

  If IsArray(FileNames) Then
    For i = LBound(FileNames) To UBound(FileNames)
      sOrig = sOrig & FileNames(i) & Chr$(0)
    Next i
  Else
    sOrig = FileNames & Chr$(0)
  End If

  With SHFileOp
    .wFunc = Operation
    .pFrom = sOrig
    .pTo = MoveTo
    If Not ShowConfirmation Then
      .fFlags = FOF_NOCONFIRMATION
    End If
    If HideProgress Then
      .fFlags = .fFlags Or FOF_SILENT
    End If
    If RenameOnCollision Then
      .fFlags = .fFlags Or FOF_RENAMEONCOLLISION
    End If
  End With

  r = SHFileOperation(SHFileOp)

  If r = 0 Then
    CopyMoveFiles = Not SHFileOp.fAborted
  Else
    CopyMoveFiles = False
  End If
End Function


Public Function DeleteFiles(FileNames As Variant, _
      Optional MoveToRecycle As Boolean = False, _
      Optional ShowConfirmation As Boolean = True, _
      Optional HideProgress As Boolean = False) As Long
  '-------------------------------------------------------------
  ' Deletes the specifed files. The user can
  ' pass in either a single filename or an array of filenames.
  '-------------------------------------------------------------
  '   API Declarations:
  '-------------------------------------------------------------
  '   Private Const FOF_ALLOWUNDO = &H40
  '   Private Const FOF_NOCONFIRMATION = &H10
  '   Private Const FOF_SILENT = &H4
  '
  '   Private Const FO_DELETE = &H3
  '
  '   Private Type SHFILEOPSTRUCT
  '       hwnd As Long
  '       wFunc As Long
  '       pFrom As String
  '       pTo As String
  '       fFlags As Integer
  '       fAborted As Boolean
  '       hNameMaps As Long
  '       sProgress As String
  '   End Type
  '
  '   Private Declare Function SHFileOperation Lib _
      '   "shell32.dll" Alias "SHFileOperationA" _
      '   (lpFileOp As SHFILEOPSTRUCT) As Long
  '-------------------------------------------------------------

  Dim sDest As String
  Dim i As Integer
  Dim SHFileOp As SHFILEOPSTRUCT
  Dim r As Long

  With SHFileOp
    If MoveToRecycle Then
      .fFlags = FOF_ALLOWUNDO
    End If
    If Not ShowConfirmation Then
      .fFlags = .fFlags Or FOF_NOCONFIRMATION
    End If
    If HideProgress Then
      .fFlags = .fFlags Or FOF_SILENT
    End If
  End With

  If IsArray(FileNames) Then
    For i = LBound(FileNames) To UBound(FileNames)
      sDest = sDest & FileNames(i) & Chr$(0)
    Next i
  Else
    sDest = FileNames & Chr$(0)
  End If

  With SHFileOp
    .wFunc = FO_DELETE
    .pFrom = sDest
  End With

  r = SHFileOperation(SHFileOp)
  If r <> 0 Then
    DeleteFiles = False
  Else
    DeleteFiles = Not SHFileOp.fAborted
  End If
End Function

Public Function IsDirectory(sPath As String) As Boolean
' Author: (Lost).
' Purpose: To determine whether is not a given path points to a folder.
' License: GPL 2.0+
  On Error GoTo ErrorThis
  IsDirectory = GetAttr(sPath) Or vbDirectory
  Exit Function
ErrorThis:
  IsDirectory = False
End Function

'Public Function IsFile(spath As String) As Boolean
'' Purpose: To determine whether or not a path points to a directory or a file
'' Example/Note: xxx
'' !! Assumes/Pre: Nothing
'' Parameters:
''  sPath as String  -
'' Returns: Boolean
''       Success-
''       Failure- Raises error on failure
'' Revision history:
''   2005-Nov-14, @ 1618 [Michael Johnson] Initial creation
'  Call TraceEnters(MODULE_NAME & "::IsFile")
'  TraceDetail = "To determine whether or not a path points to a directory or a file"
'
'            Stop    ' *** Under construction *** This portion of code may be incomplete !!!
'  IsFile = FileExists(spath) And Not IsDirectory(spath)
'ExitThis:
'  Call TraceExits
'  Exit Function
'End Function

'Public Function FileExists(spath As String) As Boolean
'' Purpose: To determine whether or not a given path points to a file.
'' Author: (Lost).
'' License: GPL 2.0+
'  On Error GoTo ErrorThis
'  FileExists = GetAttr(spath)
'  Exit Function
'ErrorThis:
'  FileExists = False
'End Function



