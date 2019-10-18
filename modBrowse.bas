Attribute VB_Name = "modBrowse"
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ppidl As Long) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBI As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Byte) As Long

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type BROWSEINFO
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Public Const CSIDL_DRIVES = 17
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SETSELECTION = &H466
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const REG_SZ = 1
Public Const KEY_WRITE = &H20006
Public Const SRCCOPY = &HCC0020
Public Const FO_MOVE = &H1
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_FILESONLY = &H80
Public Const FOF_RENAMEONCOLLISION = &H8

'************************************************************
Public Sub SWAP(X, Y)
'************************************************************
'Replacement QB Function (Used In QuickSort)
'************************************************************
  Dim tmp
    tmp = X
    X = Y
    Y = tmp
End Sub

'************************************************************
Public Sub QuickSort(ArrayToSort(), StartEl, NumEls)
'************************************************************
'Standard QuickSort Routine
'************************************************************
  Dim Temp
  Dim First, Last, i, j, StackPtr
  ReDim QStack(NumEls \ 5 + 10)
    First = StartEl
    Last = StartEl + NumEls - 1
      Do
        Do
          Temp = ArrayToSort((Last + First) \ 2)
          i = First
          j = Last
            Do
              While ArrayToSort(i) < Temp
                i = i + 1
              Wend
              While ArrayToSort(j) > Temp
                j = j - 1
              Wend
              If i > j Then Exit Do
              If i < j Then SWAP ArrayToSort(i), ArrayToSort(j)
              i = i + 1
              j = j - 1
            Loop While i <= j
          If i < Last Then
            QStack(StackPtr) = i
            QStack(StackPtr + 1) = Last
            StackPtr = StackPtr + 2
          End If
          Last = j
        Loop While First < Last
        If StackPtr = 0 Then Exit Do
        StackPtr = StackPtr - 2
        First = QStack(StackPtr)
        Last = QStack(StackPtr + 1)
      Loop
 Erase QStack
End Sub

'************************************************************
Public Function BrowseForFolder(OwnerForm As Object, sTitle As String) As String
'************************************************************
'Opens The Windows Folder Open Dialog and Returns Path or ""
'************************************************************
  On Error GoTo errorhandler
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim physpath As String
    Dim retval As Long
      bi.hwndOwner = OwnerForm.hwnd
      retval = SHGetSpecialFolderLocation(OwnerForm.hwnd, CSIDL_DRIVES, bi.pidlRoot)
      bi.pszDisplayName = Space(260)
      bi.lpszTitle = sTitle
      bi.ulFlags = 0
      bi.lpfn = DummyFunc(AddressOf BrowseCallbackProc)
      bi.lParam = 0
      bi.iImage = 0
      pidl = SHBrowseForFolder(bi)
        If pidl <> 0 Then
          bi.pszDisplayName = Left(bi.pszDisplayName, InStr(bi.pszDisplayName, vbNullChar) - 1)
          physpath = Space(260)
          retval = SHGetPathFromIDList(pidl, physpath)
            If retval = 0 Then
              BrowseForFolder = ""
            Else
              physpath = Left(physpath, InStr(physpath, vbNullChar) - 1)
              BrowseForFolder = physpath
            End If
          CoTaskMemFree pidl
        End If
      CoTaskMemFree bi.pidlRoot
      Exit Function
    
errorhandler:
  'Return Nothing
  BrowseForFolder = ""
End Function

'************************************************************
Public Function DummyFunc(ByVal param As Long) As Long
'************************************************************
'Used By BrowseForFolder
'************************************************************
  DummyFunc = param
End Function

'************************************************************
Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
'************************************************************
'Used By BrowseForFolder
'************************************************************
  Dim pathstring As String
  Dim retval As Long

  Select Case uMsg
  Case BFFM_INITIALIZED
    'pathstring = "C:\"
    'retval = SendMessage(hwnd, BFFM_SETSELECTION, ByVal CLng(1), ByVal pathstring)
  End Select
  BrowseCallbackProc = 0
End Function

