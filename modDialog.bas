Attribute VB_Name = "modDialog"

'
'    Project: The Black Box
'      Modul: XCDlg
'   Contents: Replaces with simple-to-use-calls the usage of some
'             functions of the Microsoft Windows Common Controls-OCX.
'             This will protect your install-application from
'             one of the most common version-conflicts!
'
' Specials: It is not possible to rebuild the Find-And-Replace-Functions
'           of the system in a module. For these functions you need to
'           install a new messageloop with the special Find-And-Replace-
'           messages - impossible without a form and only in a module.
'
'         Created on: 11-17-1998
'         Created by: JoeHurst@snafu.de
' Last changes on/by: 01-22-1999: JH
'                     - DLG_Open/DLG_Save now deal with FilterIndex (in/out)
'
' License: Believed to be public domain.
'
' ------------------------------------------------------------------------------------------------------
Option Explicit

Private Const WM_USER = &H400
Private Const WM_INITDIALOG = &H110
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetParent Lib "user32" (ByVal Hwnd As Long) As Long
Private Const SPI_GETWORKAREA = 48
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private hOwner As Long

' This array allows you to save custom colors for the common dialog-function
' "ChooseColor". You may read and write this array. Make sure you store only
' long values in RGB-Format!
Public THECUSTOMCOLORS(16) As Long

' Use these constantes to define the flags of DLG_Open and DLG_Save.
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20

' Use these constants to define the wCommand-Parameter in DLG_Help
Public Const HELP_COMMAND = &H102&
Public Const HELP_CONTENTS = &H3&
Public Const HELP_CONTEXT = &H1          '  Display topic in ulTopic
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_FINDER = &HB
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_HELPONHELP = &H4       '  Display help on using help
Public Const HELP_INDEX = &H3            '  Display index
Public Const HELP_KEY = &H101            '  Display topic for keyword in dwData
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_QUIT = &H2             '  Terminate help
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_SETINDEX = &H5         '  Set current Index for multi index help
Public Const HELP_SETWINPOS = &H203&

' Use these constants to define the flags of DLG_Color
Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

'Choose Font
'Private Const CF_ANSIONLY& = &H400&
Private Const CF_APPLY& = &H200&
'Private Const CF_BITMAP& = 2
Private Const CF_SCREENFONTS& = &H1
Private Const CF_PRINTERFONTS& = &H2
'Private Const CF_BOTH& = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'Private Const CF_DIB& = 8
'Private Const CF_DIF& = 5
'Private Const CF_DSPBITMAP& = &H82
'Private Const CF_DSPENHMETAFILE& = &H8E
'Private Const CF_DSPMETAFILEPICT& = &H83
'Private Const CF_DSPTEXT& = &H81
Public Const CF_EFFECTS& = &H100&
Private Const CF_ENABLEHOOK& = &H8&
Private Const CF_ENABLETEMPLATE& = &H10&
'Private Const CF_ENABLETEMPLATEHANDLE& = &H20&
'Private Const CF_ENHMETAFILE& = 14
'Private Const CF_FIXEDPITCHONLY& = &H4000&
'Private Const CF_FORCEFONTEXIST& = &H10000
'Private Const CF_GDIOBJFIRST& = &H300
'Private Const CF_GDIOBJLAST& = &H3FF
'Private Const CF_HDROP& = 15
Private Const CF_INITTOLOGFONTSTRUCT& = &H40&
Private Const CF_LIMITSIZE& = &H2000&
'Private Const CF_LOCALE& = 16
'Private Const CF_MAX& = 17
'Private Const CF_METAFILEPICT& = 3
'Private Const CF_NOFACESEL& = &H80000
'Private Const CF_NOVECTORFONTS& = &H800&
'Private Const CF_NOOEMFONTS& = CF_NOVECTORFONTS
'Private Const CF_NOSCRIPTSEL& = &H800000
'Private Const CF_NOSIMULATIONS& = &H1000&
'Private Const CF_NOSIZESEL& = &H200000
'Private Const CF_NOSTYLESEL& = &H100000
'Private Const CF_NOVERTFONTS& = &H1000000
'Private Const CF_OEMTEXT& = 7
'Private Const CF_OWNERDISPLAY& = &H80
'Private Const CF_PALETTE& = 9
'Private Const CF_PENDATA& = 10
'Private Const CF_PRIVATEFIRST& = &H200
'Private Const CF_PRIVATELAST& = &H2FF
'Private Const CF_RIFF& = 11
'Private Const CF_SCALABLEONLY& = &H20000
'Private Const CF_SCRIPTSONLY& = CF_ANSIONLY
'Private Const CF_SELECTSCRIPT& = &H400000
'Private Const CF_SHOWHELP& = &H4&
'Private Const CF_SYLK& = 4
'Private Const CF_TEXT& = 1
'Private Const CF_TIFF& = 6
'Private Const CF_TTONLY& = &H40000
'Private Const CF_UNICODETEXT& = 13
'Private Const CF_USESTYLE& = &H80&
'Private Const CF_WAVE& = 12
'Private Const CF_WYSIWYG& = &H8000&

' Use these constants to define the flags of DLG_Printer and DLG_PageSetup
'Public Const PD_ALLPAGES = &H0
'Public Const PD_COLLATE = &H10
'Public Const PD_DISABLEPRINTTOFILE = &H80000
'Public Const PD_HIDEPRINTTOFILE = &H100000
'Public Const PD_NONETWORKBUTTON = &H200000
'Public Const PD_NOPAGENUMS = &H8
'Public Const PD_NOSELECTION = &H4
'Public Const PD_NOWARNING = &H80
'Public Const PD_PAGENUMS = &H2
'Public Const PD_PRINTSETUP = &H40
'Public Const PD_PRINTTOFILE = &H20
'Public Const PD_RETURNDC = &H100
'Public Const PD_RETURNDEFAULT = &H400
'Public Const PD_RETURNIC = &H200
'Public Const PD_SELECTION = &H1
'Public Const PD_SHOWHELP = &H800
'Public Const PD_USEDEVMODECOPIES = &H40000
'Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000

' Use these constants to define the flags of DLG_PageSetup
'Public Const PSD_DEFAULTMINMARGINS = &H0 '  default (printer's)
'Public Const PSD_DISABLEMARGINS = &H10
'Public Const PSD_DISABLEORIENTATION = &H100
'Public Const PSD_DISABLEPAGEPAINTING = &H80000
'Public Const PSD_DISABLEPAPER = &H200
'Public Const PSD_DISABLEPRINTER = &H20
'Public Const PSD_ENABLEPAGEPAINTHOOK = &H40000
'Public Const PSD_ENABLEPAGESETUPHOOK = &H2000 '  must be same as PD_*
'Public Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000 '  must be same as PD_*
'Public Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000 '  must be same as PD_*
'Public Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8 '  3rd of 4 possible
'Public Const PSD_INTHOUSANDTHSOFINCHES = &H4 '  2nd of 4 possible
'Public Const PSD_INWININIINTLMEASURE = &H0 '  1st of 4 possible
'Public Const PSD_MARGINS = &H2 '  use caller's
'Public Const PSD_MINMARGINS = &H1 '  use caller's
'Public Const PSD_NOWARNING = &H80 '  must be same as PD_*
'Public Const PSD_RETURNDEFAULT = &H400 '  must be same as PD_*
'Public Const PSD_SHOWHELP = &H800 '  must be same as PD_*

' Browse for Folder
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000

Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11

Public Const BFFM_SETSELECTION = (WM_USER + 103)
Public Const BFFM_ENABLEDOK = (WM_USER + 101)

Private Type tOPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Type tCHOOSECOLOR
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As Long
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Type tPRINTDLG
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  hdc As Long
  flags As Long
  nFromPage As Integer
  nToPage As Integer
  nMinPage As Integer
  nMaxPage As Integer
  nCopies As Integer
  hInstance As Long
  lCustData As Long
  lpfnPrintHook As Long
  lpfnSetupHook As Long
  lpPrintTemplateName As String
  lpSetupTemplateName As String
  hPrintTemplate As Long
  hSetupTemplate As Long
End Type

Private Type tPOINTAPI
  X As Long
  Y As Long
End Type

Private Type tRECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type tPAGESETUPDLG
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  flags As Long
  ptPaperSize As tPOINTAPI
  rtMinMargin As tRECT
  rtMargin As tRECT
  hInstance As Long
  lCustData As Long
  lpfnPageSetupHook As Long
  lpfnPagePaintHook As Long
  lpPageSetupTemplateName As String
  hPageSetupTemplate As Long
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const LF_FACESIZE& = 32

Private Declare Function ChooseFont Lib "COMDLG32" _
    Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long

Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

'Private Declare Function MulDiv Lib "kernel32" _
'    (ByVal nNumber As Long, _
'    ByVal nNumerator As Long, _
'    ByVal nDenominator As Long) As Long

'Private Declare Function GetDeviceCaps Lib "gdi32" _
'    (ByVal hdc As Long, ByVal nIndex As Long) As Long

'Private Const LOGPIXELSY = 90

'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE - 1) As Byte
End Type

Public Type TCHOOSEFONT
    lStructSize As Long
    hwndOwner As Long ' caller's window handle
    hdc As Long ' printer DC/IC or NULL
    lpLogFont As Long
    iPointSize As Long ' 10 * size in points of selected font
    flags As Long ' enum. Type flags
    rgbColors As Long ' returned text color
    lCustData As Long ' data passed to hook fn.
    lpfnHook As Long ' ptr. to hook Function
    lpTemplateName As String ' custom template name
    hInstance As Long ' instance handle of.EXE that
    ' contains cust. dlg. template
    lpszStyle As String ' return the style field here
    ' must be LF_FACESIZE or bigger
    nFontType As Integer ' same value reported to the EnumFonts
    ' call back with the extra FONTTYPE_
    ' bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long ' minimum pt size allowed &
    nSizeMax As Long ' max pt size allowed If
    ' CF_LIMITSIZE is used
End Type
Public Enum EFontType
    Simulated_FontType = &H8000
    Printer_FontType = &H4000
    Screen_FontType = &H2000
    Bold_FontType = &H100
    Italic_FontType = &H200
    Regular_FontType = &H400
End Enum

'Public Type CHOOSEFONTRETURN
'    'lfHeight As Long
'    'lfWidth As Long
'    'lfEscapement As Long
'    'lfOrientation As Long
'    lfBold As Long
'    lfItalic As Byte
'    lfSize As Integer
'    'lfUnderline As Byte
'    'lfStrikeOut As Byte
'    'lfColor As Long
'    'lfCharSet As Byte
'    'lfOutPrecision As Byte
'    'lfClipPrecision As Byte
'    'lfQuality As Byte
'    'lfPitchAndFamily As Byte
'    lfFaceName As String
'    lfOK As Long
'End Type


Private Declare Function VarPtr Lib "msvbvm50.dll" (var As Any) As Long
'Declare Function VarPtr Lib "msvbvm60.dll" (var As Any) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias _
    "GetSaveFileNameA" (pOpenfilename As tOPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
    "GetOpenFileNameA" (pOpenfilename As tOPENFILENAME) As Long
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal Hwnd As Long, ByVal lpHelpFile As String, _
    ByVal wCommand As Long, ByVal dwData As Any) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias _
    "ChooseColorA" (pChoosecolor As tCHOOSECOLOR) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" _
    (pPrintdlg As tPRINTDLG) As Long
Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias _
    "PageSetupDlgA" (pPagesetupdlg As tPAGESETUPDLG) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
    (ByVal Hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
    ByVal hIcon As Long) As Long

'Browse for Folder
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
    "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
    "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
   lParam As Any) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
    pSource As Any, ByVal dwLength As Long)

Private Const MAX_PATH = 260
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

Private Declare Function SHSimpleIDListFromPath Lib _
   "shell32" Alias "#162" _
   (ByVal szPath As String) As Long
Private Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" _
   (lpString1 As Any, lpString2 As Any) As Long

Private Declare Function lstrlenA Lib "kernel32" _
   (lpString As Any) As Long

Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32_NT = 2
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As OSVERSIONINFO) As Long

'Netzlaufwerk verbinden / trennen
Private Const RESOURCETYPE_DISK = &H1
'Private Const RESOURCETYPE_PRINT = &H2
Private Declare Function WNetConnectionDialog Lib "mpr.dll" _
    (ByVal Hwnd As Long, ByVal dwType As Long) As Long
Private Declare Function WNetDisconnectDialog Lib "mpr.dll" _
    (ByVal Hwnd As Long, ByVal dwType As Long) As Long


Private Function BrowseCallbackProcStr(ByVal Hwnd As Long, _
    ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
                                       
Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(Hwnd, BFFM_SETSELECTIONA, _
                          True, ByVal lpData)
                          
         Case Else:
         
   End Select
          
End Function

Private Function BrowseCallbackProc(ByVal Hwnd As Long, _
    ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
 
  'Callback for the Browse PIDL method.
 
  'On initialization, set the dialog's
  'pre-selected folder using the pidl
  'set as the bi.lParam, and passed back
  'to the callback as lpData param.
 
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(Hwnd, BFFM_SETSELECTIONA, _
                          False, ByVal lpData)
                          
         Case Else:
         
   End Select

End Function


Private Function FARPROC(pfn As Long) As Long
  
  'A dummy procedure that receives and returns
  'the value of the AddressOf operator.
 
  'Obtain and set the address of the callback
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
 
  FARPROC = pfn

End Function

Public Function BrowseForFolderByPath(Hwnd As Long, Titel As String, _
    sSelPath As String) As String
'Browser for Folder mit Übergabe eines Pfades

  Dim BI As BROWSEINFO
  Dim pidl As Long
  Dim lpSelPath As Long
  Dim spath As String * MAX_PATH
  
  With BI
    .hOwner = Hwnd
    .pidlRoot = 0
    .lpszTitle = Titel
    .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
    
    lpSelPath = LocalAlloc(LPTR, Len(sSelPath))
    MoveMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath)
    .lParam = lpSelPath
    
    End With
    
   pidl = SHBrowseForFolder(BI)
   
   If pidl Then
     
      If SHGetPathFromIDList(pidl, spath) Then
         BrowseForFolderByPath = Left$(spath, InStr(1, spath, vbNullChar) - 1)
      End If
      
      Call CoTaskMemFree(pidl)
   
   End If
   
  Call LocalFree(lpSelPath)

End Function


Public Function BrowseForFolderByPIDL(Hwnd As Long, Titel As String, _
    sSelPath As String) As String
'Browser for Folder mit Übergabe von PIDL

   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim spath As String * MAX_PATH
  
   With BI
      .hOwner = Hwnd
      .pidlRoot = 0
      .lpszTitle = Titel
      .lpfn = FARPROC(AddressOf BrowseCallbackProc)
      .lParam = GetPIDLFromPath(sSelPath) 'replaces '= SHSimpleIDListFromPath(sSelPath)'
   End With
  
   pidl = SHBrowseForFolder(BI)
  
   If pidl Then
      If SHGetPathFromIDList(pidl, spath) Then
         BrowseForFolderByPIDL = Left$(spath, InStr(1, spath, vbNullChar) - 1)
      End If
     
     'free the pidl returned by call to SHBrowseForFolder
      Call CoTaskMemFree(pidl)
  End If
  
 'free the pidl set in call to GetPIDLFromPath
  Call CoTaskMemFree(BI.lParam)
  
End Function


Private Function GetPIDLFromPath(spath As String) As Long

  'return the pidl to the path supplied by calling the
  'undocumented API #162 (our name SHSimpleIDListFromPath).
  'This function is necessary as, unlike documented APIs,
  'the API is not implemented in 'A' or 'W' versions.

  If IsWinNT Then
    GetPIDLFromPath = SHSimpleIDListFromPath(StrConv(spath, vbUnicode))
  Else
    GetPIDLFromPath = SHSimpleIDListFromPath(spath)
  End If

End Function


Private Function IsWinNT() As Boolean
'Windows NT?
    Dim OSV As OSVERSIONINFO
    OSV.OSVSize = Len(OSV)
    'API returns 1 if a successful call
    If GetVersionEx(OSV) = 1 Then
        'PlatformId contains a value representing
        'the OS, so if its VER_PLATFORM_WIN32_NT,
        'return true
        IsWinNT = OSV.PlatformID = VER_PLATFORM_WIN32_NT
    End If
End Function


Private Function UnqualifyPath(spath As String) As String
'Entfernt allfällige Backslashs

  'qualifying a path usually involves assuring
  'that its format is valid, including a trailing slash
  'ready for a filename. Since the SHBrowseForFolder API
  'will pre-select the path if it contains the trailing
  'slash, I call stripping it 'unqualifying the path'.
   If Len(spath) > 0 Then
   
      If Right$(spath, 1) = "\" Then
      
         UnqualifyPath = Left$(spath, Len(spath) - 1)
         Exit Function
      
      End If
   
   End If
   
   UnqualifyPath = spath
   
End Function

' Get a filename to save something to.
' This module does not need the Microsoft Windows Common Controls-OCX! (!nl)
' For standard, use this constant-expression for the flags: (!nl)
' OFN_PATHMUSTEXIST+OFN_OVERWRITEPROMPT+OFN_HIDEREADONLY.
' (For getting a save-filename, always add OFN_HIDEREADONLY!) (!nl)
' In: ParentHwnd: hwnd of parent-form or 0. (!nl)
' Presetfilename: do not use a complete path, just a single filename! (!nl)
' Set unused parameters to vbnullstring$. (!nl)
' FilterIndex=Index of preselected filter-text and out-value of selection
' by user (!nl)
' Out: ReturnFilename=Name of file to save to or empty string
' Exists since Black-Box-Version: 1.3
Function DLG_Save(ParentHwnd As Long, ByVal DlgFilter As String, _
                  ByVal DlgPresetFilename As String, ByVal DefaultExtension As String, _
                  ByVal flags As Long, ByVal DlgTitle As String, _
                  ByRef ReturnFileName As String, _
                  ByRef FilterIndex As Long, _
                  Optional InitialDir As String) As Long

On Error GoTo 0

  Dim OFN As tOPENFILENAME
  Dim FLong As String, result As String
  Dim InitDir As String, n As String, ret As String
  Dim r As Long
  hOwner = ParentHwnd
  
  n = vbNullChar$ + vbNullChar$
  For r = 1 To Len(DlgFilter)
    If Mid$(DlgFilter, r, 1) = "|" Then Mid$(DlgFilter, r, 1) = vbNullChar$
  Next r
  ret = DlgPresetFilename + String(2048 - Len(DlgPresetFilename), 0)
  With OFN
    .lStructSize = Len(OFN)
    .hwndOwner = ParentHwnd
    .lpstrFilter = DlgFilter + n
    If Len(DlgFilter) > 2 Then .nFilterIndex = 1
    .lpstrFile = ret
    .nMaxFile = 2048
    .lpstrTitle = DlgTitle + n
    .lpstrDefExt = DefaultExtension + n
    ' allow help only if ParentHwnd is set
    If ParentHwnd = 0 Then flags = flags And Not OFN_SHOWHELP
    .flags = flags
    .nFilterIndex = FilterIndex
    .lpstrInitialDir = InitialDir
  End With
  
  r = GetSaveFileName(OFN)
  
  If r Then   ' OK
    ret = OFN.lpstrFile
    If InStr(ret, vbNullChar$) Then
      FLong = Mid(ret, 1, InStr(ret, vbNullChar$) - 1)
    Else
      FLong = ret
    End If
    FilterIndex = OFN.nFilterIndex
  Else      ' Cancel
    FLong = ""
  End If
  ReturnFileName = FLong
  DLG_Save = r
End Function

' Get a filename to load something from.
' This module does not need the Microsoft Windows Common Controls-OCX! (!nl)
' Set unused parameters to vbnullstring$. For standard, use this expression for the
' flags: (!nl)
' OFN_PATHMUSTEXIST+OFN_FILEMUSTEXIST+OFN_HIDEREADONLY  . (!nl)
' OpenWriteProtect is a returnvalue. If you get TRUE, the user wants
' to open a file writeprotected, so your application has to react correct! (!nl)
' FilterIndex=Index of preselected filter-text and out-value of selection
' by user (!nl)
' Out: ReturnFilename=Name of file to save to or empty string.
'      Result: 0 on error
' Exists since Black-Box-Version: 1.3
Function DLG_Open(ParentHwnd As Long, ByVal DlgFilter As String, _
                  ByVal DlgPresetFilename As String, _
                  ByVal DefaultExtension As String, _
                  ByVal flags As Long, _
                  ByVal DlgTitle As String, _
                  ByRef OpenWriteProtect As Boolean, _
                  ByRef ReturnFileName As String, _
                  ByRef FilterIndex As Long, _
                  Optional InitialDir As String) As Long
                  
  Dim OFN As tOPENFILENAME
  Dim FLong As String
  Dim n As String, ret As String
  Dim r As Long
  hOwner = ParentHwnd
  
  n = vbNullChar$ + vbNullChar$
  For r = 1 To Len(DlgFilter)
    If Mid$(DlgFilter, r, 1) = "|" Then Mid$(DlgFilter, r, 1) = vbNullChar$
  Next r
  ret = DlgPresetFilename + String(2048 - Len(DlgPresetFilename), 0)
  
  With OFN
    .lStructSize = Len(OFN)
    .hwndOwner = ParentHwnd
    .lpstrFilter = DlgFilter + n
    If Len(DlgFilter) > 2 Then .nFilterIndex = 1
    .lpstrFile = ret
    .nMaxFile = 2048
    .lpstrTitle = DlgTitle + n
    .lpstrDefExt = DefaultExtension + n
    ' allow help only if ParentHwnd is set
    If ParentHwnd = 0 Then flags = flags And Not OFN_SHOWHELP
    .flags = flags
    .nFilterIndex = FilterIndex
    .lpstrInitialDir = InitialDir
  End With
  
  r = GetOpenFileName(OFN)
  
  With OFN
    If r Then   ' OK
      ret = .lpstrFile
      If InStr(ret, vbNullChar$) Then
        FLong = Mid(ret, 1, InStr(ret, vbNullChar$) - 1)
      Else
        FLong = ret
      End If
      OpenWriteProtect = ((.flags And OFN_READONLY) <> 0)
      FilterIndex = OFN.nFilterIndex
    Else      ' Cancel
      FLong = ""
      OpenWriteProtect = False
    End If
  End With
  ReturnFileName = FLong
  DLG_Open = r
End Function

' Get a color to use.
' This module does not need the Microsoft Windows Common Controls-OCX! (!nl)
' A standard for ColorFlags is CC_RGBINIT. ResultColor is the result of this
' function, you may use it to preset/preselect a color. (!nl)
' Because a result of 0 can also mean that color 0 (=blackest black) was selected,
' this function returns -1 for Error/Cancel.
' Exists since Black-Box-Version: 1.3
Function DLG_Color(ParentHwnd As Long, flags As Long, cColor As Long) As Long
  Dim CC As tCHOOSECOLOR
  Dim r As Long
  
  With CC
    .lStructSize = Len(CC)
    .hwndOwner = ParentHwnd
    .rgbResult = cColor
    .lpCustColors = VarPtr(THECUSTOMCOLORS(1))
    ' allow help only if ParentHwnd is set
    If ParentHwnd = 0 Then flags = flags And Not CC_SHOWHELP
    .flags = flags Or CC_RGBINIT
  End With
  r = ChooseColor(CC)
  If r Then
    cColor = CC.rgbResult
  Else
    cColor = -1
  End If
  DLG_Color = r
End Function

'' Get (new) printer-settings.
'' This module does not need the Microsoft Windows Common Controls-OCX! (!nl)
'' NumberOfCopies and further are input values and return output values! Set
'' to zero on input if you don't want to preset the dialog! (!nl)
'' Don't panic because you have no access to the Dlg-Structure; this function
'' gives all needed information in it's parameter-variables to you: PrintToFile and
'' (PrintOnlySelection or PageString), you just have to parse PageString and Print
'' to PrinterDC (or a file), combined with NumberOfCopies.
'' Exists since Black-Box-Version: 1.3
'Function DLG_Printer(ParentHwnd As Long, flags As Long, _
'                     FromPage As Long, ToPage As Long, _
'                     MinPage As Long, MaxPage As Long, _
'                     PrinterDC As Long, _
'                     PrintToFile As Boolean, _
'                     NumberOfCopies As Long, _
'                     PrintOnlySelection As Boolean, _
'                     PageString As String, _
'                     PrintSorted As Boolean) As Long
'
'  Dim pd As tPRINTDLG
'  Dim r As Long
'
'  With pd
'    .lStructSize = Len(pd)
'    .hdc = PrinterDC
'    .hwndOwner = ParentHwnd
'    .nCopies = NumberOfCopies
'    .nFromPage = FromPage
'    .nToPage = ToPage
'    .nMaxPage = MaxPage
'    .nMinPage = MinPage
'    ' make sure a PrinterDC is always valid
'    flags = flags Or PD_RETURNDC
'    ' allow help only if ParentHwnd is set
'    If ParentHwnd = 0 Then flags = flags And Not PD_SHOWHELP
'    .flags = flags
'
'  End With
'
'  r = PrintDlg(pd)
'
'  If r Then
'    With pd
'      FromPage = .nFromPage
'      ToPage = .nToPage
'      MaxPage = .nMaxPage
'      MinPage = .nMinPage
'      PrinterDC = .hdc
'      NumberOfCopies = .nCopies
'      PageString = "" & .nMinPage & "-" & .nMaxPage
'      If (.flags And PD_PAGENUMS) Then PageString = "" & .nFromPage & "-" & .nToPage
'      PrintToFile = ((.flags And PD_PRINTTOFILE) <> 0)
'      PrintOnlySelection = ((.flags And PD_SELECTION) <> 0)
'      PrintSorted = ((.flags And PD_COLLATE) <> 0)
'      If PrintOnlySelection Then PageString = ""
'    End With
'  End If
'
'  DLG_Printer = r
'End Function

'' Call the dialog to setup the page-settings before printing a file.
'' This module does not need the Microsoft Windows Common Controls-OCX! (!nl)
'' A standard-value for flags is: PD_RETURNDC
'' Exists since Black-Box-Version: 1.3
'Function DLG_PageSetup(ParentHwnd As Long, flags As Long) As Long
'  Dim r As Long, PS As tPAGESETUPDLG
'
'  With PS
'    .lStructSize = Len(PS)
'    .hwndOwner = ParentHwnd
'    ' allow help only if ParentHwnd is set
'    If ParentHwnd = 0 Then flags = flags And Not PSD_SHOWHELP
'    .flags = PD_RETURNDC     'flags
'  End With
'
'  r = PageSetupDlg(PS)
'  If r Then
'  Else
'  End If
'  DLG_PageSetup = r
'End Function

Function DLG_Font(CurFont As Font, _
                      Optional PrinterDC As Long = -1, _
                      Optional Owner As Long = -1, _
                      Optional Color As Long = vbBlack, _
                      Optional MinSize As Long = 0, _
                      Optional MaxSize As Long = 0, _
                      Optional flags As Long = 0) As Boolean
    
    Dim m_lApiReturn As Long ', m_lExtendedError As Long
    
    hOwner = Owner
    m_lApiReturn = 0
    ' Unwanted Flags bits
    Const CF_FontNotSupported = CF_APPLY Or CF_ENABLETEMPLATE
    
    ' Flags can get reference variable or constant with bit flags
    ' PrinterDC can take printer DC
    If PrinterDC = -1 Then
        PrinterDC = 0
        If flags And CF_PRINTERFONTS Then PrinterDC = Printer.hdc
    Else
        flags = flags Or CF_PRINTERFONTS
    End If
    ' Must have some fonts
    If (flags And CF_PRINTERFONTS) = 0 Then flags = flags Or CF_SCREENFONTS
    ' Color can take initial color, receive chosen color
    If Color <> vbBlack Then flags = flags Or CF_EFFECTS
    ' MinSize can be minimum size accepted
    If MinSize Then flags = flags Or CF_LIMITSIZE
    ' MaxSize can be maximum size accepted
    If MaxSize Then flags = flags Or CF_LIMITSIZE

    ' Put in required internal flags and remove unsupported
    flags = (flags Or CF_INITTOLOGFONTSTRUCT) And Not CF_FontNotSupported
    
    ' Initialize LOGFONT variable
    Dim fnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    fnt.lfHeight = -(CurFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = CurFont.Weight
    fnt.lfItalic = CurFont.Italic
    fnt.lfUnderline = CurFont.Underline
    fnt.lfStrikeOut = CurFont.Strikethrough
    ' Other fields zero
    StrToBytes fnt.lfFaceName, CurFont.Name

    ' Initialize TCHOOSEFONT variable
    Dim cf As TCHOOSEFONT
    cf.lStructSize = Len(cf)
    If Owner <> -1 Then cf.hwndOwner = Owner
    cf.hdc = PrinterDC
    cf.lpLogFont = VarPtr(fnt)
    cf.iPointSize = CurFont.Size * 10
    cf.flags = flags
    cf.rgbColors = Color
    cf.nSizeMin = MinSize
    cf.nSizeMax = MaxSize

    ' All other fields zero
    m_lApiReturn = ChooseFont(cf)
    Select Case m_lApiReturn
    Case 1
        ' Success
        DLG_Font = True
        flags = cf.flags
        Color = cf.rgbColors
        CurFont.Bold = cf.nFontType And Bold_FontType
        'CurFont.Italic = cf.nFontType And Italic_FontType
        CurFont.Italic = fnt.lfItalic
        CurFont.Strikethrough = fnt.lfStrikeOut
        CurFont.Underline = fnt.lfUnderline
        CurFont.Weight = fnt.lfWeight
        CurFont.Size = cf.iPointSize / 10
        CurFont.Name = BytesToStr(fnt.lfFaceName)
    Case 0
        ' Cancelled
        DLG_Font = False
    Case Else
'        ' Extended error
'        m_lExtendedError = CommDlgExtendedError()
        DLG_Font = False
    End Select
        
End Function
Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

Private Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        'If UnicodeTypeLib Then
        '    Dim st As String
        '    st = StrConv(s, vbFromUnicode)
        '    CopyMemoryStr ab(LBound(ab)), st, cab
        'Else
            CopyMemoryStr ab(LBound(ab)), s, cab
        'End If
    End If
End Sub
Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function

' Open winhelp to display a helpfile.
' When you use winhelp in your application, do not ignore this hint:
' Use the propertie HelpContextID in every form and for every possible
' control. So you save very much keyboard-tracking (F1 pressed?) and
' very much single calls of any help-procedure to start winhelp. VB
' supports winhelp very good, so do not try to make it on your own.
' Exists since Black-Box-Version: 1.3
Function DLG_Winhelp(ParentHwnd As Long, ByVal HelpFile As String, _
    ByVal wCommand As Long, ByVal dwData As String) As Long
  DLG_Winhelp = WinHelp(ParentHwnd, HelpFile, wCommand, dwData)
End Function

'' Generate the standard windows About-box. This is not part of the Microsoft
'' Windows Common Controls-OCX, but is easy to use. ParentForm must be set, no
'' 0-value is allowed!
'' Exists since Black-Box-Version: 1.3
'Sub DLG_AboutBox(ParentForm As Form, Copyright As String)
'  ShellAbout ParentForm.hwnd, App.ProductName, Copyright, ParentForm.Icon
'End Sub
'
'Public Function DLG_BrowseFolder(ParentHwnd As Long, szDialogTitle As String, _
'Folder As String) As Long
''Browse for Folder normal
'
'    Dim X As Long, BI As BROWSEINFO, dwIList As Long, szPath As String, wPos As Integer
'
'    With BI
'        .hOwner = ParentHwnd
'        .lpszTitle = szDialogTitle
'        .ulFlags = BIF_RETURNONLYFSDIRS
'    End With
'    dwIList = SHBrowseForFolder(BI)
'    szPath = Space$(512)
'    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
'    DLG_BrowseFolder = X
'    If X Then
'        wPos = InStr(szPath, Chr(0))
'        Folder = Left$(szPath, wPos - 1)
'    Else
'        Folder = ""
'    End If
'
'End Function




