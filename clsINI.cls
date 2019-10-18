VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetPrivateProfileSection Lib "Kernel32" Alias _
 "GetPrivateProfileSectionA" (ByVal Section As String, ByVal Buffer As String, ByVal Size As Long, ByVal Filename As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "Kernel32" Alias _
 "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias _
 "GetPrivateProfileStringA" (ByVal Section As String, ByVal Key As String, ByVal Default As String, ByVal Buffer As String, ByVal Size As Long, ByVal Filename As String) As Long
Private Declare Function WritePrivateProfileSection Lib "Kernel32" Alias _
 "WritePrivateProfileSectionA" (ByVal Section As String, ByVal Key As String, ByVal Filename As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias _
 "WritePrivateProfileStringA" (ByVal Section As String, ByVal Key As String, ByVal Setting As String, ByVal Filename As String) As Long

Private iFilename As String, iSection As String, iKey As String, _
 iErrCode As Long, iErrVar(1 To 2) As String

Public Property Let Key(ByVal vData As String)
 iKey = vData
End Property

Public Property Get Key() As String
 Key = iKey
End Property

Public Property Let Section(ByVal vData As String)
 iSection = vData
End Property

Public Property Get Section() As String
 Section = iSection
End Property

Public Property Let Filename(ByVal vData As String)
 iFilename = Trim$(vData)
End Property

Public Property Get Filename() As String
'Attribute Filename.VB_UserMemId = 0
 Filename = iFilename
End Property

Public Function CheckIniFile(Optional Filename As String) As Boolean
 Dim tFilename As String
 CheckIniFile = False
 tFilename = GetFilename(Filename)
 If Len(tFilename) = 0 Then
  Exit Function
 End If
 If Len(Dir(tFilename)) = 0 Then
  Exit Function
 End If
 CheckIniFile = True
End Function

Public Function GetAllSectionsFromInifile(Optional ByVal Filename As String, Optional Sorted As Boolean = False) As Collection
 Dim tFilename As String, i As Long, j As Long, nBuffer As String, _
  res As Long, tStr As String, nItems() As String

 Const kBufferSize = 32767

 tFilename = GetFilename(Filename)

 If CheckIniFile(tFilename) = True Then
  Set GetAllSectionsFromInifile = New Collection
  nBuffer = Space$(kBufferSize)
  res = GetPrivateProfileSectionNames(nBuffer, kBufferSize, tFilename & vbNullString)
  If res = 0 Then
   Exit Function
  End If
  nItems = Split(Left$(nBuffer, res), Chr$(0))
  If Sorted = True Then
   For i = LBound(nItems) To UBound(nItems) - 1
    For j = i To UBound(nItems)
     If StrComp(nItems(i), nItems(j), vbTextCompare) = 1 Then
      tStr = nItems(i): nItems(i) = nItems(j): nItems(j) = tStr
     End If
    Next j
   Next i
  End If
  For i = 0 To UBound(nItems)
   If Trim$(nItems(i)) <> "" Then
    GetAllSectionsFromInifile.Add nItems(i)
   End If
  Next i
 End If
End Function

Public Function CheckSection(Optional Section As String, Optional Filename As String) As Boolean
 Dim tFilename As String, tSection As String, i As Long, tColl As Collection
 CheckSection = False
 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 Set tColl = GetAllSectionsFromInifile(tFilename)
 For i = 1 To tColl.Count
  If UCase$(tSection) = UCase$(tColl.Item(i)) Then
   CheckSection = True
   Set tColl = Nothing
   Exit Function
  End If
 Next i
 Set tColl = Nothing
End Function

Public Function GetAllKeysFromSection(Optional Section As String, Optional Filename As String, Optional WithComments As Boolean = False, Optional Sorted As Boolean = False) As Collection
 Const kBufferSize = 32767
 Dim tFilename As String, tSection As String, nBuffer As String, res As Long, _
  nItems() As String, nItem() As String, i As Long, tColl As Collection, _
  ts(1) As String, j As Long, tStr As String

 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)

 Set tColl = New Collection
 Set GetAllKeysFromSection = tColl
 If CheckSection(tSection, tFilename) = True Then
  nBuffer = Space$(kBufferSize)
  res = GetPrivateProfileSection(tSection, nBuffer, kBufferSize, tFilename)
  If res = 0 Then
   Exit Function
  End If
  nItems = Split(Left$(nBuffer, res), Chr$(0))
  If Sorted = True Then
   For i = LBound(nItems) To UBound(nItems) - 1
    For j = i To UBound(nItems)
     If StrComp(nItems(i), nItems(j), vbTextCompare) = 1 Then
      tStr = nItems(i): nItems(i) = nItems(j): nItems(j) = tStr
     End If
    Next j
   Next i
  End If
  For i = LBound(nItems) To UBound(nItems)
   If Len(Trim$(nItems(i))) > 0 Then
    If WithComments = True Then
      If InStr(nItems(i), "=") > 0 Then
        nItem = Split(nItems(i), "=")
        ts(0) = nItem(0): ts(1) = nItem(1)
        tColl.Add ts
       Else
        ts(0) = nItems(i): ts(1) = vbNullString
        tColl.Add ts
      End If
     Else
      If InStr(Trim$(nItems(i)), ";") <> 1 Then
       If InStr(nItems(i), "=") > 0 Then
         nItem = Split(nItems(i), "=")
         ts(0) = nItem(0): ts(1) = nItem(1)
         tColl.Add ts
        Else
         ts(0) = nItems(i): ts(1) = vbNullString
         tColl.Add ts
       End If
      End If
    End If

   End If
  Next i
  Set GetAllKeysFromSection = tColl
 End If
End Function

Public Function CheckKey(Optional Key As String, Optional Section As String, Optional Filename As String) As Boolean
 Dim tFilename As String, tSection As String, tKey As String, _
  i As Long, tColl As Collection
 CheckKey = False
 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 tKey = GetKeyname(Key)

 Set tColl = GetAllKeysFromSection(tSection, tFilename)
 For i = 1 To tColl.Count
  If UCase$(tKey) = UCase$(tColl.Item(i)(0)) Then
   CheckKey = True
   Set tColl = Nothing
   Exit Function
  End If
 Next i
End Function

Public Function GetKeyFromSection(Optional Key As String, Optional Section As String, Optional Filename As String) As String
 Const kBufferSize = 1024
 Dim nBuffer As String, tKey As String, tSection As String, tFilename As String, _
  tColl As Collection, i As Long, res As Long
 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 tKey = GetKeyname(Key)

 GetKeyFromSection = vbNullString
 If CheckSection(tSection, tFilename) = True Then
  Set tColl = GetAllKeysFromSection(tSection, tFilename)
  For i = 1 To tColl.Count
   If UCase$(tKey) = UCase$(tColl.Item(i)(0)) Then
    GetKeyFromSection = tColl.Item(i)(1)
    Set tColl = Nothing
    Exit Function
   End If
  Next i
 End If
End Function

Public Sub DeleteIniFile(Optional Filename As String)
 Dim tFilename As String
 tFilename = GetFilename(Filename)
 If Dir(tFilename) <> "" Then
  Kill tFilename
 End If
End Sub

Public Sub CreateIniFile(Optional Filename As String)
 Dim tFilename As String, fn As Long
 tFilename = GetFilename(Filename)

 fn = FreeFile
 Open tFilename For Output As #fn
 Close #fn
 FlushInifile
End Sub

Public Sub CreateSection(Optional Section As String, Optional Filename As String)
 Dim tFilename As String, tSection As String, res As Long, fn As Long

 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)

 If CheckSection(tSection, tFilename) = False Then
  fn = FreeFile
  Open tFilename For Append As #fn
  Print #fn, "[" & tSection & "]"
  Close #fn
 End If
 FlushInifile
End Sub

Public Sub DeleteAllSectionFromInifile(Optional Filename As String)
 Dim tKey As String, tSection As String, tFilename As String, res As Long, _
  i As Long, tColl As Collection

 tFilename = GetFilename(Filename)

 If CheckIniFile(tFilename) = True Then
  res = WritePrivateProfileString(vbNullString, vbNullString, vbNullString, tFilename)
 End If
End Sub

Public Sub DeleteSectionFromInifile(Optional Section As String, Optional Filename As String)
 Dim tKey As String, tSection As String, tFilename As String, res As Long, _
  i As Long, tColl As Collection

 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)

 If CheckSection(tSection, tFilename) = True Then
  res = WritePrivateProfileString(tSection, vbNullString, vbNullString, tFilename)
 End If
End Sub

Public Sub DeleteKeyFromSection(Optional Key As String, Optional Section As String, Optional Filename As String)
 Dim tKey As String, tSection As String, tFilename As String, res As Long

 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 tKey = GetKeyname(Key)

 If CheckKey(tKey, tSection, tFilename) = True Then
  res = WritePrivateProfileString(tSection, tKey, vbNullString, tFilename)
 End If
End Sub

Public Function SaveKey(Value As String, Optional Key As String, Optional Section As String, Optional Filename As String) As Boolean
 Dim tKey As String, tSection As String, tFilename As String, res As Long

 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 tKey = GetKeyname(Key)

 res = WritePrivateProfileString(tSection & vbNullString, tKey & vbNullString, _
  Value & vbNullString, tFilename & vbNullString)
 FlushInifile
End Function

Public Sub FlushInifile(Optional Filename As String)
 Dim res As Long, tFilename As String
 tFilename = GetFilename(Filename)
 res = WritePrivateProfileString(vbNullString, vbNullString, vbNullString, tFilename & vbNullString)
End Sub

Private Function GetFilename(ByVal Filename As String) As String
 Filename = Trim$(Filename)
 If Len(Filename) = 0 Then
   GetFilename = iFilename
  Else
   GetFilename = Filename
 End If
End Function

Private Function GetSectionname(ByVal Sectionname As String) As String
 Sectionname = Trim$(Sectionname)
 If Len(Sectionname) = 0 Then
   GetSectionname = iSection
  Else
   GetSectionname = Sectionname
 End If
End Function

Private Function GetKeyname(ByVal Keyname As String) As String
 Keyname = Trim$(Keyname)
 If Len(Keyname) = 0 Then
   GetKeyname = iKey
  Else
   GetKeyname = Keyname
 End If
End Function


