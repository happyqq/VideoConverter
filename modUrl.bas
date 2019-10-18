Attribute VB_Name = "modUrl"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, ByVal lpParameters As String, _
  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_NORMAL = 1 '����ЩAPI����������VB�������棬����vbNormalFocus��
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_SHOW = 5


Public Sub OpenURL(hWnd As Long, sURL As String)

    ShellExecute hWnd, "open", sURL, vbNullString, vbNullString, SW_SHOW
    
End Sub

