Attribute VB_Name = "modLink"

'引用 API
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

'链接到服务器
'参数 URL 为需要链接到的网页位置
Public Function OpenLink(ByVal URL As String) As Long
    
    Link = ShellExecute(0&, vbNullString, URL, _
           vbNullString, vbNullString, vbNormalFocus)
    
End Function

