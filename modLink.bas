Attribute VB_Name = "modLink"

'���� API
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

'���ӵ�������
'���� URL Ϊ��Ҫ���ӵ�����ҳλ��
Public Function OpenLink(ByVal URL As String) As Long
    
    Link = ShellExecute(0&, vbNullString, URL, _
           vbNullString, vbNullString, vbNormalFocus)
    
End Function

