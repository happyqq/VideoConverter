Attribute VB_Name = "modPrivate"

Public Function CheckRegister(ByVal RegUser As String, ByVal RegCode As String, ByVal bShowMsg As Boolean) As Boolean

CheckRegister = False

Dim cMD5 As New clsMD5

 If RegCode = UCase(cMD5.CalculateMD5(sMD5Pre & RegUser & sMD5Last)) Then
 
    CheckRegister = True
    
    If bShowMsg Then MsgBox "���ɶ������ݳɹ�", vbOKOnly + vbInformation, "��ϲ"
     
      
 Else
     CheckRegister = False
 
     If bShowMsg Then MsgBox "��Ȩ��������ɶ�������ʧ�ܣ����������������ȥ��ϵ���߰ɡ�", vbOKOnly + vbCritical, "��ѽ"
     
 
 End If
 
End Function



Public Function CheckContactURL(ByVal sURL As String, ByVal RegCode As String) As Boolean

CheckContactURL = False

Dim cMD5 As New clsMD5

 If sURL = sADUrl Then
     CheckContactURL = True
     Exit Function
 End If

 If RegCode = UCase(cMD5.CalculateMD5(sMD5Pre & sURL & sMD5Last)) And left(sURL, 22) = "http://www.mama520.cn/" Then
    CheckContactURL = True
 Else
     CheckContactURL = False
 
 End If
 
End Function



