Attribute VB_Name = "modSound"
Private Declare Function sndPlaySoundA Lib "winmm.dll" ( _
    ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub PlaySound(ByVal sPath As String)
    Call sndPlaySoundA(sPath, 1&)
End Sub

