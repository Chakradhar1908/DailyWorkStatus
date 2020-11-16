Module modSound
    Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpsStyleoundName As String, ByVal uFlags As Integer) As Integer
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10

    Public Sub PlayIt(ByVal T As String)
        sndPlaySound(T, SND_ASYNC Or SND_NODEFAULT)
    End Sub

End Module
