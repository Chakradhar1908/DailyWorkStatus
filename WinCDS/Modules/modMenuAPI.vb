Module modMenuAPI
    Private Structure POINTAPI
        Dim X As Long
        Dim Y As Long
    End Structure
    Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
    Private Const WM_GETSYSMENU As Long = &H313

    Public Sub SystemMenuOnMouseUp(ByRef M As Form)
        Dim pt As POINTAPI
        GetCursorPos(pt) ' This is relative to the screen, so we can't use the coordinates passed in the event
        SendMessage(M.Handle, WM_GETSYSMENU, 0, MakeLong(pt.Y, pt.X))
    End Sub
End Module
