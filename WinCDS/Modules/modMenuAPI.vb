Module modMenuAPI
    Public Sub SystemMenuOnMouseUp(ByRef M As Form)
        Dim pt As POINTAPI
        GetCursorPos pt ' This is relative to the screen, so we can't use the coordinates passed in the event
        SendMessage M.hWnd, WM_GETSYSMENU, 0, ByVal MakeLong(pt.Y, pt.X)
End Sub

End Module
