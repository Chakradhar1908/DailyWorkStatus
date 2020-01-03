Module modVistaUAC
    Public Function IsElevated(Optional ByVal hProcess As Long) As Boolean
        Dim hToken As Long
        Dim dwIsElevated As Long
        Dim dwLength As Long

        If hProcess = 0 Then
            hProcess = GetCurrentProcess()
        End If
        If OpenProcessToken(hProcess, TOKEN_QUERY, hToken) Then
            If GetTokenInformation(hToken, TokenElevation, dwIsElevated, 4, dwLength) Then
                IsElevated = (dwIsElevated <> 0)
            End If
            CloseHandle hToken
  End If
    End Function

End Module
