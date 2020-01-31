Module modVistaUAC
    Private Const TOKEN_QUERY As Integer = &H8
    Private Const TokenElevation As Integer = 20
    Private Declare Function GetCurrentProcess Lib "kernel32" () As Integer
    Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Integer, ByVal DesiredAccess As Integer, TokenHandle As Integer) As Integer
    Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Integer, ByVal TokenInformationClass As Integer, TokenInformation As Object, ByVal TokenInformationLength As Integer, ReturnLength As Integer) As Integer
    Public Function IsElevated(Optional ByVal hProcess As Integer = 0) As Boolean
        Dim hToken As Integer
        Dim dwIsElevated As Integer
        Dim dwLength As Integer

        If hProcess = 0 Then
            hProcess = GetCurrentProcess()
        End If
        If OpenProcessToken(hProcess, TOKEN_QUERY, hToken) Then
            If GetTokenInformation(hToken, TokenElevation, dwIsElevated, 4, dwLength) Then
                IsElevated = (dwIsElevated <> 0)
            End If
            CloseHandle(hToken)
        End If
    End Function

    Public Function LaunchAutoVNC() As Boolean
        Dim cPath As String

        '  If MsgBox("This support feature temporarily disables your User Account Control (UAC) setting in order to allow " & CompanyName & " to have full control of your computer." & vbCrLf & "During the support session, the UAC will not be active." & vbCrLf2 & "Click OK to indicate you are aware of this.", vbExclamation + vbOKCancel) = vbCancel Then
        '    Exit Function
        '  End If
        '
        cPath = CurDir()

        ChDrive(WinCDSAutoVNCFolder)
        ChDir(WinCDSAutoVNCFolder)
        MainMenu.Hide()

        ConnectCMDUpgrade()

        ShellOut.ShellOut(ConnectCMDFile)

        MainMenu.Show()
        ChDrive(cPath)
        ChDir(cPath)

        LaunchAutoVNC = True
    End Function

End Module
