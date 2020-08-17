Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports Microsoft.Win32.SafeHandles
Imports System.Security
Imports System.Runtime.ConstrainedExecution
Imports System.Text
Module modVistaUAC
    Private Const TOKEN_QUERY As Integer = &H8
    Public Const TOKEN_DUPLICATE As Integer = 2
    Public Const TOKEN_IMPERSONATE = &H4
    'Private Const TokenElevation As Integer = 20
    Private Const TokenElevation As Integer = 2
    Const TOKEN_READ As Integer = &H20008
    'Private Declare Function GetCurrentProcess Lib "kernel32" () As Integer
    'Private Declare Function GetCurrentProcess Lib "kernel32" () As IntPtr
    <DllImport("kernel32.dll", SetLastError:=True)>
    Function GetCurrentProcess() As IntPtr
    End Function

    'Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Integer, ByVal DesiredAccess As Integer, TokenHandle As Integer) As Integer
    'Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As IntPtr, ByVal DesiredAccess As Integer, TokenHandle As IntPtr) As Integer
    <DllImport("advapi32.dll", SetLastError:=True)>
    Function OpenProcessToken(processHandle As IntPtr, desiredAccess As Integer, <Out()> ByRef tokenHandle As IntPtr) As Boolean
    End Function
    'Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Integer, ByVal TokenInformationClass As Integer, TokenInformation As Object, ByVal TokenInformationLength As Integer, ReturnLength As Integer) As Integer
    '<DllImport("advapi32.dll", SetLastError:=True)>
    'Function GetTokenInformation(tokenHandle As IntPtr,
    '                                    tokenInformationClass As Integer,
    '                                    ByVal tokenInformation As IntPtr,
    '                                    tokenInformationLength As Integer,
    '                                    ByRef returnLength As Integer) As Boolean
    'End Function
    <DllImport("advapi32.dll", SetLastError:=True)>
    Function GetTokenInformation(tokenHandle As IntPtr,
                                        tokenInformationClass As Integer,
                                        ByVal tokenInformation As IntPtr,
                                        tokenInformationLength As Integer,
                                        ByRef returnLength As Integer) As Boolean
    End Function
    'extern GetTokenInformation(
    '        TokenHandle intptr,
    '        TOKEN_INFORMATION_CLASS TokenInformationClass,
    '        IntPtr TokenInformation,
    '        int TokenInformationLength,
    '        out int ReturnLength);
    Public Function IsElevated(Optional ByVal hProcess As Integer = 0) As Boolean
        Dim hToken As IntPtr = IntPtr.Zero
        'Dim dwIsElevated As Integer
        Dim dwIsElevated As IntPtr = IntPtr.Zero
        Dim dwLength As Integer
        Dim hProcess2 As IntPtr
        'Dim p As Process
        Dim Result As Boolean

        If hProcess = 0 Then
            'hProcess = GetCurrentProcess()
            hProcess2 = GetCurrentProcess()
            'p = Process.GetCurrentProcess
            'p = Process.GetProcessById()
        End If
        'If OpenProcessToken(hProcess, TOKEN_QUERY, hToken) Then
        'If OpenProcessToken(hProcess2, TOKEN_QUERY, hToken) = True Then
        If hProcess2 <> IntPtr.Zero Then
            If OpenProcessToken(hProcess2, TOKEN_QUERY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE, hToken) <> 0 Then
                'If GetTokenInformation(hToken, TokenElevation, dwIsElevated, 4, dwLength) Then
                Result = GetTokenInformation(hToken, TokenElevation, dwIsElevated, dwLength, dwLength)
                dwIsElevated = Marshal.AllocHGlobal(dwLength)
                If GetTokenInformation(hToken, TokenElevation, dwIsElevated, dwLength, dwLength) = True Then
                    IsElevated = (dwIsElevated <> 0)
                End If
                CloseHandle(hToken)
            End If
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

    Public Function UACIsAdmin() As Boolean
        UACIsAdmin = IsUserAnAdmin <> 0
    End Function

End Module
