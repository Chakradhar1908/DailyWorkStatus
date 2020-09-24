Public Class cFTP
    Private Const DBL_DAYZEROBIAS As Double = 109205.0#   ' Abs(CDbl(#01-01-1601#))
    Private Const DBL_MILLISECONDPERDAY As Double = 10000000.0# * 60.0# * 60.0# * 24.0# / 10000.0#
    Private Const BUFFERSIZE As Integer = 4 * 1024

    Private m_hOpen As Integer
    Private m_hConnection As Integer
    Private m_dwType As Integer
    Private m_dwSeman As Integer
    Private m_sErrorMessage As String
    Private m_sErrorSource As String
    Private Const MAX_PATH As Integer = 260
    Private Const INTERNET_FLAG_RELOAD As Integer = &H80000000
    Private Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
    Private Const INTERNET_FLAG_PASSIVE As Integer = &H8000000
    Private Const FORMAT_MESSAGE_FROM_HMODULE As Integer = &H800
    Private Const GENERIC_READ As Integer = &H80000000
    Private Const GENERIC_WRITE As Integer = &H40000000
    Private Const ERROR_NO_MORE_FILES As Integer = 18
    Private Const INTERNET_AUTODIAL_FORCE_ONLINE As Integer = 1
    Private Const INTERNET_OPEN_TYPE_PRECONFIG As Integer = 0
    Private Const INTERNET_INVALID_PORT_NUMBER As Integer = 0
    Private Const INTERNET_SERVICE_FTP As Integer = 1
    Private Const FTP_TRANSFER_TYPE_BINARY As Integer = &H2
    Private Const FTP_TRANSFER_TYPE_ASCII As Integer = &H1
    Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Integer, ByVal dwReserved As Integer) As Integer
    Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Integer, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Integer) As Integer
    Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Integer, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Integer, ByVal lFlags As Integer, ByVal lContext As Integer) As Integer
    Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Integer) As Integer
    Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Integer, ByVal sBuff As String, ByVal Access As Integer, ByVal Flags As Integer, ByVal Context As Integer) As Integer
    Private Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Integer, ByRef sBuffer As Byte, ByVal lNumBytesToWrite As Integer, dwNumberOfBytesWritten As Integer) As Integer
    Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean
    Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszName As String) As Boolean
    Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Integer, ByVal lpszErrorBuffer As String, ByRef lpdwErrorBufferLength As Integer) As Boolean
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByVal lpSource As Integer, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, Arguments As Integer) As Integer
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Integer
    <VBFixedString(2048)> Dim sBuffer As String
    Public Sub SetTransferBinary()
        m_dwType = FTP_TRANSFER_TYPE_BINARY
    End Sub

    Public Sub SetModePassive()
        m_dwSeman = INTERNET_FLAG_PASSIVE
    End Sub

    Public Function OpenConnection(ByVal sServer As String, ByVal sUser As String, ByVal sPassword As String, Optional ByVal Port As Integer = INTERNET_INVALID_PORT_NUMBER) As Boolean
        CloseConnection()

        If InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0) = 0 Then
            pvSetLastError(Err.LastDllError, "OpenConnection:InternetAutodial")
        End If
        m_hOpen = InternetOpen(STR_APP_NAME, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        If m_hOpen = 0 Then
            pvSetLastError(Err.LastDllError, "OpenConnection:InternetOpen")
            Exit Function
        End If
        m_hConnection = InternetConnect(m_hOpen, sServer, Port, sUser, sPassword, INTERNET_SERVICE_FTP, m_dwSeman, 0)
        If m_hConnection = 0 Then
            pvSetLastError(Err.LastDllError, "OpenConnection:InternetConnect")
            InternetCloseHandle(m_hOpen)
            m_hOpen = 0
            Exit Function
        End If
        '--- success
        OpenConnection = True
    End Function

    Public Function UploadFile(ByVal sLocal As String, ByVal sRemote As String) As Boolean
        Dim hFile As Integer
        Dim baData() As Byte
        Dim lWritten As Integer
        Dim lSize As Integer
        Dim lSum As Integer
        Dim lIdx As Integer
        Dim NFile As Integer
        Dim bCancel As Boolean

        hFile = FtpOpenFile(m_hConnection, sRemote, GENERIC_WRITE, m_dwType, 0)
        If hFile = 0 Then
            pvSetLastError(Err.LastDllError, "UploadFile:FtpOpenFile")
            Exit Function
        End If
        NFile = FreeFile()
        'Open sLocal For Binary Access Read As #NFile
        FileOpen(NFile, sLocal, OpenMode.Binary)
        ReDim baData(0 To BUFFERSIZE - 1)
        lSize = LOF(NFile)
        For lIdx = 0 To lSize \ BUFFERSIZE
            lWritten = lSize - lIdx * BUFFERSIZE
            If lWritten <= 0 Then
                Exit For
            End If
            If lWritten < BUFFERSIZE Then
                ReDim baData(0 To lWritten - 1)
            End If
            'Get #NFile, , baData
            'baData = My.Computer.FileSystem.ReadAllText(sLocal)
            baData = My.Computer.FileSystem.ReadAllBytes(sLocal)

            If InternetWriteFile(hFile, baData(0), UBound(baData) + 1, lWritten) = 0 Then
                pvSetLastError(Err.LastDllError, "UploadFile:InternetWriteFile")
                GoTo QH
            End If
            lSum = lSum + lWritten
            'RaiseEvent (FileTransferProgress(lSum, lSize, bCancel))
            If bCancel Then
                GoTo QH
            End If
        Next
        '--- success
        UploadFile = True
QH:
        'Close #NFile
        FileClose(NFile)
        InternetCloseHandle(hFile)
    End Function

    Public Sub CloseConnection()
        If m_hConnection <> 0 Then
            InternetCloseHandle(m_hConnection)
            m_hConnection = 0
        End If
        If m_hOpen <> 0 Then
            InternetCloseHandle(m_hOpen)
            m_hOpen = 0
        End If
    End Sub

    Public Function SetDirectory(ByVal sDir As String) As Boolean
        If FtpSetCurrentDirectory(m_hConnection, sDir) = 0 Then
            pvSetLastError(Err.LastDllError, "SetDirectory:FtpSetCurrentDirectory")
            Exit Function
        End If
        '--- success
        SetDirectory = True
    End Function

    Public Function CreateDirectory(ByVal sDirectory As String) As Boolean
        If FtpCreateDirectory(m_hConnection, sDirectory) = 0 Then
            pvSetLastError(Err.LastDllError, "CreateDirectory:FtpCreateDirectory")
            Exit Function
        End If
        '--- success
        CreateDirectory = True
    End Function

    Private Sub pvSetLastError(ByVal dwError As Integer, ByRef sFunc As String)
        'Dim sBuffer As String * 2048

        m_sErrorSource = sFunc
        If dwError = 12003 Then
            ' Extended error information was returned
            InternetGetLastResponseInfo(dwError, sBuffer, 2048)
        Else
            FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, GetModuleHandle("wininet.dll"), dwError, 0, sBuffer, 2048, 0)
        End If
        If InStr(sBuffer, ChrW(0)) > 0 Then
            m_sErrorMessage = Replace(Left(sBuffer, InStr(sBuffer, ChrW(0)) - 1), vbCrLf, "") & " (" & dwError & ")"
        Else
            m_sErrorMessage = "Error " & dwError
        End If
    End Sub

    Private Function STR_APP_NAME() As String
        'STR_APP_NAME = App.ProductName
        STR_APP_NAME = Application.ProductName
    End Function
End Class
