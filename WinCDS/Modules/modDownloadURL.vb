Module modDownloadURL
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Integer, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Integer, ByVal lpfnCB As Integer) As Integer
    Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Integer

    Public Function DownloadURLToString(ByVal URL As String, Optional ByRef FailureMessage As String = "") As String
        Dim lFN As String
        On Error Resume Next
        lFN = TempFile()
        If DownloadURLToFileAPI(URL, lFN, , FailureMessage) Then DownloadURLToString = ReadFile(lFN)
        Kill(lFN)
    End Function
    Public Function DownloadURLToFileAPI(ByVal URL As String, ByVal localFileName As String, Optional ByVal OverWrite As Boolean = True, Optional ByRef FailureMessage As String = "") As Boolean
        Dim ErrCode As Integer
        On Error Resume Next
        If OverWrite Then Kill(localFileName)
        If FileExists(localFileName) Then FailureMessage = "Could not secure destination." : Exit Function

        ErrCode = URLDownloadToFile(0, URL, localFileName, olelib.BINDF.BINDF_GETNEWESTVERSION, 0)
        DownloadURLToFileAPI = (ErrCode = 0) And FileExists(localFileName)
        If Not DownloadURLToFileAPI Then
            Select Case ErrCode
                Case 0 : FailureMessage = "Transfer Cancelled."
                Case -2146697212 : FailureMessage = "Download Failure.  Check URL: " & URL
                Case Else : FailureMessage = "Download Failure.  EC = " & ErrCode
            End Select
        End If
    End Function

    Public Function DownloadURLToFile(ByVal URL As String, ByVal localFileName As String, Optional ByVal ExpectedSize As Integer = -1, Optional ByRef FailureMessage As String = "", Optional ByVal AltPrg As Object = Nothing, Optional ByVal AltPrg2 As Object = Nothing, Optional ByVal APIOnly As Boolean = False) As Boolean
        Dim OK As Boolean, EC As Integer
        On Error Resume Next
        Kill(localFileName)
        If FileExists(localFileName) Then FailureMessage = "Could not secure destination." : Exit Function

        FailureMessage = ""
        OK = DoDownload(URL, localFileName, EC, AltPrg, AltPrg2, APIOnly)

        If OK Then
            DownloadURLToFile = True

            If Not FileExists(localFileName) Then
                DownloadURLToFile = False
                FailureMessage = "Transfer Cancelled."
            End If

            If ExpectedSize > 0 And FileLen(localFileName) <> ExpectedSize Then
                DownloadURLToFile = False
                FailureMessage = "Transfer Interrupted."
            End If

            If FileLen(localFileName) = 0 Then
                DownloadURLToFile = False
                FailureMessage = "No file."
            End If
        Else
            FailureMessage = "Download failed.  EC=" & EC
        End If
    End Function

    Private Function DoDownload(ByVal URL As String, ByVal localFileName As String, Optional ByRef ErrCode As Integer = 0, Optional ByVal AltPrg As Object = Nothing, Optional ByVal AltPrg2 As Object = Nothing, Optional ByVal APIOnly As Boolean = False) As Boolean
        Const BINDF_GETNEWESTVERSION As Integer = &H10

        If HasOleLib() And Not APIOnly Then
            DoDownload = DoDownloadProgress(URL, localFileName, ErrCode, AltPrg, AltPrg2)
        Else
            ErrCode = URLDownloadToFile(0, URL, localFileName, BINDF_GETNEWESTVERSION, 0)
            DoDownload = (ErrCode = 0)
        End If
    End Function

    Private Function HasOleLib() As Boolean
        HasOleLib = FileExists(GetWindowsSystemDir() & "\olelib.tlb")
    End Function

    Private Function DoDownloadProgress(ByVal URL As String, ByVal localFileName As String, Optional ByRef ErrCode As Integer = 0, Optional ByVal AltPrg As Object = Nothing, Optional ByVal AltPrg2 As Object = Nothing) As Boolean
        Dim C As URLDL
        C = New URLDL
        DoDownloadProgress = C.DownloadFileProgress(URL, localFileName, ErrCode, AltPrg, AltPrg2)
        C = Nothing
    End Function

End Module
