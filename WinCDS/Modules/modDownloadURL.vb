
Module modDownloadURL

    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller as integer, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved as integer, ByVal lpfnCB as integer) as integer
    Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) as integer
    Public Function DownloadURLToString(ByVal URL As String, Optional ByRef FailureMessage As String = "") As String
        Dim lFN As String
        On Error Resume Next
        lFN = TempFile()
        If DownloadURLToFileAPI(URL, lFN, , FailureMessage) Then DownloadURLToString = ReadFile(lFN)
        Kill(lFN)
    End Function
    Public Function DownloadURLToFileAPI(ByVal URL As String, ByVal localFileName As String, Optional ByVal OverWrite As Boolean = True, Optional ByRef FailureMessage As String = "") As Boolean
        Dim ErrCode as integer
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

End Module
