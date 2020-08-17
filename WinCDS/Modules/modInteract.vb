Module modInteract
    Public Function WI_SendSale(ByVal SaleNo As String, Optional ByVal StoreNo As Integer = 0) As Boolean
        Dim FN As String, X As String
        FN = Format(StoreNo, "00") & "-" & SaleNo & ".html"
        X = WI_UploadStringAsFile("SendSale", SaleToHTML(SaleNo, StoreNo, , , False), FN)
    End Function

    Public Function WI_UploadStringAsFile(ByVal Operation As String, ByVal Text As String, Optional ByVal RemoteFileName As String = "") As String
        On Error Resume Next
        Dim Result As String
        UploadStringToURL(Text, WinCDSInteractiveURL(Operation), , RemoteFileName, Result)
        WI_UploadStringAsFile = Result
    End Function

    Private Function WinCDSInteractiveURL(ByVal Op As String) As String
        Dim A As String

        A = ""
        A = A & WebUpdateURL
        A = A & "Interact.asp"
        'A = A & "?a=" & CDbl(Now)
        A = A & "?a=" & Now
        A = A & "&s=" & ProtectValueForURL(Trim(LCase(StoreSettings(1).Name)))
        A = A & "&o=" & Op
        WinCDSInteractiveURL = URLEncode(A)
    End Function
End Module
