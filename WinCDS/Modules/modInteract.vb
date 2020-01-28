Module modInteract
    Public Function WI_SendSale(ByVal SaleNo As String, Optional ByVal StoreNo As Long) As Boolean
        Dim FN As String, X As String
        FN = Format(StoreNo, "00") & "-" & SaleNo & ".html"
        X = WI_UploadStringAsFile("SendSale", SaleToHTML(SaleNo, StoreNo, , , False), FN)
    End Function

End Module
