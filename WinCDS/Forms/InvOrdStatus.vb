Public Class InvOrdStatus
    Dim ActiveMarginLine As Long
    Dim ActiveSaleNo As String
    Public ShowCost As Boolean   ' Set by MainMenu.
    Dim WithEvents mMailCheck As MailCheck

    Public Sub CheckDeliveryStatus()
        Show()
        cmdNextSale.Value = True
    End Sub

End Class