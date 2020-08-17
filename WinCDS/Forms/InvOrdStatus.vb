Public Class InvOrdStatus
    Dim ActiveMarginLine as integer
    Dim ActiveSaleNo As String
    Public ShowCost As Boolean   ' Set by MainMenu.
    Dim WithEvents mMailCheck As MailCheck

    Public Sub CheckDeliveryStatus()
        Show()
        'cmdNextSale.Value = True
        cmdNextSale.PerformClick()
    End Sub

End Class