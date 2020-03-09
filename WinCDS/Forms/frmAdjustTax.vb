Public Class frmAdjustTax
    Public Sub LoadSale(ByVal vSaleNo As String)
        Dim I As Long, S As String, R As Double, X As Currency, C As Boolean, TT As Currency
        Dim xH As cHolding
  Set xH = New cHolding
  If Not xH.Load(vSaleNo, "LeaseNo") Then
            DisposeDA xH
    Exit Sub
        End If

        SaleNo = vSaleNo
        GrossSale = xH.Sale
        Taxable = GetPrice(xH.Sale) - GetPrice(xH.NonTaxable)

        lstTaxes.Clear
        R = StoreSettings.SalesTax
        X = OnScreenReport.SaleTax1Amount
        If X = 0 Then
            X = R * Taxable
            C = False
        Else
            C = True
        End If

        TT = X
        S = AlignString(Format(R, "0.000"), 11, vbAlignLeft, True)
        S = S & " " & FormatCurrency(X)
        lstTaxes.AddItem S
  lstTaxes.itemData(lstTaxes.NewIndex) = 0
        lstTaxes.Selected(lstTaxes.NewIndex) = C

        For I = 0 To SalesTax2Count() - 1
            R = QuerySalesTax2Rate(I)
            X = OnScreenReport.SaleTax2Amount(I)
            If X = 0 Then
                X = R * Taxable
                C = False
            Else
                C = True
            End If
            TT = TT + X
            S = AlignString(QuerySalesTax2(I), 11, vbAlignLeft, True)
            S = S & " " & FormatCurrency(X)
            lstTaxes.AddItem S
    lstTaxes.itemData(lstTaxes.NewIndex) = I + 1
            lstTaxes.Selected(lstTaxes.NewIndex) = C
        Next

        DisposeDA xH
End Sub

End Class