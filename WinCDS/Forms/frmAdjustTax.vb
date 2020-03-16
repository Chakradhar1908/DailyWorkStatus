Public Class frmAdjustTax
    Public Sub LoadSale(ByVal vSaleNo As String)
        Dim I As Integer, S As String, R As Double, X As Decimal, C As Boolean, TT As Decimal
        Dim xH As cHolding
        xH = New cHolding
        If Not xH.Load(vSaleNo, "LeaseNo") Then
            DisposeDA(xH)
            Exit Sub
        End If

        SaleNo = vSaleNo
        GrossSale = xH.Sale
        Taxable = GetPrice(xH.Sale) - GetPrice(xH.NonTaxable)

        'lstTaxes.Clear
        lstTaxes.Items.Clear()
        R = StoreSettings.SalesTax
        X = OnScreenReport.SaleTax1Amount
        If X = 0 Then
            X = R * Taxable
            C = False
        Else
            C = True
        End If

        TT = X
        S = AlignString(Format(R, "0.000"), 11, VBRUN.AlignConstants.vbAlignLeft, True)
        S = S & " " & FormatCurrency(X)
        'lstTaxes.AddItem S
        'lstTaxes.itemData(lstTaxes.NewIndex) = 0
        'lstTaxes.Selected(lstTaxes.NewIndex) = C
        Dim AddedItem As Integer
        AddedItem = lstTaxes.Items.Add(New ItemDataClass(S, 0))
        lstTaxes.SetSelected(AddedItem, True)

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
            S = AlignString(QuerySalesTax2(I), 11, VBRUN.AlignConstants.vbAlignLeft, True)
            S = S & " " & FormatCurrency(X)
            'lstTaxes.AddItem S
            'lstTaxes.itemData(lstTaxes.NewIndex) = I + 1
            'lstTaxes.Selected(lstTaxes.NewIndex) = C
            AddedItem = lstTaxes.Items.Add(New ItemDataClass(S, I + 1))
            lstTaxes.SetSelected(AddedItem, True)
        Next

        DisposeDA(xH)
    End Sub

    Public Property SaleNo() As String
        Get
            SaleNo = txtSaleNo.Text
        End Get
        Set(value As String)
            txtSaleNo.Text = value
        End Set
    End Property

    Public Property GrossSale() As Decimal
        Get
            GrossSale = GetPrice(txtGrossSale.Text)
        End Get
        Set(value As Decimal)
            txtGrossSale.Text = FormatCurrency(value)
        End Set
    End Property

    Public Property Taxable() As Decimal
        Get
            Taxable = GetPrice(txtTaxable.Text)
        End Get
        Set(value As Decimal)
            txtTaxable.Text = FormatCurrency(value)
        End Set
    End Property
End Class