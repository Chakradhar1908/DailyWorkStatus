Public Class OrdVoid
    Public Function VoidOrder(ByVal SaleNo As String, Optional ByRef ParentForm As Form = Nothing) As Boolean
        ' This is the public access function.
        ' It shows the form modal, with an optional parent.
        ' OK and Cancel unload the form, which cancels the modal show.
        ' This function is then set to True if OK was clicked,
        ' or false if Cancel/Unload are called.
        ' The calling form calls: Success=OrdVoid.VoidOrder(SaleNo)

        'OrderVoided = False
        'VoidSaleNo = SaleNo
        'dteVoidDate.Value = Today
        'SaleTaxCode = 1 ' Default to default sales tax, in case the sale didn't include any?

        Dim Margin As CGrossMargin
        Margin = New CGrossMargin
        'If Margin.Load(SaleNo) Then
        ' The sale exists.  Load the payment detail into combo boxes.
        Do Until Margin.DataAccess.Record_EOF
                ' This block of hackishness compensates for old style Adjustments refunds.
                If Trim(Margin.Style) = "PAYMENT" Or (Trim(Margin.Style) = "NOTES" And Microsoft.VisualBasic.Left(Margin.Desc, 13) = "STORE FINANCE") Then
                    If Margin.Quantity = 0 Then
                        If Microsoft.VisualBasic.Left(Margin.Desc, 11) = "Refund By: " Then
                        Margin.SellPrice = -Math.Abs(Margin.SellPrice)
                        Select Case Trim(Mid(Margin.Desc, 12))
                                Case "CASH" : Margin.Quantity = 1
                                Case "CHECK" : Margin.Quantity = 2
                                Case "VISA CARD" : Margin.Quantity = 3
                                Case "MASTER CARD" : Margin.Quantity = 4
                                Case "DISCOVER CARD" : Margin.Quantity = 5
                                Case "AMEX CARD" : Margin.Quantity = 6
                                Case "DEBIT CARD" : Margin.Quantity = 9
                                Case "COMPANY CHECK" : Margin.Quantity = 2
                                Case Else : Margin.Quantity = 1 ' Error condition, treat as cash.
                            End Select
                        Else
                            Margin.Quantity = 1 ' Bad payment type!  Treat as cash.
                        End If
                    End If
                    ' End of Adjustment Refund hack block.
                    'If Margin.SellPrice = 0 And PayTypeIsFinance(Val(Margin.Quantity), False) Then
                    'Margin.SellPrice = BillOSale.SaleTotal
                    'End If
                    'AddPaymentLine Margin.Quantity, Margin.SellPrice
                End If
            'If Trim(Margin.Style) = "TAX1" Or Trim(Margin.Style) = "TAX2" Then
            ' Save the sale's tax code!
            'SaleTaxCode = Margin.Quantity
            'End If
            Margin.DataAccess.Records_MoveNext()
        Loop
        'optRefundType(0).Value = True  ' Default to Refund As Paid
        '    Me.Show vbModal
        'VoidOrder = OrderVoided
        If VoidOrder Then
                Dim Typ As String
                Typ = "" & Margin.Quantity
            'If SwipeCreditCards And IsIn(Typ, "3", "4", "5", "6") Then
            'BillOSale.cmdPrint.Value = True
        End If
        'End If
        'Else
        ' Void always fails if the sale has no margin lines.
        VoidOrder = False
        'End If
        DisposeDA(Margin)
    End Function

End Class