Public Class OrdVoid
    Private VoidSaleNo As String
    Private SaleTaxCode As Integer
    Private AddedVoidLine As Boolean
    Private PaymentCount As Integer    ' Count of payment types used in the sale
    Private OrderVoided As Boolean  ' Confirmation that the order was voided.
    Private OrdVoidFormLoad As Boolean
    Private Enum cdsRefundType
        cdsrft_AsPaid = 0
        cdsrft_CompanyCheck = 1
        cdsrft_StoreCredit = 2
        cdsrft_ForfeitDeposit = 3
        cdsrft_ApplyToSale = 4
    End Enum

    Public Function VoidOrder(ByVal SaleNo As String, Optional ByRef ParentForm As Form = Nothing) As Boolean
        ' This is the public access function.
        ' It shows the form modal, with an optional parent.
        ' OK and Cancel unload the form, which cancels the modal show.
        ' This function is then set to True if OK was clicked,
        ' or false if Cancel/Unload are called.
        ' The calling form calls: Success=OrdVoid.VoidOrder(SaleNo)
        '<CT>
        Dim StyleValue As String = ""
        '</CT>

        OrderVoided = False
        VoidSaleNo = SaleNo
        dteVoidDate.Value = Today
        If OrdVoidFormLoad = False Then
            OrdVoid_Load(Me, New EventArgs)
            OrdVoidFormLoad = True
        End If
        SaleTaxCode = 1 ' Default to default sales tax, in case the sale didn't include any?

        Dim Margin As CGrossMargin
        Margin = New CGrossMargin
        If Margin.Load(SaleNo) Then
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
                    If Margin.SellPrice = 0 And PayTypeIsFinance(Val(Margin.Quantity), False) Then
                        Margin.SellPrice = BillOSale.SaleTotal
                    End If
                    AddPaymentLine(Margin.Quantity, Margin.SellPrice)
                End If
                If Trim(Margin.Style) = "TAX1" Or Trim(Margin.Style) = "TAX2" Then
                    ' Save the sale's tax code!
                    SaleTaxCode = Margin.Quantity
                End If

                Margin.DataAccess.Records_MoveNext()
                Margin.cDataAccess_GetRecordSet(Margin.DataAccess.RS)
            Loop
            'optRefundType(0).Value = True  ' Default to Refund As Paid
            optRefundType0.Checked = True
            'Me.Show vbModal
            '<CT>
            StyleValue = BillOSale.UGridIO1.GetValue(BillOSale.UGridIO1.LastRowUsed, BillColumns.eStyle)
            BillOSale.SetStyle(BillOSale.UGridIO1.LastRowUsed, StyleValue)
            '</CT>
            MailCheck.Close()
            Me.ShowDialog()
            VoidOrder = OrderVoided

            If VoidOrder Then
                Dim Typ As String
                Typ = "" & Margin.Quantity
                If SwipeCreditCards() And IsIn(Typ, "3", "4", "5", "6") Then
                    'BillOSale.cmdPrint.Value = True
                    BillOSale.cmdPrint.PerformClick()
                End If
            End If
        Else
            ' Void always fails if the sale has no margin lines.
            VoidOrder = False
        End If
        DisposeDA(Margin)
    End Function

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        'Unload Me
        Me.Close()
    End Sub

    Private ReadOnly Property RefundType() As cdsRefundType
        Get
            'For RefundType = optRefundType.LBound To optRefundType.UBound
            '    If optRefundType(RefundType) Then Exit Property
            'Next
            If optRefundType0.Checked = True Then Exit Property
            If optRefundType1.Checked = True Then Exit Property
            If optRefundType2.Checked = True Then Exit Property
            If optRefundType3.Checked = True Then Exit Property
            If optRefundType4.Checked = True Then Exit Property
        End Get
    End Property

    Private Function DescribeVoidType(ByVal PaidAs As String, Optional ByVal Forfeit As Decimal = 0) As String
        Dim F As String
        If Forfeit > 0 Then F = " - Fft: " & FormatCurrency(Forfeit)
        Select Case RefundType
            Case cdsRefundType.cdsrft_AsPaid : DescribeVoidType = "(As Paid: " & PaidAs & F & ")"
            Case cdsRefundType.cdsrft_CompanyCheck : DescribeVoidType = "(Company Check" & F & ")"
            Case cdsRefundType.cdsrft_StoreCredit : DescribeVoidType = "(Store Credit" & F & ")"
            Case cdsRefundType.cdsrft_ForfeitDeposit : DescribeVoidType = "(Forfeit " & FormatCurrency(Forfeit) & ")"
            Case cdsRefundType.cdsrft_ApplyToSale : DescribeVoidType = "(" & FormatCurrency(GetPrice(txtRefundSpecial.Text)) & " Applied to Sale No " & txtApplyToSaleNo.Text & F & ")"
        End Select
    End Function

    Private Function SaleHasBlindCredit(ByVal SaleNo As String, ByRef TotalAmount As Decimal) As Boolean
        Dim S As sSale, I As Integer
        S = New sSale
        S.LoadSaleNo(SaleNo)
        For I = 1 To S.ItemCount
            If S.Item(I).Style = "PAYMENT" Then
                If SwipeCreditCards() And Val(S.Item(I).Quantity) = 3 Then
                    TotalAmount = TotalAmount + S.Item(I).Price
                    If GetPrice(S.Item(I).Price) < 0 Then
                        SaleHasBlindCredit = True
                    End If
                End If
            End If
        Next

Done:
        DisposeDA(S)
    End Function

    Private Sub cmdOk_Click(sender As Object, e As EventArgs) Handles cmdOk.Click
        Dim Holding As cHolding, HoldingLoaded As Boolean, PaidAs As String, TotalRefunded As Boolean, TotalAmt As Decimal
        Dim Written As Decimal, Tax As Decimal, tST As String, Appr As String
        Dim Margin As CGrossMargin, Count As Integer, Typ As String
        Dim M As String

        ' Validate and perform the void.
        If GetPrice(lblRefundTotal.Text) > GetPrice(lblTotalPaid.Text) Then
            MessageBox.Show("You can't refund more than the total deposit.", "WinCDS")
            Exit Sub
        End If
        '  If RefundType = cdsrft_ForfeitDeposit And GetPrice(lblRefundTotal.Caption) > 0 Then
        '    MsgBox "The entire deposit needs to be forfeit.", vbCritical
        '    Exit Sub
        '  End If
        If GetPrice(lblRefundTotal.Text) < 0 Then
            M = "Do you really want to refund a negative amount?" & vbCrLf & "This means you are taking money from the customer!"
            If MessageBox.Show(M, "WinCDS", MessageBoxButtons.YesNo) = DialogResult.No Then Exit Sub
        End If
        If GetPrice(lblRefundTotal.Text) < GetPrice(lblTotalPaid.Text) Then
            MessageBox.Show("You have to refund the whole deposit.", "WinCDS")
            Exit Sub
        End If

        If RefundType = cdsRefundType.cdsrft_ApplyToSale Then
            If txtApplyToSaleNo.Text = "" Then MessageBox.Show("Enter a sale number into the box.", "No Sale number") : Exit Sub
            If Not LeaseNoExists(txtApplyToSaleNo.Text) Then MessageBox.Show("That sale number does not exist.", "Invalid Sale Number") : Exit Sub
            If Not IsIn(GetLeaseNoStatus(txtApplyToSaleNo.Text), "OPEN", "Lay-A-Way", "30 Day Lay-A-Way", "60 Day Lay-A-Way", "90 Day Lay-A-Way", "120 Day Lay-A-Way") Then MessageBox.Show("You can only apply the refund to an open sale.", "Sale not open") : Exit Sub
            If GetLeaseNoMailIndex(txtApplyToSaleNo.Text) <> GetLeaseNoMailIndex(VoidSaleNo) Then
                If MessageBox.Show("The names on the two sales do not match." & vbCrLf2 & "Are you sure you want apply the amount to this sale?", "Different Customers Selected", MessageBoxButtons.OKCancel) = DialogResult.Cancel Then
                    Exit Sub
                End If
            End If
        End If

        LogFile("voidsale", "void proc: " & VoidSaleNo, False)
        ' Go through the GrossMargin table and void each record.

        TotalRefunded = False

        Margin = New CGrossMargin
        Margin.Load(VoidSaleNo, "SaleNo")               ' This needs to be ordered by MarginLine.
        PaidAs = ""
        Do Until Margin.DataAccess.Record_EOF
            tST = Trim(Margin.Style)
            If IsIn(tST, "TAX1", "TAX2") Then
                Tax = Tax + Margin.SellPrice
            ElseIf tST = "PAYMENT" Then
                If TotalRefunded Then GoTo NoMoreRefundsThisSale
                PaidAs = QueryPaymentDescription(Margin.Quantity)
                Typ = "" & Margin.Quantity
                If SwipeCards() Then
                    If Margin.SellPrice < 0 Then
                        MessageBox.Show("The amount of " & FormatCurrency(Margin.SellPrice) & " indicates that this is a CC return." & vbCrLf & "No action will be performed for this already returned amount.", "Ignoring previously returned CC Purchase")
                    Else
                        If Not IsIn(RefundType, cdsRefundType.cdsrft_StoreCredit, cdsRefundType.cdsrft_ForfeitDeposit, cdsRefundType.cdsrft_ApplyToSale) Then
                            If SwipeCreditCards() And IsIn(Typ, cdsPayTypes.cdsPT_Visa, cdsPayTypes.cdsPT_MCard, cdsPayTypes.cdsPT_Disc, cdsPayTypes.cdsPT_amex) Then
                                Dim CCRRes As Boolean
                                If SaleHasBlindCredit(Margin.SaleNo, TotalAmt) Then
                                    ' if there are any credits, we must blind credit the resulting amount...
                                    CCRRes = ProcessCCReturn(TotalAmt, Appr, "#" & Margin.SaleNo, Margin.SellDte)
                                    TotalRefunded = True
                                Else
                                    CCRRes = ProcessCCReturn(Margin.SellPrice, Appr, Margin.TransID, Margin.SellDte)
                                End If

                                If Not CCRRes Then
                                    MessageBox.Show("Unable to return credit card purchase of " & FormatCurrency(Margin.SellPrice) & "." & vbCrLf & "You will have to manually return the amount from the Credit Card Manager function on the main menu to refund this amount.", "Credit Card Return Failed")
                                Else
                                    AddVoidLine(Margin.Index, VoidSaleNo, Margin.Name, StoresSld, dteVoidDate.Value, "CC Void Appr=" & Appr, "", Margin.Salesman, Margin.Phone)
                                    AddedVoidLine = True
                                End If
                            ElseIf SwipeDebitCards() And IsIn(Typ, "9") Then
                                If Not ProcessDebitReturn(Margin.SellPrice, Appr) Then
                                    MessageBox.Show("Unable to return debit card purchase." & vbCrLf & "You will have to manually return the amount from the Credit Card Manager function on the main menu to refund this amount.", "Debit Card Return Failed")
                                Else
                                    AddVoidLine(Margin.Index, VoidSaleNo, Margin.Name, StoresSld, dteVoidDate.Value, "Debit Void Appr=" & Appr, "", Margin.Salesman, Margin.Phone)
                                    AddedVoidLine = True
                                End If
                            ElseIf SwipeGiftCards() And IsIn(Typ, "12") Then
                                If Not ProcessGiftCardReturn(Margin.SellPrice, Appr) Then
                                    MessageBox.Show("Unable to return gift card purchase." & vbCrLf & "You will have to manually return the amount from the Credit Card Manager function on the main menu to refund this amount.", "Gift Card Return Failed")
                                Else
                                    AddVoidLine(Margin.Index, VoidSaleNo, Margin.Name, StoresSld, dteVoidDate.Value, "Gift Card Void Appr=" & Appr, "", Margin.Salesman, Margin.Phone)
                                    AddedVoidLine = True
                                End If
                            End If
                            'BFH20120815 not for type 4, apply to sale...  don't want to void the payment, we want to transfer it!
                            'BFH20121005 not for forfeit or store credit either!!
                        End If
                    End If
                End If
NoMoreRefundsThisSale:
            ElseIf Not IsIn(tST, "PAYMENT", "SUB", "--- Adj ---") Then
                Written = Written + Margin.SellPrice
            End If
            Margin.Void(dteVoidDate.Value)                         ' Void each item, returning to stock as necessary.
            BillOSale.SetStatus(Count, Margin.Status)  ' Update the display on bos2.
            Margin.DataAccess.Records_MoveNext()
            Margin.cDataAccess_GetRecordSet(Margin.DataAccess.RS)
            Count = Count + 1
        Loop
        '<CT>
        BillOSale.SetStyle(BillOSale.UGridIO1.LastRowUsed, BillOSale.UGridIO1.GetValue(BillOSale.UGridIO1.LastRowUsed, BillColumns.eStyle))
        '</CT>
        BillOSale.BalDue.Text = "0.00"                         ' Update the display on bos2.
        BillOSale.Refresh()

        LogFile("voidsale", "void proc loop done", False)

        ' And add a VOID line for the note..
        Dim VoidType As String
        If GetPrice(txtRefundSpecial.Text) <> 0 And (lblSpecialPaymentType.Tag = "41500" Or lblSpecialPaymentType.Tag = "21500" Or lblSpecialPaymentType.Text = "STORE CREDIT") Then VoidType = Trim(lblSpecialPaymentType.Text & " $" & CurrencyFormat(txtRefundSpecial.Text))
        If Not AddedVoidLine Then
            If SwipeCreditCards() Then
                If IsIn(Typ, "3", "4", "5", "6") Then
                    AddVoidLine(Margin.Index, VoidSaleNo, Margin.Name, StoresSld, dteVoidDate.Value, txtVoidNote.Text, VoidType, Margin.Salesman, Margin.Phone)
                Else
                    VoidType = Trim(VoidType & " " & DescribeVoidType(PaidAs, GetPrice(txtForfeit.Text))) & " " & vbShortDate
                    AddVoidLine(Margin.Index, VoidSaleNo, Margin.Name, StoresSld, dteVoidDate.Value, txtVoidNote.Text, VoidType, Margin.Salesman, Margin.Phone, False)
                End If
            Else
                VoidType = Trim(VoidType & " " & DescribeVoidType(PaidAs, GetPrice(txtForfeit.Text)))
                AddVoidLine(Margin.Index, VoidSaleNo, Margin.Name, StoresSld, dteVoidDate.Value, txtVoidNote.Text, VoidType, Margin.Salesman, Margin.Phone, False)
            End If
        End If

        LogFile("voidsale", "void load hold", False)

        Holding = New cHolding
        HoldingLoaded = Holding.Load(VoidSaleNo)

        ' Also add void payments for each non-zero txtRefundAmount.
        ' AddNewCashJournalRecord automatically filters out zero-sum entries.
        Dim El As Object
        Dim L As Label
        Dim A() As TextBox
        Dim I As Integer
        Dim Ptext As String, Ptag As String

        For Each C As Control In Me.fraPaymentSummary.Controls
            If Mid(C.Name, 1, 15) = "txtRefundAmount" Then
                ReDim Preserve A(I)
                A(I) = C
                I = I + 1
            End If
        Next

        I = 0
        'For Each El In txtRefundAmount.Text
        For Each El In A
            '    If GetPrice(El.Text) > 0 Then
            'L.Name = "lblPaymentType" & El.index
            'L = New Label
            If I = 0 Then
                Ptext = lblPaymentType.Text
                Ptag = lblPaymentType.Tag
            Else
                For Each C As Control In Me.fraPaymentSummary.Controls
                    If C.Name = "lblPaymentType" & I Then
                        Ptext = C.Text
                        Ptag = C.Tag
                        Exit For
                    End If
                Next
            End If

            'If IsIn(Holding.Status, "E", "S", "O") And PayTypeIsFinance(lblPaymentType(El.Index), False) Then  ' No CASH adjustment for these OPEN sale types...
            If IsIn(Holding.Status, "E", "S", "O") And PayTypeIsFinance(Ptext, False) Then  ' No CASH adjustment for these OPEN sale types...
                ' In the case of Open Credit/Store Finance, the
                'ElseIf Holding.Status = "F" And Val(lblPaymentType(El.Index).Tag) = cdsPayTypes.cdsPT_StoreFinance Then
            ElseIf Holding.Status = "F" And Val(Ptag) = cdsPayTypes.cdsPT_StoreFinance Then
                'ElseIf Holding.Status = "F" And PayTypeIsOutsideFinance(lblPaymentType(El.Index)) Then
            ElseIf Holding.Status = "F" And PayTypeIsOutsideFinance(Ptext) Then
                AddNewCashJournalRecord("11300", -GetPrice(El.Text), VoidSaleNo, "", dteVoidDate.Value)
                'ElseIf Holding.Status = "C" And Val(lblPaymentType(El.Index).Tag) = cdsPayTypes.cdsPT_StoreFinance Then
            ElseIf Holding.Status = "C" And Val(Ptag) = cdsPayTypes.cdsPT_StoreFinance Then
                AddNewCashJournalRecord("11300", -GetPrice(El.Text), VoidSaleNo, "", dteVoidDate.Value)
                'ElseIf Holding.Status = "C" And PayTypeIsOutsideFinance(lblPaymentType(El.Index)) Then
            ElseIf Holding.Status = "C" And PayTypeIsOutsideFinance(Ptext) Then
                'ElseIf Holding.Status = "D" And PayTypeIsOutsideFinance(lblPaymentType(El.Index)) Then
            ElseIf Holding.Status = "D" And PayTypeIsOutsideFinance(Ptext) Then
                '      AddNewCashJournalRecord "11300", -GetPrice(El.Text), VoidSaleNo, "", dteVoidDate
            Else
                'AddNewCashJournalRecord(lblPaymentType(El.Index).Tag, -GetPrice(El.Text), VoidSaleNo, Margin.Name, dteVoidDate.Value)
                AddNewCashJournalRecord(Ptag, -GetPrice(El.Text), VoidSaleNo, Margin.Name, dteVoidDate.Value)
            End If
            I = I + 1
        Next

        If GetPrice(txtRefundSpecial.Text) <> 0 Then
            If lblSpecialPaymentType.Tag <> "" Then
                AddNewCashJournalRecord(lblSpecialPaymentType.Tag, -GetPrice(txtRefundSpecial.Text), VoidSaleNo, Margin.Name, dteVoidDate.Value)
            Else
                If RefundType = cdsRefundType.cdsrft_ApplyToSale Then
                    ApplyPaymentToSale(Margin.SaleNo, txtApplyToSaleNo.Text, GetPrice(txtRefundSpecial.Text), dteVoidDate.Value, Margin.Name, Margin.Quantity, Margin.TransID)

                    AddNewCashJournalRecord("1", -GetPrice(txtRefundSpecial.Text), Margin.SaleNo, Margin.Name, dteVoidDate.Value)
                    AddNewCashJournalRecord("1", GetPrice(txtRefundSpecial.Text), txtApplyToSaleNo.Text, "", dteVoidDate.Value)

                    AddNewAuditRecord(txtApplyToSaleNo.Text, "PA ", DateFormat(dteVoidDate), 0, 0, 0, -GetPrice(txtRefundSpecial.Text), 0, 0, 0, 0, "")
                Else
                    ' BFH20051223
                    ' This should leave the record in the cash journal of where the money went!
                    AddNewCashJournalRecord("1", -GetPrice(txtRefundSpecial.Text), VoidSaleNo, Margin.Name, dteVoidDate.Value) ' First, pretend to refund the credit amount in cash.
                    AddNewCashJournalRecord("1", GetPrice(txtRefundSpecial.Text), VoidSaleNo & "SC", Margin.Name, dteVoidDate.Value)  ' Then, put that money back into the system as forfeit.

                    RemoveSCDepositFromHolding(VoidSaleNo, GetPrice(txtRefundSpecial.Text))
                    CreateCreditMemo(VoidSaleNo & "SC", GetPrice(txtRefundSpecial.Text), dteVoidDate.Value, Margin)
                    AddSCLine(Margin, VoidSaleNo, dteVoidDate.Value, lblRefundTotal.Text)
                End If
            End If
        End If

        'BFH20071219 forfeits now handled separately
        If GetPrice(txtForfeit.Text) > 0 Then
            ' First, pretend to refund the forfeit amount in cash.
            AddNewCashJournalRecord("1", -GetPrice(txtForfeit.Text), VoidSaleNo, Margin.Name, dteVoidDate.Value)
            ' Then, put that money back into the system as forfeit.
            AddNewCashJournalRecord("41500", GetPrice(txtForfeit.Text), VoidSaleNo, Margin.Name, dteVoidDate.Value)
        End If
        ' Need: Written, Tax

        LogFile("voidsale", "void save audit", False)


        If HoldingLoaded Then   ' Save a new audit record too..
            ' bfh20050819 - changed from "132" to "11200"
            Dim Tp As Decimal
            Tp = BillOSale.SaleTotal("paid") -
          BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance)) +
          BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance2)) +
          BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance3)) +
          BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance4)) +
          BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance5)) - BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_StoreFinance))
            Select Case Holding.Status
                Case "B"
                    AddNewAuditRecord(VoidSaleNo, "V" & Holding.Status & " " & Margin.Name, dteVoidDate.Value, -Written, -Tax, 0, 0, 0, -Math.Abs(Written), -Tax, SaleTaxCode, Margin.Salesman, -GetPrice(Holding.NonTaxable))
                    AddNewCashJournalRecord("11200", Holding.Deposit - Holding.Sale, VoidSaleNo, "", dteVoidDate.Value)       ' Add a cash record for the backorder.
                Case "S" ' Open Store Finance
                    AddNewAuditRecord(VoidSaleNo, "V" & Holding.Status & " " & Margin.Name, dteVoidDate.Value, -Written, -Tax, -BillOSale.SaleTotal("gross"), Tp, -BillOSale.SaleTotal("gross"), 0, 0, SaleTaxCode, Margin.Salesman, -GetPrice(Holding.NonTaxable))
                Case "F"
                    AddNewAuditRecord(VoidSaleNo, "V" & Holding.Status & " " & Margin.Name, dteVoidDate.Value, -Written, -Tax, 0, 0, 0, -Math.Abs(Written), -Tax, SaleTaxCode, Margin.Salesman, -GetPrice(Holding.NonTaxable))
                    AddNewCashJournalRecord("11300", Holding.Deposit - Holding.Sale, VoidSaleNo, "", dteVoidDate.Value)       ' Add a cash record for the backorder.
                Case "E"
                    AddNewAuditRecord(VoidSaleNo, "V" & Holding.Status & " " & Margin.Name, dteVoidDate.Value, -Written, -Tax, -BillOSale.SaleTotal("gross"), Tp, -BillOSale.SaleTotal("gross"), 0, 0, SaleTaxCode, Margin.Salesman, -GetPrice(Holding.NonTaxable))
                Case "C"
                    AddNewAuditRecord(VoidSaleNo, "V" & Holding.Status & " " & Margin.Name, dteVoidDate.Value, -Written, -Tax, 0, 0, 0, -Math.Abs(Written), -Tax, SaleTaxCode, Margin.Salesman, -GetPrice(Holding.NonTaxable))
                    AddNewCashJournalRecord("11300", Holding.Deposit - Holding.Sale, VoidSaleNo, "", dteVoidDate.Value)        ' Add a cash record for the backorder.
                Case "D"
                    AddNewAuditRecord(VoidSaleNo, "VD " & Margin.Name, dteVoidDate.Value, -Written, -Tax, 0, 0, 0, -Written, -Tax, SaleTaxCode, Margin.Salesman, -GetPrice(Holding.NonTaxable))
                    AddNewCashJournalRecord("11300", -BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance)), VoidSaleNo, "", dteVoidDate.Value)      ' Add a cash record for the backorder.
                    AddNewCashJournalRecord("11300", -BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance2)), VoidSaleNo, "", dteVoidDate.Value)        ' Add a cash record for the backorder.
                    AddNewCashJournalRecord("11300", -BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance3)), VoidSaleNo, "", dteVoidDate.Value)        ' Add a cash record for the backorder.
                    AddNewCashJournalRecord("11300", -BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance4)), VoidSaleNo, "", dteVoidDate.Value)        ' Add a cash record for the backorder.
                    AddNewCashJournalRecord("11300", -BillOSale.SaleTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance5)), VoidSaleNo, "", dteVoidDate.Value)        ' Add a cash record for the backorder.
                Case "O", "L", "1", "2", "3", "4"
                    AddNewAuditRecord(VoidSaleNo, "VO " & Margin.Name, dteVoidDate.Value, -(GetPrice(Holding.Sale) - GetPrice(Tax)), -GetPrice(Tax), -GetPrice(Holding.Sale), GetPrice(Holding.Deposit), -GetPrice(Holding.Sale), 0, 0, SaleTaxCode, Margin.Salesman, -GetPrice(Holding.NonTaxable))
                Case Else
                    ' Already void, maybe?
            End Select
        Else
            ' Can't load the sale, but the sale loaded this!!
        End If

        LogFile("voidsale", "void done", False)
        ' Set the success variable.
        OrderVoided = True
        DisposeDA(Margin, Holding)
        'Unload Me
        Me.Close()
    End Sub

    Private Sub AddVoidLine(ByVal Index As Integer, ByVal SaleNo As String, ByVal LastName As String, ByVal Store As Integer, ByVal VoidDate As Date, ByVal VoidNote As String, ByVal VoidType As String, ByVal Salesman As String, ByVal Tele As String, Optional ByVal ShowAuthorizationLine As Boolean = True)
        Dim Margin As New CGrossMargin
        Margin.DataAccess.DataBase = GetDatabaseAtLocation(Store)
        Margin.Name = LastName
        Margin.Index = Index
        Margin.SaleNo = SaleNo
        Margin.Location = Store
        Margin.SellDte = VoidDate
        Margin.Style = "NOTES"
        Margin.Vendor = ""
        Margin.Status = "VOID"
        Margin.Quantity = 0
        If VoidNote = "" Then
            Margin.Desc = Trim(ArrangeString("Voided:" & " " & VoidType & " " & DateFormat(VoidDate), 46) & IIf(ShowAuthorizationLine, ArrangeString("I Authorize The Above Transaction:", 46) & "X ___________________________________", ""))
        Else ' BFH20150509 - The note made the void date get cut off.  This replaces the "I Authorize" line with the note.
            Margin.Desc = Trim(ArrangeString("Voided:" & " " & VoidType & " " & DateFormat(VoidDate), 46) & IIf(ShowAuthorizationLine, ArrangeString(VoidNote, 46) & "X ___________________________________", ""))
        End If
        Margin.SellPrice = 0
        Margin.Detail = 0
        Margin.Commission = ""
        Margin.Salesman = Salesman
        Margin.Phone = Tele
        Margin.Save()

        BillOSale.AddMarginRow(Margin)  '' Very hackish, MJK20070724
        DisposeDA(Margin)
    End Sub

    Public Sub RemoveSCDepositFromHolding(ByVal SaleNo As String, ByVal CreditAmt As Decimal)
        Dim Ch As cHolding
        Ch = New cHolding
        If Not Ch.Load(SaleNo, "LeaseNo") Then Exit Sub
        Ch.Deposit = Ch.Deposit - CreditAmt
        Ch.Save()
        DisposeDA(Ch)
    End Sub

    Public Sub CreateCreditMemo(ByVal SaleNo As String, ByVal Credit As Decimal, ByVal SaleDate As Date, ByRef Margin As CGrossMargin)
        ' Margin is a record from the parent sale we'll be copying...
        ' This is an ugly way to deal with the problem......

        Dim Holding As cHolding
        Holding = New cHolding
        If Not Holding.Load(SaleNo, "LeaseNo") Then
            Holding.LeaseNo = SaleNo                          ' Use the provided lease number.
            Holding.Index = Margin.Index                      ' Use the same customer number.
            Holding.Deposit = 0                               ' Amount paid.
            Holding.Sale = 0                                  ' Total amount of sale, with tax
            Holding.NonTaxable = 0                            ' Amount that's not taxable..
            Holding.LastPay = Today                            ' Paid today.
            Holding.Salesman = Margin.Salesman                ' Retain the salesman..
            Holding.Comm = "N"                                ' Commission isn't paid.
        End If
        Holding.Status = "O"                                ' All Store Credits are open sales.
        Holding.Deposit = Holding.Deposit + Credit          ' Amount paid.
        Holding.Save()

        ' The only entry on the credit memo sale is a payment.
        AddNewMarginRecord(Holding.LeaseNo, "PAYMENT", "STORE CREDIT" & Space(5) & DateFormat(SaleDate), 1, Credit,
  "", "0", "000", 0, 0, 0, "", "", "", Holding.Salesman,
  StoresSld, DateFormat(SaleDate), DateFormat(SaleDate), StoresSld, Margin.Name, DateFormat(SaleDate), Margin.Phone,
  Margin.Index, "100", 0, "", "", "", "", "")

        ' Account for the payment in the cash & sales journals.
        '  AddNewAuditRecord SaleNo, "NS " & Margin.Name, DateFormat(SaleDate), 0, 0, 0, -GetPrice(txtRefundSpecial.Text), 0, 0, 0, SaleTaxCode, Margin.Salesman
        AddNewAuditRecord(SaleNo, "NS " & Margin.Name, DateFormat(SaleDate), 0, 0, 0, -Credit, 0, 0, 0, SaleTaxCode, Margin.Salesman, 0)

        ' Print the credit memo...
        PrintSale(Holding.LeaseNo, , 1)

        DisposeDA(Holding)
    End Sub

    Public Sub AddSCLine(ByRef vM As CGrossMargin, ByVal SaleNo As String, ByVal VoidDate As Date, ByVal Amt As Object, Optional ByVal AsDelivered As Decimal = False)
        Dim Margin As New CGrossMargin
        Margin.DataAccess.DataBase = GetDatabaseAtLocation()
        Margin.Name = vM.Name
        Margin.Index = vM.Index
        Margin.SaleNo = SaleNo
        Margin.Location = StoresSld
        Margin.SellDte = VoidDate
        '  Margin.Style = "NOTES"
        Margin.Style = "PAYMENT" ' bfh20060113 CHANGED FROM (+) NOTE TO (-) PAYMENT
        Margin.Vendor = ""
        Margin.Status = IIf(AsDelivered, "DEL", "VOID")
        Margin.Quantity = 0
        Margin.Desc = "Store Credit for " & FormatCurrency(GetPrice(Amt)) & " (" & SaleNo & "SC) " & vbShortDate
        '  Margin.SellPrice = GetPrice(Amt) '0 'IIf(AsDelivered, GetPrice(Amt), 0)
        Margin.SellPrice = -GetPrice(Amt) '0 'IIf(AsDelivered, GetPrice(Amt), 0)
        Margin.Detail = 0
        Margin.Commission = ""
        Margin.Salesman = vM.Salesman
        Margin.Phone = vM.Phone
        Margin.Save()
        DisposeDA(Margin)
    End Sub

    Public Sub ApplyPaymentToSale(ByVal OldSaleNo As String, ByVal SaleNo As String, ByVal Credit As Decimal, ByVal SaleDate As Date, ByVal Name As String, Optional ByVal PayType As String = "", Optional ByVal TransID As String = "")
        Dim H As cHolding

        H = New cHolding
        H.Load(SaleNo, "LeaseNo")

        AddNewMarginRecord(SaleNo, "SUB", "               Sub Total = ", 0, H.Sale - H.Deposit)
        AddNewMarginRecord(SaleNo, "PAYMENT", UCase(PayTypeName(PayType)) & ": Transfer From Sale No #" & OldSaleNo, PayType, Credit, , , , , , , , , , , , , , , , , , , , , , , , , TransID)

        H.Deposit = H.Deposit + Credit
        H.LastPay = SaleDate
        H.Save()
        DisposeDA(H)
    End Sub

    Private Sub AddPaymentLine(ByVal PayType As Integer, ByVal PayAmount As Decimal)
        Dim pType As String
        Dim PRow As Integer

        If PayType = 7 Or PayAmount = 0 Then Exit Sub

        pType = UCase(TranslateAccountCode(CStr(PayType), "Unknown"))
        If PaymentCount = 0 Then
            ' Load into the existing row.
            PRow = 0
        Else
            ' If this payment type is already in use, add to its existing row.
            Dim El As Object

            PRow = -1
            For Each El In lblPaymentType.Text
                If El.text = pType Then
                    PRow = El.Index
                    Exit For
                End If
            Next

            If PRow = -1 Then
                'PRow = lblPaymentType.UBound + 1
                Dim C As Control, Cnt As Integer
                For Each C In Me.Controls
                    If Mid(C.Text, 1, 14) = "lblPaymentType" Then
                        Cnt = Cnt + 1
                    End If
                Next

                PRow = Cnt + 1
                ' Create controls for a new row.
                'Load lblPaymentType(PRow)
                'Load lblAmountPaid(PRow)
                'Load txtRefundAmount(PRow)

                ' Move the controls to their new homes.
                Dim RowTop As Integer
                'RowTop = lblPaymentType(0).Top + 360 * PRow
                'lblPaymentType(PRow).Move lblPaymentType(0).Left, RowTop, lblPaymentType(0).Width, lblPaymentType(0).Height
                'lblAmountPaid(PRow).Move lblAmountPaid(0).Left, RowTop, lblAmountPaid(0).Width, lblAmountPaid(0).Height
                'txtRefundAmount(PRow).Move txtRefundAmount(0).Left, RowTop, txtRefundAmount(0).Width, txtRefundAmount(0).Height
                RowTop = lblPaymentType.Top + 36 * PRow
                Dim L As New Label
                L.Name = "lblPaymentType" & PRow
                L.Tag = PayType
                L.Text = pType
                L.Location = New Point(lblPaymentType.Left, RowTop)
                L.Size = New Size(lblPaymentType.Width, lblPaymentType.Height)
                Me.Controls.Add(L)

                L = New Label
                L.Name = "lblAmountPaid" & PRow
                L.Text = "0.00"
                L.Text = CurrencyFormat(GetPrice(L.Text) + PayAmount)
                L.TextAlign = ContentAlignment.TopRight
                L.Location = New Point(lblAmountPaid.Left, RowTop)
                L.Size = New Size(lblAmountPaid.Width, lblAmountPaid.Height)
                Me.Controls.Add(L)

                Dim T As TextBox
                T = New TextBox
                T.Name = "txtRefundAmount" & PRow
                T.Text = "0.00"
                T.Text = CurrencyFormat(GetPrice(T.Text) + PayAmount)
                T.Tag = T.Text
                T.TextAlign = HorizontalAlignment.Right
                T.Location = New Point(txtRefundAmount.Left, RowTop)
                T.Size = New Size(txtRefundAmount.Width, txtRefundAmount.Height)
                Me.Controls.Add(T)


                ' Make the new controls visible.
                'lblPaymentType(PRow).Visible = True
                'lblAmountPaid(PRow).Visible = True
                'txtRefundAmount(PRow).Visible = True

                ' Set extended properties in the new controls..
                'lblAmountPaid(PRow).Alignment = 1
                'txtRefundAmount(PRow).Alignment = 1
                'lblAmountPaid(PRow).Caption = "0.00"
                'txtRefundAmount(PRow).Text = "0.00"
                'txtRefundAmount(PRow).TabIndex = txtRefundAmount(PRow - 1).TabIndex + 1

                '      Debug.Print "Row " & PaymentCount & " moved to " & RowTop & "."

                ' Move the special controls and totals.
                'RowTop = lblPaymentType(0).Top + 360 * (PRow + 1)
                RowTop = lblPaymentType.Top + 360 * (PRow + 1)
                lblSpecialPaymentType.Top = RowTop
                txtRefundSpecial.Top = RowTop
                '      Debug.Print "Special Row moved to " & RowTop & "."

                'RowTop = lblPaymentType(0).Top + 360 * (PRow + 2)
                RowTop = lblPaymentType.Top + 360 * (PRow + 2)
                lblForfeit.Top = RowTop
                txtForfeit.Top = RowTop

                'RowTop = lblPaymentType(0).Top + 360 * (PRow + 3)
                RowTop = lblPaymentType.Top + 360 * (PRow + 3)
                lblTotalPaidLabel.Top = RowTop
                lblTotalPaid.Top = RowTop
                lblRefundTotal.Top = RowTop
                '      Debug.Print "Totals Row moved to " & RowTop & "."

                ' Adjust the frame and form heights.
                fraPaymentSummary.Height = lblRefundTotal.Top + lblRefundTotal.Height + 60
                '      Debug.Print "Frame height changed to " & fraPaymentSummary.Height & "."
                'Height = fraPaymentSummary.Top + fraPaymentSummary.Height + 60 + (Height - ScaleHeight)
                Height = fraPaymentSummary.Top + fraPaymentSummary.Height + 60 + (Height - Me.ClientSize.Height)
                '      Debug.Print "ScaleHeight changed to " & ScaleHeight & "."
            End If
        End If

        ' Set the values in the visible controls.
        'lblPaymentType(PRow).Caption = pType
        'lblPaymentType(PRow).Tag = PayType
        'lblAmountPaid(PRow).Caption = CurrencyFormat(GetPrice(lblAmountPaid(PRow).Caption) + PayAmount)
        'txtRefundAmount(PRow).Text = CurrencyFormat(GetPrice(txtRefundAmount(PRow).Text) + PayAmount)
        'txtRefundAmount(PRow).Tag = txtRefundAmount(PRow).Text

        ' Adjust the sale totals.
        'lblTotalPaid.Caption = "$" & CurrencyFormat(GetPrice(lblTotalPaid.Caption) + PayAmount)
        lblTotalPaid.Text = "$" & CurrencyFormat(GetPrice(lblTotalPaid.Text) + PayAmount)
        RecalculateRefundTotal

        PaymentCount = PaymentCount + 1
    End Sub

    Private Sub RecalculateRefundTotal()
        Dim El As Object, X As Decimal, Total As Decimal
        Dim A() As TextBox
        Dim I As Integer

        For Each C As Control In Me.fraPaymentSummary.Controls
            If Mid(C.Name, 1, 15) = "txtRefundAmount" Then
                ReDim Preserve A(I)
                A(I) = C
                I = I + 1
            End If
        Next

        Total = GetPrice(txtRefundSpecial.Text)
        'For Each El In txtRefundAmount.Text
        For Each El In A
            Total = Total + GetPrice(El.Text)
        Next

        X = GetPrice(lblTotalPaid.Text) - Total
        If X < 0 Then X = 0
        txtForfeit.Text = FormatCurrency(X)
        lblForfeit.Visible = X > 0
        txtForfeit.Visible = X > 0
        txtForfeit.ForeColor = Color.Blue
        lblForfeit.ForeColor = Color.Blue
        lblRefundTotal.Text = FormatCurrency(Total + X)
        lblRefundTotal.ForeColor = IIf(GetPrice(lblRefundTotal.Text) = GetPrice(lblTotalPaid.Text), Color.Black, Color.Red)
    End Sub

    Private Sub dteVoidDate_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles dteVoidDate.Validating
        If Not CheckAccess("Change Sale Date", True, True) Then e.Cancel = True
    End Sub

    Private Sub OrdVoid_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If OrdVoidFormLoad = True Then Exit Sub
        SetButtonImage(cmdCancel, 3)
        SetButtonImage(cmdOk, 2)
        'optRefundType_Click 0
        optRefundTypeClick(optRefundType0, New EventArgs)
    End Sub

    Private Sub optRefundTypeClick(sender As Object, e As EventArgs) Handles optRefundType0.CheckedChanged, optRefundType1.CheckedChanged, optRefundType2.CheckedChanged, optRefundType3.CheckedChanged, optRefundType4.CheckedChanged
        Dim optName As String

        optName = CType(sender, RadioButton).Name
        Dim El As Object
        Dim A() As TextBox
        Dim I As Integer

        For Each C As Control In Me.fraPaymentSummary.Controls
            If Mid(C.Name, 1, 15) = "txtRefundAmount" Then
                ReDim Preserve A(I)
                A(I) = C
                I = I + 1
            End If
        Next

        'For Each El In txtRefundAmount : El.Text = "0.00" : Next
        On Error Resume Next
        For Each El In A : El.Text = "0.00" : Next
        txtRefundSpecial.Text = "0.00"
        lblSpecialPaymentType.Text = ""
        lblSpecialPaymentType.Tag = ""
        txtRefundSpecial.Visible = False
        lblSpecialPaymentType.Visible = False
        txtApplyToSaleNo.Visible = False
        txtApplyToSaleNo.Text = ""

        'Select Case Index
        Select Case optName
            'Case 0                                                                ' Return as paid
            Case "optRefundType0"
                For Each El In A
                    'El.Text = lblAmountPaid(El.Index)
                    If optName = "optRefundType0" Then
                        El.Text = lblAmountPaid.Text
                    ElseIf optName = "optRefundType1" Then
                        Dim L As Label
                        L.Name = "lblAmountPaid1"
                        El.Text = L.Text
                    ElseIf optName = "optRefundType2" Then
                        Dim L As Label
                        L.Name = "lblAmountPaid2"
                        El.Text = L.Text
                    ElseIf optName = "optRefundType3" Then
                        Dim L As Label
                        L.Name = "lblAmountPaid3"
                        El.Text = L.Text
                    ElseIf optName = "optRefundType4" Then
                        Dim L As Label
                        L.Name = "lblAmountPaid4"
                        El.Text = L.Text
                    End If
                Next
                RecalculateRefundTotal()
            'Case 1                                                                ' Company Check
            Case "optRefundType1"
                txtRefundSpecial.Text = CurrencyFormat(lblTotalPaid.Text)
                lblSpecialPaymentType.Text = "COMPANY CHECK"
                lblSpecialPaymentType.Tag = "21500"    ' BFH20050412
                txtRefundSpecial.Visible = True
                lblSpecialPaymentType.Visible = True
                RecalculateRefundTotal()
            'Case 2                                                                ' Store Credit
            Case "optRefundType2"
                txtRefundSpecial.Text = CurrencyFormat(lblTotalPaid.Text)
                lblSpecialPaymentType.Text = "STORE CREDIT"
                lblSpecialPaymentType.Tag = ""    ' No entry for store credits.
                txtRefundSpecial.Visible = True
                lblSpecialPaymentType.Visible = True
                RecalculateRefundTotal()
            'Case 3                                                                ' Forfeit
            Case "optRefundType3"
                ' Validate this choice by password.
                If Not CheckAccess("Forfeit Deposits", True) Then
                    'optRefundType(0).Value = True
                    optRefundType0.Checked = True
                    'optRefundType(0).SetFocus
                    optRefundType0.Select()
                    Exit Sub
                End If
                RecalculateRefundTotal()
            'Case 4                                                                ' apply to sale
            Case "optRefundType4"
                txtApplyToSaleNo.Visible = True
                On Error Resume Next
                txtApplyToSaleNo.Select()
                txtRefundSpecial.Text = CurrencyFormat(lblTotalPaid.Text)
                lblSpecialPaymentType.Text = "APPLY TO SALE"
                lblSpecialPaymentType.Tag = ""
                txtRefundSpecial.Visible = True
                lblSpecialPaymentType.Visible = True
                RecalculateRefundTotal()
            Case Else
        End Select
    End Sub

    Private Sub OrdVoid_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        PaymentCount = 0
        VoidSaleNo = ""
    End Sub

    Private Sub txtRefundAmount_DoubleClick(sender As Object, e As EventArgs) Handles txtRefundAmount.DoubleClick
        Dim txtName As String
        Dim T As TextBox

        txtName = CType(sender, TextBox).Name
        T = CType(sender, TextBox)

        'If GetPrice(txtRefundAmount(Index)) = 0 Then
        '    txtRefundAmount(Index) = CurrencyFormat(GetPrice(lblAmountPaid(Index)))
        'Else
        '    txtRefundAmount(Index) = "0.00"
        'End If
        'txtRefundAmount_LostFocus Index

        If GetPrice(T.Text) = 0 Then
            Dim N As Integer
            Dim L As Label

            N = Mid(txtName, 16)
            L.Name = "lblAmountPaid" & N
            T.Text = CurrencyFormat(GetPrice(L.Text))
        Else
            T.Text = "0.00"
        End If
        txtRefundAmount_Leave(T, New EventArgs)
    End Sub

    Private Sub txtRefundAmount_Leave(sender As Object, e As EventArgs) Handles txtRefundAmount.Leave
        Dim T As TextBox

        T = CType(sender, TextBox)

        On Error Resume Next
        'If Not ValidPrice(txtRefundAmount(Index).Text) Then
        '    MsgBox "Invalid refund price.", vbCritical
        '    txtRefundAmount(Index).SetFocus
        '    Exit Sub
        'End If
        '' Should we allow this value to be greater than the original amount?
        'txtRefundAmount(Index).Text = CurrencyFormat(txtRefundAmount(Index).Text)
        'RecalculateRefundTotal()

        If Not ValidPrice(T.Text) Then
            MessageBox.Show("Invalid refund price.", "WinCDS")
            T.Select()
            Exit Sub
        End If

        T.Text = CurrencyFormat(T.Text)
        RecalculateRefundTotal()
    End Sub

    Private Function ValidPrice(ByVal Price As String) As Boolean
        If Trim(Price) = "" Then Price = "0"
        If Not IsNumeric(Price) Then Exit Function
        If GetPrice(Price) < 0 Then Exit Function
        ValidPrice = True
    End Function

    Private Sub txtRefundSpecial_Enter(sender As Object, e As EventArgs) Handles txtRefundSpecial.Enter
        SelectContents(txtRefundSpecial)
    End Sub

    Private Sub txtRefundAmount_Enter(sender As Object, e As EventArgs) Handles txtRefundAmount.Enter
        'SelectContents txtRefundAmount(Index)
        SelectContents(sender)
    End Sub

    Private Sub txtRefundSpecial_DoubleClick(sender As Object, e As EventArgs) Handles txtRefundSpecial.DoubleClick
        If GetPrice(txtRefundSpecial.Text) = 0 Then
            txtRefundSpecial.Text = CurrencyFormat(GetPrice(txtForfeit.Text))
        Else
            txtRefundSpecial.Text = "0.00"
        End If
        'txtRefundSpecial_LostFocus
        txtRefundSpecial_Leave(sender, New EventArgs)
    End Sub

    Private Sub txtRefundSpecial_Leave(sender As Object, e As EventArgs) Handles txtRefundSpecial.Leave
        If Not ValidPrice(txtRefundSpecial.Text) Then
            MessageBox.Show("Invalid refund price.", "WinCDS")
            txtRefundSpecial.Select()
            Exit Sub
        End If
        txtRefundSpecial.Text = CurrencyFormat(txtRefundSpecial.Text)
        RecalculateRefundTotal()
    End Sub
End Class