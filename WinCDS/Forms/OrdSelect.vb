Public Class OrdSelect
    Public X As Object
    Public Sale As Decimal
    Public BalDue As Decimal
    Dim Amt As Decimal
    Public NonTaxable As Decimal
    Dim Delivery As Double
    Dim Labor As Decimal
    Public SalesTax1 As Decimal
    Public SalesTax2 As Decimal
    Public Store As Decimal
    Dim TaxRate As Object
    Dim Rate As Object
    Dim Num As Object
    Public TaxApplied As String
    Public ArStatus As String

    Private LastListClick As Integer
    Private LastListItemData As Integer
    Private LastListItemText As String

    Dim WithEvents mInvCkStyle As InvCkStyle
    Public RowChangeOK As Boolean                 ' Allow BillOSale to change its adding row
    Private mInvCkStyleShown As Boolean           ' Is the style selector shown?

    Private ReadOnly Property UseOutsideCreditOnNewSale() As Boolean
        Get
            UseOutsideCreditOnNewSale = False
            UseOutsideCreditOnNewSale = UseOutsideCreditOnNewSale Or IsDevelopment()
            UseOutsideCreditOnNewSale = UseOutsideCreditOnNewSale Or IsPitUSA
            UseOutsideCreditOnNewSale = UseOutsideCreditOnNewSale Or IsSidesFurniture
        End Get
    End Property

    Private Sub Unload_InvCkStyle()
        '!-11/16/98:AA:    Unload mInvCkStyle
        mInvCkStyleShown = False
        If Not mInvCkStyle Is Nothing Then
            'Unload mInvCkStyle
            mInvCkStyle.Close()
        End If
        mInvCkStyle = Nothing
        '  Debug.Print "Unloading mInvCkStyle"
    End Sub

    Private Sub chkPayAll_Click(sender As Object, e As EventArgs) Handles chkPayAll.Click
        ' If no payment style is selected, prompt for one.
        If chkPayAll.Checked = False Then Exit Sub
        If LastListClick = 0 And LastListItemText = "" Then
            'optPayment.Value = True
            optPayment.Checked = True
            LoadPaymentOptionsIntoListbox(lstOptions)
            'ShowListbox
        End If
        On Error Resume Next
        'lstOptions.SetFocus
        lstOptions.Select()

    End Sub

    Private Sub LoadPaymentOptionsIntoListbox(ByRef lst As ListBox)
        lstOptions.Items.Clear()
        If DoPayType(cdsPayTypes.cdsPT_Cash) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_Cash), , cdsPayTypes.cdsPT_Cash)
        If DoPayType(cdsPayTypes.cdsPT_Check) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_Check), , cdsPayTypes.cdsPT_Check)
        If DoPayType(cdsPayTypes.cdsPT_Visa) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_Visa), , cdsPayTypes.cdsPT_Visa)
        If DoPayType(cdsPayTypes.cdsPT_MCard) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_MCard), , cdsPayTypes.cdsPT_MCard)
        If DoPayType(cdsPayTypes.cdsPT_Disc) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_Disc), , cdsPayTypes.cdsPT_Disc)
        If DoPayType(cdsPayTypes.cdsPT_amex) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_amex), , cdsPayTypes.cdsPT_amex)
        If DoPayType(cdsPayTypes.cdsPT_DebitCard) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_DebitCard), , cdsPayTypes.cdsPT_DebitCard)
        If DoPayType(cdsPayTypes.cdsPT_StoreCreditCard) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_StoreCreditCard), , cdsPayTypes.cdsPT_StoreCreditCard)
        If DoPayType(cdsPayTypes.cdsPT_ECheck) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_ECheck), , cdsPayTypes.cdsPT_ECheck)

        If UseOutsideCreditOnNewSale Then
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_OutsideFinance), , cdsPayTypes.cdsPT_OutsideFinance)
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance2) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_OutsideFinance2), , cdsPayTypes.cdsPT_OutsideFinance2)
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance3) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_OutsideFinance3), , cdsPayTypes.cdsPT_OutsideFinance)
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance4) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_OutsideFinance4), , cdsPayTypes.cdsPT_OutsideFinance)
            If DoPayType(cdsPayTypes.cdsPT_OutsideFinance5) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_OutsideFinance5), , cdsPayTypes.cdsPT_OutsideFinance)
        End If

        If Installment Then
            If DoPayType(cdsPayTypes.cdsPT_StoreFinance) Then AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_StoreFinance), , cdsPayTypes.cdsPT_StoreFinance)
        End If

        If IsFormLoaded("BilloSale") Then
            If Val(BillOSale.BalDue.Text) < 0 Then
                AddListItem(lst, PayListItem(cdsPayTypes.cdsPT_CompanyCheck), , 21500) ' exchange for check refunds (was 21400)
            End If
        End If

        chkPayAll.Enabled = True
        On Error Resume Next
        'lst.Selected(0) = True
        lst.SetSelected(0, True)
    End Sub

    Private Sub AddListItem(ByRef lst As ListBox, ByVal Item As String, Optional ByVal ListIndex As Integer = -1, Optional ByVal itemData As Integer = 0)
        If ListIndex < 0 Then
            'lst.AddItem Item
            'lst.Items.Add(Item)
            lst.Items.Add(New ItemDataClass(Item, itemData))
        Else
            'lst.AddItem Item, ListIndex
            'lst.Items.Insert(ListIndex, Item)
            lst.Items.Insert(ListIndex, New ItemDataClass(Item, itemData))
        End If
        'lst.itemData(lst.NewIndex) = itemData
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdOk_Click(sender As Object, e As EventArgs) Handles cmdOk.Click
        Dim X As Integer, IsStoreFinance As Boolean, Xx As Double
        Dim OkToProcess As Boolean, DelRate As Double
        OkToProcess = False

        X = BillOSale.NewStyleLine   ' Was BillOSale.X
        BillOSale.X = X

        If optEnterStyle.Checked = True Then                              ' ******************************** style
            BillOSale.LocEnabled = True
            BillOSale.StatusEnabled = True
            BillOSale.QuanEnabled = False
            BillOSale.DescEnabled = False
            Hide()
            mInvCkStyleShown = True
            If mInvCkStyle Is Nothing Then
                mInvCkStyle = New InvCkStyle
                mInvCkStyle.Tag = "x"
            End If
            mInvCkStyle.Show()             ' Get the new item's style number, or SS/SO/FND.

        ElseIf optStain.Checked = True Then                               ' ********************************stain prot
            BillOSale.RowClear(X)
            BillOSale.SetStyle(X, "STAIN")
            If IsBFMyer Then
                BillOSale.SetDesc(X, "SAFEWARE PROTECTION PLAN")
            Else
                BillOSale.SetDesc(X, "STAIN PROTECTION")
            End If
            BillOSale.StyleEnabled = False
            BillOSale.MfgEnabled = False
            BillOSale.LocEnabled = False
            BillOSale.StatusEnabled = False
            BillOSale.QuanEnabled = False
            BillOSale.DescEnabled = True
            BillOSale.PriceFocus()
            BillOSale.StyleAddEnd()

        ElseIf optDelivery.Checked = True Then                            ' ******************************** delivery
            BillOSale.RowClear(X)
            BillOSale.SetStyle(X, "DEL")
            DelRate = Val(StoreSettings.DelPercent) '   frmSetup BillOSale.QueryDelPercent
            If DelRate > 0 Then
                BillOSale.SetDesc(X, "DELIVERY CHARGE (" & StoreSettings.DelPercent & ")")
            Else
                BillOSale.SetDesc(X, "DELIVERY CHARGE")
            End If
            BillOSale.StyleEnabled = False
            BillOSale.MfgEnabled = False
            BillOSale.LocEnabled = False
            BillOSale.StatusEnabled = False
            BillOSale.QuanEnabled = False
            BillOSale.DescEnabled = True
            If DelRate > 0 Then
                BillOSale.SetPrice(X, GetPrice(BillOSale.BalDue.Text) * DelRate / 100)
            End If
            BillOSale.PriceFocus()
            BillOSale.StyleAddEnd()

        ElseIf optLabor.Checked = True Then                               ' ******************************** labor
            BillOSale.RowClear(X)
            BillOSale.SetStyle(X, "LAB")
            BillOSale.SetDesc(X, "LABOR")
            BillOSale.StyleEnabled = False
            BillOSale.MfgEnabled = False
            BillOSale.LocEnabled = False
            BillOSale.StatusEnabled = False
            BillOSale.QuanEnabled = False
            BillOSale.DescEnabled = True
            BillOSale.PriceFocus()
            BillOSale.StyleAddEnd()

        ElseIf optTax1.Checked = True Then                                '******************************** tax 1
            OkToProcess = True
            BillOSale.RowClear(X)
            BillOSale.RowClear(X + 1)
            BillOSale.SetStyle(X, "SUB")
            BillOSale.SetDesc(X, "               Sub Total =")
            '.SetPrice X, BillOSale.Written - BillOSale.Deposit
            BillOSale.SetPrice(X, CurrencyFormat(BillOSale.SubTotal(X - 1))) ' format(BillOSale.Written - BillOSale.Deposit + SalesTax1, "###,###.00")

            BillOSale.NewStyleLine = X + 1
            BillOSale.X = X + 1
            BillOSale.SetStyle(X + 1, "TAX1")
            BillOSale.SetQuan(X + 1, 1)
            BillOSale.SetDesc(X + 1, "SALES TAX:   (" & (StoreSettings.SalesTax * 100) & "%" & IIf(IsDoddsLtd, " HST", "") & ")")

            BillOSale.StyleEnabled = False
            BillOSale.MfgEnabled = False
            BillOSale.LocEnabled = False
            BillOSale.StatusEnabled = False
            BillOSale.QuanEnabled = True       'BFH20170717 - Enabled Quan changing after tax
            BillOSale.DescEnabled = False


            'If Left(StoreSettings.Name, 7) = "Palazzo" And Val(BillOSale.Sale) > 1600 And BillOSale.X = 2 Then
            '    SalesTax1 = ((BillOSale.Sale - 1600#) * 0.06) + "132.00"
            '    BillOSale.Desc = "SALES TAX    Over $1600.00 Adjusted"
            '   Else:
            SalesTax1 = CurrencyFormat(GetStoreTax1() * (BillOSale.SubTotal(X + 1, True))) ' BillOSale.Sale - BillOSale.NonTaxable)
            'End If

            '.Price = SalesTax1
            BillOSale.SetPrice(X + 1, CurrencyFormat(SalesTax1))
            BillOSale.BalDue.Text = CurrencyFormat(GetPrice(BillOSale.BalDue.Text) + SalesTax1)
            BillOSale.PriceFocus()
            TaxApplied = "Y"
            BillOSale.StyleAddEnd()

            'see optTax2.getfocus for load
        ElseIf optTax2.Checked = True Then                        '********************************variable sales tax
            OkToProcess = True
            BillOSale.RowClear(X)
            BillOSale.RowClear(X + 1)
            BillOSale.SetStyle(X, "SUB")
            '.SetPrice X, BillOSale.Written - BillOSale.Deposit
            BillOSale.SetPrice(X, CurrencyFormat(BillOSale.SubTotal(X - 1))) 'format(BillOSale.Written - BillOSale.Deposit + SalesTax2, "###,###.00")
            BillOSale.SetDesc(X, "               Sub Total =")

            BillOSale.NewStyleLine = X + 1
            BillOSale.X = X + 1
            BillOSale.SetStyle(X + 1, "TAX2")
            BillOSale.StyleEnabled = False
            BillOSale.MfgEnabled = False
            BillOSale.LocEnabled = False
            BillOSale.StatusEnabled = False
            BillOSale.QuanEnabled = False
            BillOSale.SetQuan(X + 1, CLng(LastListClick) + 1) ' lstOptions.ListIndex + 1
            BillOSale.DescEnabled = False
            BillOSale.SetDesc(X + 1, "SALES TAX      " & LastListItemText & " = ") '& lstOptions.List(LastListClick)    ' Removed spaces for integer tax descriptions: "                     "

            GetTax()
            SalesTax2 = Rate * BillOSale.SubTotal(X + 1, True) ' (BillOSale.Sale - BillOSale.NonTaxable)
            '.Price = SalesTax2
            BillOSale.SetPrice(X + 1, CurrencyFormat(SalesTax2))
            SalesTax2 = BillOSale.QueryPrice(X + 1)
            BillOSale.BalDue.Text = CurrencyFormat(GetPrice(BillOSale.BalDue.Text) + SalesTax2 + SalesTax1)
            BillOSale.PriceFocus()
            TaxApplied = "Y"
            BillOSale.StyleAddEnd()

        ElseIf optNoTax.Checked = True Or optNoTax2.Checked = True Then                                  '******************************** No Tax
            OkToProcess = optNoTax.Checked
            Dim Which As String
            If optNoTax.Checked = True Then Which = optNoTax.Name
            If optNoTax2.Checked = True Then Which = optNoTax2.Name

            If Not CheckAccess("Give Discounts", False, True, False) Then
                If MessageBox.Show("You are not authorized to give discounts.  Request approval?", "Not Authorized", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    If Not RequestManagerApproval("Give Discounts", True) Then
                        optTax1.Select()
                        MessageBox.Show("You are not authorized to give discounts.", "Not Authorized", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit Sub
                    End If
                Else
                    optTax1.Select()
                    Exit Sub
                End If
            End If

            Dim M As String, SubTot As Double, Tax As Double, AdjTot As Double

            ' was getting cleared in requesting manager approval..  Simply reset it
            If Which = "optNoTax" Then optNoTax.Checked = True
            If Which = "optNoTax2" Then optNoTax2.Checked = True

            If optNoTax.Checked = True Then
                Xx = GetStoreTax1()
            Else
                On Error Resume Next
                M = lstOptions.Text
                If M = "" Then
                    MessageBox.Show("Please select an option from the list.", "No Rate Selected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
                M = Split(M, " ")(0)
                Xx = CDbl(M)
            End If
            SubTot = BillOSale.SubTotal(X - 1, True)
            AdjTot = Format(SubTot / (1 + Xx), "0.00")
            Tax = SubTot - AdjTot
            ' pennies... but will mess up b/c of recalculate
            'If Tax <> Xx * AdjTot Then Tax = Xx * AdjTot

            BillOSale.RowClear(X)
            BillOSale.SetStyle(X, "SUB")
            BillOSale.SetDesc(X, "TAX INCLUDED PRICE")
            BillOSale.SetPrice(X, FormatCurrency(SubTot))
            X = X + 1
            BillOSale.NewStyleLine = X
            BillOSale.X = X

            BillOSale.RowClear(X)
            BillOSale.SetStyle(X, "NOTES")
            BillOSale.SetDesc(X, "PRICE WITH TAX BACKED OUT: " & FormatCurrency(AdjTot))
            BillOSale.SetPrice(X, FormatCurrency(-Tax))
            X = X + 1
            BillOSale.NewStyleLine = X
            BillOSale.X = X

            BillOSale.RowClear(X)
            If optNoTax.Checked = True Then
                BillOSale.SetStyle(X, "TAX1")
                BillOSale.SetDesc(X, "SALES TAX")
                BillOSale.SetQuan(X, 1)
            Else
                BillOSale.SetStyle(X, "TAX2")
                BillOSale.SetDesc(X, "SALES TAX  " & lstOptions.Text)
                BillOSale.SetQuan(X, CLng(LastListClick) + 1)
            End If
            BillOSale.SetPrice(X, FormatCurrency(Tax))
            X = X + 1
            BillOSale.NewStyleLine = X
            BillOSale.X = X

            ' pennies... but will mess up b/c of recalculate
            If Format(Tax, "#.00") <> Format(Xx * AdjTot, "#.00") Then
                BillOSale.RowClear(X)
                BillOSale.SetStyle(X, "NOTES")
                BillOSale.SetDesc(X, "ADDITIONAL ADJUSTMENT")
                BillOSale.SetPrice(X, FormatCurrency(Tax - Xx * AdjTot))
                X = X + 1
                BillOSale.NewStyleLine = X
                BillOSale.X = X
            End If


            BillOSale.RowClear(X)
            BillOSale.SetStyle(X, "SUB")
            BillOSale.SetDesc(X, "TAX INCLUDED PRICE")
            BillOSale.SetPrice(X, FormatCurrency(SubTot))

            BillOSale.StyleAddEnd()
            BillOSale.StyleEnabled = False
            BillOSale.MfgEnabled = False
            BillOSale.LocEnabled = False
            BillOSale.StatusEnabled = False
            BillOSale.QuanEnabled = False
            BillOSale.DescEnabled = False
            BillOSale.PriceFocus()

        ElseIf optNotes.Checked = True Then                     '********************************Notes
            BillOSale.RowClear(X)
            BillOSale.SetStyle(X, "NOTES")
            BillOSale.StyleEnabled = False
            BillOSale.MfgEnabled = True
            BillOSale.LocEnabled = True
            BillOSale.SetLoc(X, StoresSld)
            BillOSale.StatusEnabled = False
            BillOSale.QuanEnabled = False
            BillOSale.DescEnabled = True
            BillOSale.DescFocus()
            BillOSale.SetDesc(X, "") 'NOTE: This code line is not in vb6.0 project. But added here because in Notes option, desc will be showing value 1 which is actually blank. So, to clear it added this line.
            BillOSale.StyleAddEnd()

        ElseIf optPayment.Checked = True Then                '********************************payment
            frmCCAd.Advertize()
            OkToProcess = True
            If LastListItemText = "" Then
                If MessageBox.Show("Payment Type Not Selected!", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation) = DialogResult.OK Then
                    ShowListbox()
                    '        OrdSelect.Width = 5650
                    '        lstOptions.Width = 2000
                    '        lstOptions.Visible = True
                    '        lstOptions.List(lstOptions.ListIndex) = ""  ' Don't clear the list!
                    Exit Sub
                End If
            End If

            Dim Amount As Decimal, pType As String, App As String, CCPayment As Boolean, TypeCode As Integer
            Dim TransID As String, BalanceReturned As Decimal
            Dim L As String
            Amount = GetPrice(BillOSale.BalDue.Text)
            pType = LastListItemText
            TypeCode = LastListItemData
            CCPayment = False

            If SwipeCreditCards() And IsIn(CStr(LastListItemData), "3", "4", "5", "6") Then
                If chkPayAll.Checked = True Then Amount = BillOSale.BalDue.Text
                If Not ProcessCC(Amount, pType, App, TypeCode, TransID, , BalanceReturned) Then Exit Sub
                txtPayMemo.Text = Trim(App & " " & txtPayMemo.Text)
                CCPayment = True
                BillOSale.SaleHasCCTransactions = True
            ElseIf SwipeDebitCards() And IsIn(CStr(LastListItemData), "9") Then
                If chkPayAll.Checked = True Then Amount = GetPrice(BillOSale.BalDue.Text)
                If Not ProcessDebit(Amount, pType, App, TypeCode, TransID) Then Exit Sub
                txtPayMemo.Text = Trim(App & " " & txtPayMemo.Text)
                CCPayment = True
                BillOSale.SaleHasCCTransactions = True
            ElseIf SwipeGiftCards() And IsIn(CStr(LastListItemData), "12") Then
                If chkPayAll.Checked = True Then Amount = BillOSale.BalDue.Text
                If Not ProcessGiftCard(Amount, pType, App, TypeCode) Then Exit Sub
                txtPayMemo.Text = Trim(App & " " & txtPayMemo.Text)
                CCPayment = True
                BillOSale.SaleHasCCTransactions = True
            End If

            BillOSale.RowClear(X)
            BillOSale.RowClear(X + 1)
            BillOSale.SetStyle(X, "SUB")
            '.Price = BillOSale.Written - BillOSale.Deposit
            BillOSale.SetPrice(X, CurrencyFormat(BillOSale.SubTotal(X - 1)))
            BillOSale.SetDesc(X, "               Sub Total =")

            BillOSale.NewStyleLine = X + 1
            BillOSale.X = X + 1
            If LastListItemData = 11 Then
                BillOSale.SetStyle(X + 1, "NOTES")
            Else
                BillOSale.SetStyle(X + 1, "PAYMENT")
            End If
            BillOSale.StyleEnabled = False
            BillOSale.MfgEnabled = False
            BillOSale.LocEnabled = False
            BillOSale.StatusEnabled = False
            BillOSale.SetQuan(X + 1, CStr(TypeCode)) ' Trim(Right(lstOptions.List(lstOptions.ListIndex), 2))
            BillOSale.QuanEnabled = False
            L = Microsoft.VisualBasic.Left(pType, 14) & " " & BillOSale.dteSaleDate.Value & " " & Replace(txtPayMemo.Text, "Approval", "Appr")
            '' 52 was 46 -- mjk20070724
            If IsSleepWorks Then
            Else
                If CCPayment Then L = ArrangeString(L, 52) & ArrangeString("I Authorize The Above Transaction:", 46) & "X ______________________________"
            End If
            BillOSale.SetDesc(X + 1, L)
            BillOSale.DescEnabled = False

            BillOSale.SetTransID(X + 1, TransID)

            If CCPayment Then
                BillOSale.SetPrice(X + 1, Amount)
            Else
                If Not LastListItemData = 11 Then
                    If chkPayAll.Checked = True Then BillOSale.SetPrice(X + 1, BillOSale.BalDue.Text)  'pay in full
                End If
            End If

            BillOSale.PriceFocus()
            IsStoreFinance = (LastListItemData = 11)
            BillOSale.StyleAddEnd()   ' Clears LastListItemData!!
            BillOSale.X = BillOSale.X + 1

            If BalanceReturned <> 0 Then
                BillOSale.SetStyle(X + 2, "NOTES")
                BillOSale.SetDesc(X + 2, "CARD BALANCE: " & CurrencyFormat(BalanceReturned))
                BillOSale.StyleAddEnd()   ' Clears LastListItemData!!
                BillOSale.NewStyleLine = X + 1
                BillOSale.X = X + 1
            End If


        ElseIf optStoreCredit.Checked = True Then                '********************************gift card
            BillOSale.RowClear(X)
            BillOSale.SetStyle(X, "NOTES")
            BillOSale.StyleEnabled = False
            BillOSale.MfgEnabled = False
            BillOSale.SetLoc(X, StoresSld)
            BillOSale.LocEnabled = True
            BillOSale.StatusEnabled = False
            BillOSale.QuanEnabled = False
            BillOSale.SetDesc(X, "GIFT CARD  " & Today)
            BillOSale.DescEnabled = True
            BillOSale.PriceFocus()
            BillOSale.StyleAddEnd()
        ElseIf optCarpet.Checked = True Then
            '      frmYardage.Show
            '      frmYardage.Left = 0
            '      frmYardage.Top = Screen.Height / 2 - frmYardage.Height / 2
            '
            BillOSale.LocEnabled = True
            BillOSale.StatusEnabled = True
            BillOSale.QuanEnabled = False
            BillOSale.DescEnabled = False
            Hide()
            mInvCkStyleShown = True
            If mInvCkStyle Is Nothing Then
                mInvCkStyle = New InvCkStyle
            End If
            mInvCkStyle.Show()             ' Get the new item's style number, or SS/SO/FND.
        End If

        If IsStoreFinance Then                    'Right(lstOptions.List(lstOptions.ListIndex), 2) = "11"
            If Val(BillOSale.Index) = 0 Then
                'MsgBox("You cannot set up an Installment Contract without the Customer's Name and Address!", vbCritical)
                MessageBox.Show("You cannot set up an Installment Contract without the Customer's Name and Address!", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If Len(Trim(BillOSale.CustomerPhone1.Text)) < 1 Then
                'MsgBox("You cannot set up an Installment Contract without the Customer's Telephone Number.", vbCritical)
                MessageBox.Show("You cannot set up an Installment Contract without the Customer's Telephone Number.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ArStatus = "C"  'sets status to call lease no

            '04-24-2003 - need non modal for deliver order and store finance balance
            ARPaySetUp.Show()  '-NEW 2003-02-11AA:  vbModal

            If ARPaySetUp.UnloadARPaySetUp <> "" Then
                If ARPaySetUp.UnloadARPaySetUp = True Then Exit Sub 'unload "X"
            End If
            ARPaySetUp.Show()
        End If

        ' Hide
        'Unload OrdSelect

        If OrderMode("A") Then
            If Not OkToProcess Then
                BillOSale.cmdProcessSale.Enabled = False
            End If
        End If
    End Sub

    Private Function GetTax() As Object
        Rate = 0
        Num = 0
        For TaxRate = 1 To Len(LastListItemText) + 1   'lstOptions.List(LastListClick)
            Num = Mid(LastListItemText, TaxRate, 1)     'lstOptions.List(LastListClick)
            If Num = " " Then Exit For
            Rate = Rate & Num
        Next
    End Function

    Private Sub ShowListbox(Optional ByVal UseMemoBox As Boolean = False)
        chkPayAll.Checked = False
        'Me.Width = 6030
        Me.Width = 395
        lstOptions.Visible = True
        txtPayMemo.Visible = UseMemoBox
        lblPayMemo.Visible = UseMemoBox
        cmdOk.Width = fraOpt.Width
        chkPayAll.Enabled = False
        On Error Resume Next
        lstOptions.Select()
    End Sub

    Private Sub cmdProcessSale_Click(sender As Object, e As EventArgs) Handles cmdProcessSale.Click
        'BillOSale.cmdProcessSale_Click()
        BillOSale.cmdProcessSale_Click(sender, e)
    End Sub

    Private Sub OrdSelect_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        'SetCustomFrame(Me, ncBasicTool)   ->This line is not required. It is to set font and color properties using cSkinConfiguration and cNeoCaption of modNeoCaption.
        optEnterStyle.Checked = True
        optTax1.Enabled = Not BillOSale.HasTax1
        optNoTax.Enabled = Not BillOSale.HasTax1
    End Sub

    Private Sub OrdSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim X As Integer
        'Left = (Screen.Width - Width / 2) / 2
        'Left = (Screen.PrimaryScreen.Bounds.Width - Width / 2) / 2
        'X = BillOSale.Height / 2 - Height / 2
        'If X < 0 Then X = 0
        'Top = X + 1000

        'Me.Location = New Point(100, 100)
        optEnterStyle.Checked = True
        lstOptions.Visible = False
        txtPayMemo.Visible = False
        lblPayMemo.Visible = False
        RowChangeOK = True
        '    Set mInvCkStyle = New InvCkStyle
        '    mInvCkStyle.Tag = "1"
        '    Debug.Print "Loading mInvCkStyle."
        '    mInvCkStyle.ParentForm = Name

        If Not BillOSale.PrintBill Then
            BillOSale.cmdProcessSale.Enabled = True
        End If

        If IsBFMyer Then optStain.Text = "Safeware"


        ' Disable sales tax if it's already used.
        If BillOSale.HasTax1 Then optTax1.Enabled = False '- -----> Remove this line comment later.
        '    BillOSale.HasTax1 replaces the following check...  BFH20050113
        '    Dim I As integer
        '    For I = 0 To BillOSale.UGridIO1.MaxRows - 1
        '      If Trim(BillOSale.QueryStyle(I)) = "TAX1" Then
        '        optTax1.Enabled = False
        '        Exit For
        '      End If
        '    Next

        ' Disable variable tax if it's not defined.
        optTax2.Enabled = False
        optNoTax2.Enabled = False
        If Trim(QuerySalesTax2(0)) <> "" Then optTax2.Enabled = True : optNoTax2.Enabled = True '- -----------> Moved this line To Sub New.
        If IsCanadian() Then optNoTax2.Enabled = False                                                             '----------> Remove the comment later.

        If BillOSale.IsGridFull Then                                                                               '----------> Remove the if block comment later.
            optEnterStyle.Enabled = False
            optStain.Enabled = False
            optDelivery.Enabled = False
            optLabor.Enabled = False
            optTax2.Enabled = False
            optNotes.Enabled = False
            optPayment.Enabled = False
        End If
        HideListbox()
    End Sub

    Private Sub OrdSelect_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'Note: This event is replacement for unload and queryunload events of vb6.0
        Unload_InvCkStyle()
        'RemoveCustomFrame Me
        LastListClick = 0
        LastListItemData = 0
        LastListItemText = ""
    End Sub

    Private Sub lstOptions_Click(sender As Object, e As EventArgs) Handles lstOptions.Click
        Try
            'LastListClick = lstOptions.ListIndex
            LastListClick = lstOptions.SelectedIndex
            LastListItemText = lstOptions.SelectedItem.ToString
            'LastListItemData = lstOptions.itemData(LastListClick)
            LastListItemData = CType(lstOptions.SelectedItem, ItemDataClass).ItemData

            'idc = lstOptions.Items(lstOptions.SelectedIndex)
            'LastListItemData = idc.ItemData
            'LastListItemText = lstOptions.List(LastListClick)
            'LastListItemText = lstOptions.Items(lstOptions.SelectedIndex).ToString

            'HideListbox()
        Catch ex As Exception
            LastListItemData = 0
        End Try
    End Sub

    Private Sub lstOptions_DoubleClick(sender As Object, e As EventArgs) Handles lstOptions.DoubleClick
        'cmdOk.Value = True
        'cmdOk.PerformClick()
        cmdOk_Click(cmdOk, New EventArgs)
    End Sub

    Private Sub mInvCkStyle_OKClicked(ByRef Override As Boolean, ByVal Picked As String, ByVal IsNew As Boolean) Handles mInvCkStyle.OKClicked
        Override = True
        Dim BSLine As Integer
        BSLine = BillOSale.NewStyleLine
        RowChangeOK = False
        BillOSale.X = BSLine
        BillOSale.RowClear(BSLine)
        '    BillOSale.CheckAddRow
        ' Set the style, and clear all other fields.  This allows adding lines on top of existing data.
        BillOSale.SetStyle(BSLine, Picked)
        BillOSale.SetStatus(BSLine, "")
        BillOSale.SetQuan(BSLine, "")
        BillOSale.SetDesc(BSLine, "")
        BillOSale.SetLoc(BSLine, "")
        BillOSale.SetPrice(BSLine, "", True)
        BillOSale.LoadStyle(Picked)
        '    BillOSale.Canceled = False
        If Microsoft.VisualBasic.Left(BillOSale.QueryStyle(BSLine), 4) <> KIT_PFX Then  'And Order = "A" Then
            '      BillOSale.NewStyle = mInvCkStyle.NewStyle  ' No integerer needed.
            mInvCkStyle.Hide()

            If mInvCkStyle.NewStyle Then
                ' It's a new style.  Find out what InvDefault says to do.
                Select Case InvDefault.ShowAndTell
                    Case InvDefault.ENoStyle.eNoStyle_ReEnter
                        ' OrdSelect.Show  ' Is this necessary?
                        ' Something is needed, it unloads the forms and waits for re-clicks now.
                        ' It should re-show mInvCkStyle.
                        optEnterStyle.Checked = True
                        'cmdOk.Value = True
                        'cmdOk.PerformClick()
                        cmdOk_Click(cmdOk, New EventArgs)
                        Exit Sub
                    Case InvDefault.ENoStyle.eNoStyle_NotInDBase
                        If Trim(BillOSale.QueryMfg(BillOSale.NewStyleLine - 1)) = "" Then BillOSale.MfgFocus() Else BillOSale.QuanFocus()  ' Leave focus on mfg.
                        GoTo SkipPriceFocus
                    Case InvDefault.ENoStyle.Unload_BillOfSale
                        'Unload BillOSale
                        BillOSale.Close()
                        MainMenu.Show()
                        Exit Sub
                    Case InvDefault.ENoStyle.eNoStyle_EnterItem
                        If Trim(BillOSale.QueryMfg(BillOSale.X)) = "" Then
                            BillOSale.MfgFocus()
                        Else
                            BillOSale.DescFocus()
                        End If
                        GoTo SkipPriceFocus
                    Case Else
                        GoTo SkipPriceFocus
                End Select
            Else
                ' Not a new style, so get the quantity.
                'OrdStatus.Show vbModal, MainMenu
                OrdStatus.ShowDialog(MainMenu)
            End If
        Else
            '      BillOSale.NewStyle = False
            ' It's a kit.  Quantity defaults to 1 for all items, irrelevant for the kit..
        End If
        BillOSale.PriceFocus()  ' Forces grid to display new data.
SkipPriceFocus:
        Unload_InvCkStyle()
        BillOSale.cmdProcessSale.Enabled = False

        If OrderMode("A") Then
            If Microsoft.VisualBasic.Left(BillOSale.QueryStyle(BSLine), 4) = KIT_PFX Then
                BillOSale.PriceFocus()
                Show()
            End If
        End If
        RowChangeOK = True
    End Sub

    Public Sub ShowToBillOSale2()
        On Error Resume Next
        If mInvCkStyleShown Then
            'mInvCkStyle.Show , BillOSale 
            mInvCkStyle.Show(BillOSale)
        Else
            'Show , BillOSale
            Show(BillOSale)
        End If
    End Sub

    Private Sub mInvCkStyle_CancelClicked(ByRef Override As Boolean) Handles mInvCkStyle.CancelClicked
        Override = True
        BillOSale.X = BillOSale.NewStyleLine
        '    BillOSale.Canceled = True
        BillOSale.SetDesc(BillOSale.X, "")
        BillOSale.SetLoc(BillOSale.X, "")
        BillOSale.SetMfg(BillOSale.X, "")
        BillOSale.SetMfgNo(BillOSale.X, "")
        BillOSale.SetPrice(BillOSale.X, "")
        BillOSale.SetQuan(BillOSale.X, "")
        BillOSale.SetStatus(BillOSale.X, "")
        BillOSale.SetStyle(BillOSale.X, "")
        BillOSale.StyleAddEnd(True)
        Unload_InvCkStyle()
    End Sub

    Private Sub optEnterStyle_Click(sender As Object, e As EventArgs) Handles optEnterStyle.Click
        HideListbox
    End Sub

    Private Sub HideListbox()
        lstOptions.Visible = False
        'lstOptions.Width = 0
        'Me.Width = 3885
        Me.Width = 280
        txtPayMemo.Visible = False
        lblPayMemo.Visible = False
        chkPayAll.Checked = False
        chkPayAll.Enabled = False
        cmdOk.Width = fraOpt.Width
        chkPayAll.Enabled = False
    End Sub

    Private Sub optNoTax_Click(sender As Object, e As EventArgs) Handles optNoTax.Click
        HideListbox()
    End Sub

    Private Sub optNoTax2_Click(sender As Object, e As EventArgs) Handles optNoTax2.Click
        If optNoTax2.Checked = True Then
            LoadSalesTaxOptionsIntoListbox(lstOptions)
            ShowListbox()
        Else
            HideListbox()
        End If
    End Sub

    Private Sub LoadSalesTaxOptionsIntoListbox(ByRef lst As ListBox)
        'lst.Clear
        lst.Items.Clear()
        Dim El As Object
        For Each El In QuerySalesTax2List()
            'lst.AddItem El
            lst.Items.Add(El)
        Next
        On Error Resume Next
        'lst.Selected(0) = True
        lst.SetSelected(0, True)
    End Sub

    Private Sub optStain_Click(sender As Object, e As EventArgs) Handles optStain.Click
        HideListbox()
    End Sub

    Private Sub optDelivery_Click(sender As Object, e As EventArgs) Handles optDelivery.Click
        HideListbox()
    End Sub

    Private Sub optLabor_Click(sender As Object, e As EventArgs) Handles optLabor.Click
        HideListbox()
    End Sub

    Private Sub optStoreCredit_Click(sender As Object, e As EventArgs) Handles optStoreCredit.Click
        HideListbox()
    End Sub

    Private Sub optTax1_Click(sender As Object, e As EventArgs) Handles optTax1.Click
        HideListbox()
    End Sub

    Private Sub optTax2_Click(sender As Object, e As EventArgs) Handles optTax2.Click
        'variable sales tax
        If optTax2.Checked = True Then
            LoadSalesTaxOptionsIntoListbox(lstOptions)
            ShowListbox()

            'If BillOSale.cboTaxZone.ListIndex <> -1 Then
            '    lstOptions.ListIndex = BillOSale.cboTaxZone.ListIndex
            'End If
            If BillOSale.cboTaxZone.SelectedIndex <> -1 Then
                lstOptions.SelectedIndex = BillOSale.cboTaxZone.SelectedIndex
            End If
        Else
            HideListbox()
        End If
    End Sub

    Private Sub optNotes_Click(sender As Object, e As EventArgs) Handles optNotes.Click
        HideListbox()
    End Sub

    Private Sub optPayment_Click(sender As Object, e As EventArgs) Handles optPayment.Click
        If optPayment.Checked = True Then
            LoadPaymentOptionsIntoListbox(lstOptions)
            ShowListbox(True)
        Else
            HideListbox()
        End If
        chkPayAll.Enabled = True 'pay all
    End Sub

    Private Sub lstOptions_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstOptions.SelectedIndexChanged
        lstOptions_Click(lstOptions, New EventArgs)
    End Sub
End Class