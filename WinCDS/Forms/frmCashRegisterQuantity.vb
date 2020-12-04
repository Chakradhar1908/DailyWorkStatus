Imports Microsoft.VisualBasic.Interaction
Public Class frmCashRegisterQuantity
    ' DEBIT TEST SWIPE (DOUBLE CLICK TOP LEFT CORNER OF SWIPE SCREEN, COPY THIS (WITHOUT LEADING '), AND RIGHT CLICK - PASTE INTO SWIPE BOX
    '%B5499990123456781^TESTCARD/MASTERCARD/NDCECOMMERCE^0908101543219871234567890?
    ';5499990123456781=09081015432198712345?
    '  CVC 123
    '  Pin 1234

    ' Fake card:
    '%B4242424242424242^TOTALLY/FAKE^0101555000011112222333344?;5499990123456781=09081015432198712345?
    '
    ' Transaction Central card, Made up swipe
    '%B4111111111111111^TRANSACTION/CENTRAL^1201555000011112222333344?;4111111111111111=12011015432198712345?
    '   CVV 123
    '   EXP (MMDD): 0112
    '
    'MC:   %B5400111122223332^KATZ/GERALD             ^140810100000008802000000115000000?;5400111122223332=1408101088020115?
    'AMEX: %B370000000000002^KATZ/GA                   ^1503110362315?;370000000000002=1503110362315?

    Public DBG As String
    Public RefId As String
    Private tPayType As String
    Private Quantity As Double, Total As Decimal, Taxable As Boolean, Cancelled As Boolean
    Dim XC As clsXCharge, XCXL As clsXChargeXpressLink, TC As clsTransactionCentral, cM As clsCredomatic, cP As clsChargeItPro

    Public Function GetQuantityAndPrice(ByVal Style As String, ByVal Desc As String, ByVal Price As Decimal, ByVal NonTax As Boolean) As clsSaleItem
        Dim CCFunctions() As Object
        If SwipeCreditCards() Then ArrAdd(CCFunctions, "3")    '  3 = Credit
        If SwipeDebitCards() Then ArrAdd(CCFunctions, "9")     '  9 = Debit
        If SwipeGiftCards() Then ArrAdd(CCFunctions, "12")     ' 12 = Store Credit / Gift Card

        Cancelled = False
        Quantity = 0
        Total = 0
        cboChargeType.Visible = False   ' reset all control states to 'normal'
        lblItem.Visible = True
        lblDescCap.Visible = True
        lblDesc.Visible = True
        lblSwipe.Visible = False
        cmdApply.Enabled = True
        cmdApply.Visible = True
        cmdSwipe.Visible = False
        txtSwipe.Visible = False
        txtPrice.ReadOnly = False
        'cmdApply.Default = True
        Me.AcceptButton = cmdApply
        fraDiscType.Visible = False


        If Style = "PAYMENT" Then
            If Desc = "3" Then  ' This only happens for Charge payments.
                LoadCreditCardTypes()
                cboChargeType.Visible = True
                lblItem.Visible = False
                lblDescCap.Visible = False
            Else
                'lblStyle.Caption = QueryPaymentDescription(Desc)
                lblStyle.Text = QueryPaymentDescription(Desc)
            End If

            If IsInArray(Desc, CCFunctions) Then ' FOR CREDIT OR DEBIT
                cboChargeType.Visible = False
                lblSwipe.Visible = True
                lblDescCap.Visible = False
                cmdApply.Enabled = False
                cmdApply.Visible = False
                cmdSwipe.Visible = True
                If SwipeOnThisForm Then
                    txtSwipe.Visible = True
                    shpSwipe.Visible = True
                    shpSwipe.FillColor = Color.Green
                End If
            End If

            'lblStyle.Caption = Desc
            'txtPrice.ToolTipText = "Enter the payment amount here."
            ToolTip1.SetToolTip(txtPrice, "Enter the payment amount here.")
            lblPrice.Text = "Payment:"
            cmdTax.Visible = False
            Taxable = False
            lblQuan.Visible = False
            txtQuantity.Visible = False
            txtQuantity.Text = Desc
            SelectContents(txtPrice)
            'cmdCancel.Move cmdCancel.Left, cmdTax.Top
            cmdCancel.Location = New Point(cmdCancel.Left, cmdTax.Top)
            'cmdApply.Move cmdApply.Left, cmdTax.Top
            cmdApply.Location = New Point(cmdApply.Left, cmdTax.Top)
            'cmdSwipe.Move cmdApply.Left, cmdApply.Top
            cmdSwipe.Location = New Point(cmdApply.Left, cmdApply.Top)
            'txtSwipe.Move cmdApply.Left, cmdApply.Top
            txtSwipe.Location = New Point(cmdApply.Left, cmdApply.Top)
            'lblPrice.Move lblPrice.Left, lblQuan.Top
            lblPrice.Location = New Point(lblPrice.Left, lblQuan.Top)
            'txtPrice.Move txtPrice.Left, txtQuantity.Top
            txtPrice.Location = New Point(txtPrice.Left, txtQuantity.Top)
            Height = Height - txtPrice.Height
        ElseIf Style = "DISCOUNT" Then
            lblStyle.Text = Style
            lblPrice.Text = "Percentage:"
            'txtPrice.ToolTipText = "Enter the discount percentage here."
            ToolTip1.SetToolTip(txtPrice, "Enter the discount percentage here.")
            fraDiscType.Visible = True
            optDiscType0.Checked = True
            cmdTax.Visible = False
            Taxable = False
            lblQuan.Visible = False
            txtQuantity.Visible = False
            txtQuantity.Text = "1"
            SelectContents(txtPrice)
            'cmdCancel.Move cmdCancel.Left, cmdTax.Top
            cmdCancel.Location = New Point(cmdCancel.Left, cmdTax.Top)
            'cmdApply.Move cmdApply.Left, cmdTax.Top
            cmdApply.Location = New Point(cmdApply.Left, cmdTax.Top)
            'cmdSwipe.Move cmdApply.Left, cmdApply.Top
            cmdSwipe.Location = New Point(cmdApply.Left, cmdApply.Top)
            'lblPrice.Move lblPrice.Left, lblQuan.Top
            lblPrice.Location = New Point(lblPrice.Left, lblQuan.Top)
            'txtPrice.Move txtPrice.Left, txtQuantity.Top
            txtPrice.Location = New Point(txtPrice.Left, txtQuantity.Top)
            Height = Height - txtPrice.Height
        Else
            lblStyle.Text = Style                ' Show item style number.
            lblDesc.Text = Desc                  ' Show item description.
            If NonTax Then cmdTax.Visible = False : Taxable = False Else Taxable = True
            Quantity = 0                            ' Default hidden quantity to 0.  This is used if we cancel the form.
            txtQuantity.Text = "1"                  ' Default visible quantity to 1.
            SelectContents(txtQuantity)              ' Select the quantity, for easy replacement.
            'txtPrice.ToolTipText = "Enter the price of one item.  Regular price is shown by default."
            ToolTip1.SetToolTip(txtPrice, "Enter the price of one item.  Regular price is shown by default.")
        End If

        txtPrice.Text = CurrencyFormat(Price)
        txtPrice.Tag = txtPrice.Text

        DBG = "001"
Again:
        DBG = "002"
        GetQuantityAndPrice = Nothing
        DBG = "003"
        tPayType = Desc
        DBG = "004"


        txtSwipe.Text = ""
        DBG = "005"
        ResetSwipeTimer()
        DBG = "006"
        'Show vbModal                 ' Show on top of POS form.  Should this be a generic parent?
        ShowDialog()
        DBG = "007"
        ResetSwipeTimer(True)
        DBG = "008"

        If Not Cancelled Then
            DBG = "009"
            GetQuantityAndPrice = New clsSaleItem
            DBG = "010"
            GetQuantityAndPrice.Style = Style
            DBG = "011"
            If Style = "PAYMENT" Then
                DBG = "012"
                If IsInArray(Desc, CCFunctions) Then
                    DBG = "013"
                    If StoreSettings.CCProcessor = CCPROC_XC Then
                        DBG = "014"
                        XC = New clsXCharge
                        DBG = "015"
                        XC.FormHandle = 0
                        DBG = "016"
                        XC.Clerk = "Cash Register"
                        DBG = "017"
                        XC.Receipt = "000"
                        DBG = "018"
                        XC.Amount = Total
                        DBG = "019"
                        If Desc = "3" Then
                            DBG = "020"
                            If Not XC.ExecPurchase(True) Then GoTo Again ' Set GetQuantityAndPrice = Nothing: Unload Me: Exit Function 'GoTo Again
                            DBG = "021"
                            GetQuantityAndPrice.Desc = XC.CCTypeName
                            GetQuantityAndPrice.Balance = GetPrice(XC.BalanceAmountResult)
                            Total = XC.Amount                             ' partial approval says amount can change
                            DBG = "022"
                        ElseIf Desc = "9" Then
                            If Not XC.ExecDebitPurchase(True) Then GoTo Again 'Set GetQuantityAndPrice = Nothing: Unload Me: Exit Function 'GoTo Again
                            GetQuantityAndPrice.Desc = XC.CCTypeName & " DEBIT"
                            Total = XC.Amount                             ' partial approval says amount can change
                        ElseIf Desc = "12" Then
                            If Not XC.ExecGiftRedeem(True) Then GoTo Again ' Set GetQuantityAndPrice = Nothing: Unload Me: Exit Function 'GoTo Again
                            GetQuantityAndPrice.Desc = XC.CCTypeName & " GIFT"
                        End If
                        GetQuantityAndPrice.Extra1 = XC.XCC & "  Approval=" & XC.ApprovalCode
                        DisposeDA(XC)
                    ElseIf StoreSettings.CCProcessor = CCPROC_XL Then
                        XCXL = New clsXChargeXpressLink
                        XCXL.Receipt = NextReceiptNumber()
                        If IsFormLoaded("BillOSale") Then
                            XCXL.Zip = BillOSale.CustomerZip.Text
                            XCXL.Address = BillOSale.CustomerAddress.Text
                        End If
                        XCXL.Amount = Total
                        If Desc = "3" Then
                            If Not XCXL.ExecPurchase() Then GoTo Again ' Set GetQuantityAndPrice = Nothing: Unload Me: Exit Function 'GoTo Again
                            GetQuantityAndPrice.Desc = XCXL.CCTypeName
                            GetQuantityAndPrice.Balance = GetPrice(XCXL.BalanceAmountResult)
                            Total = XCXL.Amount                             ' partial approval says amount can change
                        ElseIf Desc = "9" Then
                            If Not XCXL.ExecDebitPurchase() Then GoTo Again 'Set GetQuantityAndPrice = Nothing: Unload Me: Exit Function 'GoTo Again
                            GetQuantityAndPrice.Desc = XCXL.CCTypeName & " DEBIT"
                            Total = XCXL.Amount                             ' partial approval says amount can change
                        ElseIf Desc = "12" Then
                            If Not XCXL.ExecGiftRedeem() Then GoTo Again ' Set GetQuantityAndPrice = Nothing: Unload Me: Exit Function 'GoTo Again
                            GetQuantityAndPrice.Desc = XCXL.CCTypeName & " GIFT"
                        End If
                        GetQuantityAndPrice.Extra1 = XCXL.XCC & "  Approval=" & XCXL.ApprovalCode
                        DisposeDA(XCXL)
                    ElseIf StoreSettings.CCProcessor = CCPROC_TC Then
                        TC = New clsTransactionCentral
                        TC.Amount = Total
                        TC.RefId = RefId
                        If txtSwipe.Text <> "" Then
                            ' BFH20110405 - REQUIRES TRACK 2 NOW, CUZ TRANSACTION CENTRAL CAN'T READ TRACK 1
                            If Not CreditCardSwipeValid(txtSwipe.Text, True) Then
                                MessageBox.Show("Could not get card data.", "Did not read card", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                GoTo Again
                            End If
                        End If
                        DBG = "022"

                        TC.BlindSwipe(txtSwipe.Text)
                        DBG = "023"
                        If Desc = "3" Then
                            If Not TC.ExecPurchase(txtSwipe.Text = "") Then GoTo Again
                            GetQuantityAndPrice.Desc = TC.CCTypeName
                        ElseIf Desc = "9" Then
                            If Not TC.ExecDebitPurchase(txtSwipe.Text = "") Then GoTo Again
                            GetQuantityAndPrice.Desc = TC.CCTypeName & " DEBIT"
                        ElseIf Desc = "12" Then
                            If Not TC.ExecGiftRedeem(txtSwipe.Text = "") Then GoTo Again
                            GetQuantityAndPrice.Desc = TC.CCTypeName & " GIFT"
                        End If
                        GetQuantityAndPrice.Extra1 = TC.XCC & "  Approval=" & TC.ApprovalCode
                        GetQuantityAndPrice.TransID = TC.TransID
                        DisposeDA(TC)
                    ElseIf StoreSettings.CCProcessor = CCPROC_CM Then
                        '          Set cM = New clsCredoMatic
                        '          cM.Amount = Total
                        '          cM.RefID = RefID
                        '          If txtSwipe.Text <> "" Then
                        '' BFH20110405 - REQUIRES TRACK 2 NOW, CUZ TRANSACTION CENTRAL CAN'T READ TRACK 1
                        '            If Not CreditCardSwipeValid(txtSwipe.Text, True) Then
                        '              MsgBox "Could not get card data.", vbExclamation, "Did not read card"
                        '              GoTo Again
                        '            End If
                        '          End If
                        '
                        '          cM.BlindSwipe txtSwipe
                        '          If Desc = "3" Then
                        '            If Not cM.ExecPurchase(txtSwipe = "") Then GoTo Again
                        '            GetQuantityAndPrice.Desc = cM.CCTypeName
                        '          ElseIf Desc = "9" Then
                        '            If Not cM.ExecDebitPurchase(txtSwipe = "") Then GoTo Again
                        '            GetQuantityAndPrice.Desc = cM.CCTypeName & " DEBIT"
                        '          ElseIf Desc = "12" Then
                        '            If Not cM.ExecGiftRedeem(txtSwipe = "") Then GoTo Again
                        '            GetQuantityAndPrice.Desc = cM.CCTypeName & " GIFT"
                        '          End If
                        '          GetQuantityAndPrice.Extra1 = cM.XCC & "  Approval=" & cM.ApprovalCode
                        '          GetQuantityAndPrice.TransID = cM.TransID
                        '          DisposeDA cM
                    ElseIf StoreSettings.CCProcessor = CCPROC_CI Then
                        cP = New clsChargeItPro
                        cP.Amount = Total
                        If Desc = "3" Then
                            If Not cP.ExecPurchase(True) Then GoTo Again ' Set GetQuantityAndPrice = Nothing: Unload Me: Exit Function 'GoTo Again
                            GetQuantityAndPrice.Desc = cP.CCTypeName
                            GetQuantityAndPrice.Balance = GetPrice(cP.BalanceAmountResult)
                            Total = cP.Amount                             ' partial approval says amount can change
                        ElseIf Desc = "9" Then
                            If Not cP.ExecDebitPurchase(True) Then GoTo Again 'Set GetQuantityAndPrice = Nothing: Unload Me: Exit Function 'GoTo Again
                            GetQuantityAndPrice.Desc = cP.CCTypeName & " DEBIT"
                            Total = cP.Amount                             ' partial approval says amount can change
                        ElseIf Desc = "12" Then
                            If Not cP.ExecGiftRedeem(True) Then GoTo Again ' Set GetQuantityAndPrice = Nothing: Unload Me: Exit Function 'GoTo Again
                            GetQuantityAndPrice.Desc = cP.CCTypeName & " GIFT"
                        End If
                        GetQuantityAndPrice.Extra1 = cP.XCC & "  Approval=" & cP.ApprovalCode
                        GetQuantityAndPrice.TransID = cP.TransIDResult
                        DisposeDA(cP)
                    End If
                Else
                    GetQuantityAndPrice.Desc = QueryPaymentDescription(Quantity)
                End If
            Else
                GetQuantityAndPrice.Desc = Desc
            End If
            DBG = "024"
            GetQuantityAndPrice.Quantity = Quantity       ' Return the form's quantity value.  This is set to nonzero by cmdApply, or zero otherwise.

            If GetQuantityAndPrice.Style = "DISCOUNT" Then
                'GetQuantityAndPrice.Quantity = Switch(optDiscType2, 2, optDiscType1, 1, True, 0)
                GetQuantityAndPrice.Quantity = Switch(optDiscType2.Checked = True, 2, optDiscType1.Checked = True, 1, True, 0)
            End If
            DBG = "025"
            GetQuantityAndPrice.Price = Total             ' Also return the form's price value. This is also set by cmdApply.
            GetQuantityAndPrice.DisplayPrice = Total
            GetQuantityAndPrice.NonTaxable = Not Taxable
            Taxable = True
        End If
        DBG = "026"
        Desc = ""
        DBG = "028"
        'Unload Me
        Me.Close()
        DBG = "027"
    End Function

    Private Sub LoadCreditCardTypes()
        cboChargeType.Items.Clear()
        AddItemToComboBox(cboChargeType, "VISA CARD", 3)
        AddItemToComboBox(cboChargeType, "MASTER CARD", 4)
        AddItemToComboBox(cboChargeType, "DISCOVER CARD", 5)
        AddItemToComboBox(cboChargeType, "AMEX", 6)
        'cboChargeType.ListIndex = 0
        cboChargeType.SelectedIndex = 0
    End Sub

    Private ReadOnly Property SwipeOnThisForm() As Boolean
        Get
            SwipeOnThisForm = StoreSettings.bUseCCMachine And IsIn(StoreSettings.CCProcessor, CCPROC_TC, CCPROC_CM)
        End Get
    End Property

    Private Sub ResetSwipeTimer(Optional ByVal Disable As Boolean = False)
        tmrSwipe.Enabled = False
        If Disable Then Exit Sub
        tmrSwipe.Interval = 2000
        tmrSwipe.Enabled = True
    End Sub

    Public Function DoReturn(ByVal Price As Decimal, ByVal PayType As String, ByVal TransID As String, ByVal SaleDate As String) As clsSaleItem
        Dim CCFunctions() As Object
        ArrAdd(CCFunctions, "")
        If SwipeCreditCards() Then ArrAdd(CCFunctions, "3")    '  3 = Credit
        ' Debit Cards don't do returns ---
        '       Yes, I am serious.  We do not have any control over this.  It is dictated by the processors (Global Payments, TransFirst, etc).
        '       There is really nothing for us to 'work on.'   As I stated, it is a 'rule' set by the processors.
        '       Julia Chen | X-Charge Integrations Manager
        '  If SwipeDebitCards Then ArrAdd CCFunctions, "9"     '  9 = Debit
        If SwipeGiftCards() Then ArrAdd(CCFunctions, "12")     ' 12 = Store Credit / Gift Card

        Dim AsCredit As Boolean, AsDebit As Boolean, AsGiftCard As Boolean
        Dim PayDesc As String

        AsCredit = IsIn(PayType, "3", "4", "5", "6")
        AsDebit = IsIn(PayType, "9")
        AsGiftCard = IsIn(PayType, "12")
        If PayType = "3" Then
            PayDesc = "CREDIT CARD"
        Else
            PayDesc = QueryPaymentDescription(Val(PayType))
        End If

        Cancelled = False
        Quantity = 0
        Total = 0

        cboChargeType.Visible = False
        lblItem.Visible = True

        lblDescCap.Visible = False
        lblDesc.Visible = False

        lblStyle.Text = PayDesc

        lblSwipe.Visible = False
        cmdApply.Enabled = True
        cmdApply.Visible = True
        cmdSwipe.Visible = False
        txtSwipe.Visible = False
        txtPrice.ReadOnly = True

        'txtPrice.ToolTipText = "Enter the payment amount here."
        ToolTip1.SetToolTip(txtPrice, "Enter the payment amount here.")
        lblPrice.Text = "Return:"
        cmdTax.Visible = False
        Taxable = False
        lblQuan.Visible = False
        txtQuantity.Visible = False

        If IsInArray(PayType, CCFunctions) Then
            cboChargeType.Visible = False
            lblSwipe.Visible = True
            lblSwipe.Text = "Click Swipe to Return"
            lblDescCap.Visible = False
            If SwipeOnThisForm Then txtSwipe.Visible = True
            cmdApply.Visible = False
            cmdSwipe.Visible = True
        Else
            If AsCredit Then
                cboChargeType.Visible = True
                LoadCreditCardTypes()
            End If
        End If

        SelectContents(txtPrice.Text)
        'cmdCancel.Move(cmdCancel.Left, cmdTax.Top)
        cmdCancel.Location = New Point(cmdCancel.Left, cmdTax.Top)
        'cmdApply.Move cmdApply.Left, cmdTax.Top
        cmdApply.Location = New Point(cmdApply.Left, cmdTax.Top)
        'txtSwipe.Move cmdApply.Left, cmdApply.Top
        txtSwipe.Location = New Point(cmdApply.Left, cmdApply.Top)
        'cmdSwipe.Move cmdApply.Left, cmdApply.Top
        cmdSwipe.Location = New Point(cmdApply.Left, cmdApply.Top)
        'lblPrice.Move lblPrice.Left, lblQuan.Top
        lblPrice.Location = New Point(lblPrice.Left, lblQuan.Top)
        'txtPrice.Move txtPrice.Left, txtQuantity.Top
        txtPrice.Location = New Point(txtPrice.Left, txtQuantity.Top)
        Height = Height - txtPrice.Height

        txtPrice.Text = CurrencyFormat(Price)

Again:
        DoReturn = Nothing
        txtSwipe.Text = ""
        'Show vbModal                 ' Show on top of POS form.  Should this be a generic parent?
        ShowDialog()

        If Not Cancelled Then
            DoReturn = New clsSaleItem
            DoReturn.Style = "PAYMENT"

            If IsInArray(PayType, CCFunctions) Then
                If StoreSettings.CCProcessor = CCPROC_XC Then
                    XC = New clsXCharge
                    XC.FormHandle = 0
                    XC.Clerk = "Cash Register"
                    XC.Receipt = "000"
                    XC.Amount = Total
                    If PayType = "3" Then
                        If Not XC.ExecReturn(True) Then GoTo Again
                        Total = XC.Amount
                        DoReturn.Balance = GetPrice(XC.BalanceAmountResult)
                    ElseIf PayType = "9" Then
                        If Not XC.ExecDebitReturn(True) Then GoTo Again
                    ElseIf PayType = "12" Then
                        If Not XC.ExecGiftReturn(True) Then GoTo Again
                    End If
                    DoReturn.Desc = XC.CCTypeName
                    DoReturn.Extra1 = XC.XCC & "  Approval=" & XC.ApprovalCode
                    DisposeDA(XC)
                ElseIf StoreSettings.CCProcessor = CCPROC_XL Then
                    XCXL = New clsXChargeXpressLink
                    XCXL.Receipt = NextReceiptNumber()
                    If IsFormLoaded("BillOSale") Then
                        XCXL.Zip = BillOSale.CustomerZip.Text
                        XCXL.Address = BillOSale.CustomerAddress.Text
                    End If
                    XCXL.Amount = Total
                    If PayType = "3" Then
                        If Not XCXL.ExecReturn(True) Then GoTo Again
                        Total = XCXL.Amount
                        DoReturn.Balance = GetPrice(XCXL.BalanceAmountResult)
                    ElseIf PayType = "9" Then
                        If Not XCXL.ExecDebitReturn() Then GoTo Again
                    ElseIf PayType = "12" Then
                        If Not XCXL.ExecGiftReturn() Then GoTo Again
                    End If
                    DoReturn.Desc = XCXL.CCTypeName
                    DoReturn.Extra1 = XCXL.XCC & "  Approval=" & XCXL.ApprovalCode
                    DisposeDA(XCXL)
                ElseIf StoreSettings.CCProcessor = CCPROC_TC Then
                    TC = New clsTransactionCentral
                    TC.Clerk = "Cash Register"
                    TC.Receipt = "000"
                    TC.Amount = Total
                    TC.TransID = TransID
                    TC.RefId = RefId
                    If txtSwipe.Text <> "" Then
                        'ActiveLog "frmCashRegisterQuantity::DoReturn: txtSwipe=" & txtSwipe.Text, 9
                        If Not CreditCardSwipeValid(txtSwipe.Text) Then
                            MessageBox.Show("Could not get card data.", "Did not read card", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            GoTo Again
                        End If
                        TC.BlindSwipe(txtSwipe.Text)
                    End If
                    If PayType = "3" Then
                        If Not TC.ExecBlindCredit(txtSwipe.Text = "") Then GoTo Again
                    ElseIf PayType = "9" Then
                        '          If Not TC.ExecDebitReturn(True) Then GoTo Again
                    ElseIf PayType = "12" Then
                        '          If Not TC.ExecGiftReturn(True) Then GoTo Again
                    End If
                    DoReturn.Desc = TC.CCTypeName
                    DoReturn.Extra1 = TC.XCC & "  Approval=" & TC.ApprovalCode
                    DoReturn.TransID = TC.TransID
                    DisposeDA(TC)
                ElseIf StoreSettings.CCProcessor = CCPROC_CM Then
                    cM = New clsCredomatic
                    cM.Amount = Total
                    If txtSwipe.Text <> "" Then
                        'ActiveLog "frmCashRegisterQuantity::DoReturn: txtSwipe=" & txtSwipe.Text, 9
                        If Not CreditCardSwipeValid(txtSwipe.Text) Then
                            MessageBox.Show("Could not get card data.", "Did not read card", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            GoTo Again
                        End If
                    End If
                    If PayType = "3" Then
                        If Not cM.ExecVoid(SaleDate) Then GoTo Again
                    ElseIf PayType = "9" Then
                        '          If Not cm.ExecDebitReturn(True) Then GoTo Again
                    ElseIf PayType = "12" Then
                        '          If Not cm.ExecGiftReturn(True) Then GoTo Again
                    End If
                    DoReturn.Desc = cM.CCTypeName
                    DoReturn.Extra1 = cM.XCC & "  Approval=" & cM.ApprovalCode
                    DoReturn.TransID = cM.TransID
                    DisposeDA(cM)
                ElseIf StoreSettings.CCProcessor = CCPROC_CI Then
                    cP = New clsChargeItPro
                    cP.FormHandle = 0
                    cP.Clerk = "Cash Register"
                    cP.Receipt = "000"
                    cP.Amount = Total
                    If PayType = "3" Then
                        If Not cP.ExecReturn(True) Then GoTo Again
                        Total = cP.Amount
                        DoReturn.Balance = GetPrice(cP.BalanceAmountResult)
                    ElseIf PayType = "9" Then
                        If Not cP.ExecDebitReturn(True) Then GoTo Again
                    ElseIf PayType = "12" Then
                        If Not cP.ExecGiftReturn(True) Then GoTo Again
                    End If
                    DoReturn.Desc = cP.CCTypeName
                    DoReturn.Extra1 = cP.XCC & "  Approval=" & cP.ApprovalCode
                    DisposeDA(cP)
                End If
            Else
                DoReturn.Desc = QueryPaymentDescription(PayType)
            End If

            DoReturn.Quantity = 1 'Quantity       ' Return the form's quantity value.  This is set to nonzero by cmdApply, or zero otherwise.
            ' BFH20080108 Negated price because it came out wrong...  Did this affect anything else?  hard to imagine this went this long w/o being noticed...
            DoReturn.Price = -Total             ' Also return the form's price value. This is also set by cmdApply.
            DoReturn.DisplayPrice = -Total
            DoReturn.NonTaxable = Not Taxable
            Taxable = True
        End If
        'Unload Me
        Me.Close()
    End Function

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        If Not ValidateQuantity(txtQuantity.Text, False) Then Exit Sub
        If Not ValidatePrice(txtPrice.Text, False) Then Exit Sub
        If cboChargeType.Visible Then
            'Quantity = cboChargeType.itemData(cboChargeType.ListIndex)
            Quantity = CType(cboChargeType.Items(cboChargeType.SelectedIndex), ItemDataClass).ItemData
        Else
            Quantity = GetPrice(txtQuantity.Text)      ' This is the only place quantity is set to non-zero.
        End If
        Total = GetPrice(txtPrice.Text)              ' This is the only place price is set.
        Cancelled = False
        Hide()                                    ' And unload the form to clean up.

    End Sub

    Private Function ValidateQuantity(ByVal Quan As String, ByVal Noisy As Boolean) As Boolean
        If Trim(Quan) = "" Then Exit Function ' Blanks are okay, they translate to zero.
        If Not IsNumeric(Quan) Then           ' Reject non-numeric quantities.
            ValidateQuantity = False
            If Noisy Then MessageBox.Show("Invalid quantity.  Please enter a number.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ElseIf GetPrice(Quan) < 0 Then        ' Reject negative quantities.
            ValidateQuantity = False
            If Noisy Then MessageBox.Show("Invalid quantity.  If you need to enter a return, use the Returns button on the Point of Sale screen.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            ValidateQuantity = True             ' All positive quantities are okay.  It is possible to sell fractional units.
        End If
    End Function

    Private Function ValidatePrice(ByVal Prc As String, Optional ByVal Noisy As Boolean = False) As Boolean
        If Trim(Prc) = "" Then ValidatePrice = True : Exit Function ' Blanks are okay.
        If Not IsNumeric(txtPrice.Text) Then
            ValidatePrice = False
            If Noisy Then MessageBox.Show("Invalid price.  Please enter a number.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ElseIf GetPrice(txtPrice.Text) < 0 And cboChargeType.Visible = False Then
            ValidatePrice = False
            If Noisy Then MessageBox.Show("Invalid price.  If you need to enter a return, use the Returns button on the Point of Sale screen.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            ValidatePrice = True
        End If
    End Function

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Quantity = 0
        Total = 0
        Taxable = True
        Cancelled = True
        Hide()                             ' True Quantity is never set unless the user hits Apply.
    End Sub

    Private Sub cmdTax_Click(sender As Object, e As EventArgs) Handles cmdTax.Click
        If cmdTax.Text = "&Taxable" Then
            cmdTax.Text = "No &Tax"
            'cmdTax.ToolTipText = "Item is nontaxable.  Click here to add tax to item."
            ToolTip1.SetToolTip(cmdTax, "Item is nontaxable.  Click here to add tax to item.")
            Taxable = False
        Else
            cmdTax.Text = "&Taxable"
            'cmdTax.ToolTipText = "Item is taxable.  Click here to remove tax from item."
            ToolTip1.SetToolTip(cmdTax, "Item is taxable.  Click here to remove tax from item.")
            Taxable = True
        End If
    End Sub

    Private Sub frmCashRegisterQuantity_Load(sender As Object, e As EventArgs) Handles Me.Load
        SetButtonImage(cmdApply, 2)
        SetButtonImage(cmdCancel, 3)
        SetButtonImage(cmdSwipe, 2)
        If IsIn(StoreSettings.CCProcessor, CCPROC_XC, CCPROC_XL, CCPROC_CI) Then
            lblSwipe.Text = "Click Swipe to Charge"
        ElseIf IsIn(StoreSettings.CCProcessor, CCPROC_TC, CCPROC_CM) Then
            lblSwipe.Text = "Swipe Card to Charge"
            cmdSwipe.Text = "Manual"
            'cmdSwipe.ToolTipText = "Click to manually enter the CC."
            ToolTip1.SetToolTip(cmdSwipe, "Click to manually enter the CC.")
        End If
        fraDiscType.Visible = False
    End Sub

    Private Sub tmrSwipe_Tick(sender As Object, e As EventArgs) Handles tmrSwipe.Tick
        ResetSwipeTimer(True)

        ' BFH20110405
        ' Require track 2 since we have to send it...
        ' Was: CreditCardSwipeValid(txtSwipe, IsIn(tPayType, "9"))
        If CreditCardSwipeValid(txtPrice.Text, True) Then
            txtSwipe.Text = txtPrice.Text
            txtPrice.Text = txtPrice.Tag
        End If

        If CreditCardSwipeValid(txtSwipe.Text, True) Then
            SwipeDone()
        Else
            If txtSwipe.Text <> "" Then MessageBox.Show("Invalid swipe.  Please try again.", "WinCDS")
            SwipeDone(True)
        End If
    End Sub

    Private Sub imgCheat_Click(sender As Object, e As EventArgs) Handles imgCheat.Click
        If Not IsDevelopment() Then Exit Sub
        If txtSwipe.Left = 0 Then
            'txtSwipe.Move cmdApply.Left, cmdApply.Top, cmdApply.Width
            txtSwipe.Location = New Point(cmdApply.Left, cmdApply.Top)
            txtSwipe.Size = New Size(cmdApply.Top, cmdApply.Width)
        Else
            'txtSwipe.Move 0, 150, ScaleWidth
            txtSwipe.Location = New Point(0, 15)
            txtSwipe.Size = New Size(Me.ClientSize.Width, txtSwipe.Height)
        End If
    End Sub

    Private Sub tmrLocate_Tick(sender As Object, e As EventArgs) Handles tmrLocate.Tick
        On Error Resume Next
        If txtSwipe.Text <> "" Then Exit Sub
        If txtSwipe.Visible Then txtPrice.Select()
    End Sub

    'Private Sub optDiscType_Click(Index As Integer)  -> Below is the replacement for this click event.
    Private Sub optDiscTypeCheckedChanged(sender As Object, e As EventArgs) Handles optDiscType0.CheckedChanged, optDiscType1.CheckedChanged, optDiscType2.CheckedChanged
        'If Index = 0 Then
        '    lblPrice = "Percentage:"
        'ElseIf Index = 1 Then
        '    lblPrice = "Percentage:"
        'Else
        '    lblPrice = "Amount:"
        'End If

        Dim optDiscType As String
        optDiscType = CType(sender, RadioButton).Name
        If optDiscType = "optDiscType0" Then
            lblPrice.Text = "Percentage:"
        ElseIf optDiscType = "optDiscTyp1" Then
            lblPrice.Text = "Percentage:"
        ElseIf optDiscType = "optDiscType2" Then
            lblPrice.Text = "Amount:"
        End If
    End Sub

    Private Sub txtSwipe_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSwipe.KeyPress
        'If KeyAscii = 13 Or KeyAscii = 10 Then KeyAscii = 0
        If e.KeyChar = Chr(13) Or e.KeyChar = Chr(10) Then
            e.KeyChar = ""
        End If
    End Sub

    Private Sub txtSwipe_Enter(sender As Object, e As EventArgs) Handles txtSwipe.Enter
        ResetSwipeTimer()
    End Sub

    Private Sub BlockControls(Optional ByVal En As Boolean = True)
        cmdSwipe.Enabled = En
        cmdCancel.Enabled = En
        cmdSwipe.Enabled = En
        cmdTax.Enabled = En
    End Sub

    Private Sub SwipeDone(Optional ByVal ThrowAway As Boolean = False)
        On Error Resume Next
        ResetSwipeTimer(True)

        BlockControls(True)
        shpSwipe.FillColor = Color.Green

        txtQuantity.Select()
        If ThrowAway Then
            txtSwipe.Text = ""
            Exit Sub
        End If
        Quantity = GetPrice(txtQuantity.Text)        ' This is the 1st of 2 places quantity is set to non-zero.
        Total = GetPrice(txtPrice.Text)              ' This is the 1st of 2 places price is set.
        Cancelled = False
        Hide()                                    ' And unload the form to clean up.
    End Sub

    Private Sub txtSwipe_TextChanged(sender As Object, e As EventArgs) Handles txtSwipe.TextChanged
        On Error Resume Next
        If txtSwipe.Text = "" Then Exit Sub
        ResetSwipeTimer()   ' start the timer
        If Microsoft.VisualBasic.Left(txtSwipe.Text, 1) <> "%" Then
            txtSwipe.Text = ""
            txtPrice.Select()
            shpSwipe.FillColor = Color.Green
            BlockControls(True)
        End If
    End Sub

    Private Sub txtPrice_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrice.KeyPress
        On Error Resume Next
        'If txtSwipe.Visible Then
        '    If KeyAscii = Asc("%") Then
        '        txtSwipe = "%"
        '        txtSwipe.SelStart = 1
        '        KeyAscii = 0
        '        txtSwipe.SetFocus
        '        shpSwipe.FillColor = vbRed
        '        BlockControls False
        '        tmrLocate.Enabled = False
        '    End If
        'End If

        If txtSwipe.Visible Then
            If e.KeyChar = "%" Then
                txtSwipe.Text = "%"
                txtSwipe.SelectionStart = 1
                e.KeyChar = ""
                txtSwipe.Select()
                shpSwipe.FillColor = Color.Red
                BlockControls(False)
                tmrLocate.Enabled = False
            End If
        End If
    End Sub

    Private Sub txtSwipe_DoubleClick(sender As Object, e As EventArgs) Handles txtSwipe.DoubleClick
        If txtSwipe.Left = 0 Then
            'txtSwipe.Move cmdApply.Left, cmdApply.Top, cmdApply.Width
            txtSwipe.Location = New Point(cmdApply.Left, cmdApply.Top)
            txtSwipe.Size = New Size(cmdApply.Width, txtSwipe.Height)
        Else
            'txtSwipe.Move 0, 150, ScaleWidth
            txtSwipe.Location = New Point(0, 15)
            txtSwipe.Size = New Size(Me.ClientSize.Width, txtSwipe.Height)
        End If
    End Sub

    Private Sub txtPrice_Enter(sender As Object, e As EventArgs) Handles txtPrice.Enter
        SelectContents(txtPrice)
    End Sub

    Private Sub txtQuantity_Enter(sender As Object, e As EventArgs) Handles txtQuantity.Enter
        SelectContents(txtQuantity)
    End Sub

    Private Sub txtQuantity_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuantity.KeyPress
        'If KeyAscii = 10 Or KeyAscii = 13 Then KeyAscii = 0
        If e.KeyChar = Chr(10) Or e.KeyChar = Chr(13) Then e.KeyChar = ""
    End Sub

    Private Sub txtQuantity_Leave(sender As Object, e As EventArgs) Handles txtQuantity.Leave
        If Not ValidateQuantity(txtQuantity.Text, True) Then
            txtQuantity.Select()                ' Return focus to the quantity field.
        End If
    End Sub

    Private Sub txtPrice_Leave(sender As Object, e As EventArgs) Handles txtPrice.Leave
        On Error Resume Next
        If Not ValidatePrice(txtPrice.Text, True) Then
            txtPrice.Select()
        Else
            txtPrice.Text = CurrencyFormat(GetPrice(txtPrice.Text))
        End If
        '  If lblSwipe.Visible Then txtPrice.SetFocus
    End Sub

    Private Sub cmdSwipe_Click(sender As Object, e As EventArgs) Handles cmdSwipe.Click
        Quantity = GetPrice(txtQuantity.Text)        ' This is the 2nd of 3 places quantity is set to non-zero.
        Total = GetPrice(txtPrice.Text)              ' This is the 2nd of 3 places price is set.
        Cancelled = False
        Hide()
    End Sub

    Public Function DeveloperEx() As String
        Dim T As String
        T = "None"
        DeveloperEx = T
    End Function

End Class