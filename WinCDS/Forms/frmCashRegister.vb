Public Class frmCashRegister
    Dim SaleItems() As clsSaleItem
    Dim Processing As Boolean
    Dim RunningTotal As Decimal
    Dim ReturnMode As Boolean
    Dim ReceiptPrinted As Boolean
    Dim NonTaxable As Boolean
    Dim TaxableAmt As Decimal
    Dim SaleComplete As Boolean
    Dim SaleNo As String
    Private TapeScaleWidth As Integer
    Private TapePageLength As Integer
    Dim WithEvents MC As MailCheck
    Dim GotCust As Boolean
    Private Const QtyCol As Integer = 800
    Private Const ItemCol As Integer = 1000
    Private Const PriceCol As Integer = 3500
    Private Const DYMO_QtyCol As Integer = 500
    Private Const DYMO_ItemCol As Integer = 700
    Private Const DYMO_PriceCol As Integer = 2900

    Public ReadOnly Property MailZip() As String
        Get
            MailZip = ""
            If MailIndex <> 0 Then
                Dim M As clsMailRec
                M = New clsMailRec
                If M.Load(frmCashRegisterAddress.MailIndex, "#Index") Then
                    MailZip = M.Zip
                End If
                DisposeDA(M)
            End If
        End Get
    End Property

    Public ReadOnly Property MailIndex() As Integer
        Get
            MailIndex = Val(lblCust.Tag)
        End Get
    End Property

    Public Sub BeginSale()
        Const GD As Boolean = True
        Dim F As Integer
        ' Start a new sale.
        ' Prepare the form, set focus to the SKU entry box.
        'frmCashRegister.HelpContextID = 42500

        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        cmdComm.Tag = "1"
        'cmdComm.Value = True
        cmdComm_Click(cmdComm, New EventArgs)
        cboSalesList.Tag = ""               ' this will force the commissions person not to be carried over from sale to sale

        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        lblTotal.Text = "0.00"
        lblTax.Text = "0.00"
        lblTendered.Text = "0.00"
        lblDue.Text = "0.00"
        txtSku.Text = ""                    ' Clear any lingering SKUs.
        RunningTotal = 0                    ' Clear any lingering total..
        TaxableAmt = 0                      ' Clear any lingering taxes..
        fraSaleButtons.Visible = True       ' Show the sale buttons.
        SetReturnMode(False)                 ' Enter normal scan mode.
        SetNonTaxable(False)                 ' Default to taxable.
        SaleComplete = False                ' Default to non-complete.
        ReceiptPrinted = False              ' Clear receipt-printed flag.
        ShowButtons(0)                       ' Show the charge/management buttons.
        Erase SaleItems                     ' Erase any item history we might remember.
        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        cmdPayment.Enabled = True
        cmdReturn.Enabled = True
        cmdDiscount.Enabled = False         ' Don't allow discounts until an item has been sold.
        cmdPrint.Enabled = False            ' Don't allow reprints until the sale is complete.
        cmdDone.Enabled = False             ' Don't allow sale completion with no sale.
        cmdCancelSale.Text = "Cancel Sale"
        'cmdCancelSale.ToolTipText = "Click to cancel the sale, discarding all purchase information."
        ToolTip1.SetToolTip(cmdCancelSale, "Click to cancel the sale, discarding all purchase information.")
        vsbReceipt.SmallChange = 1
        'vsbReceipt.LargeChange = picReceiptContainer.ScaleHeight / picReceipt.TextHeight("X")
        vsbReceipt.LargeChange = picReceiptContainer.ClientRectangle.Height / CreateGraphics.MeasureString("X", Me.Font).Height

        LoadStoreLogo(imgLogo, StoresSld, True)  ' Load the store logo.
        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        'picReceipt.Cls
        picReceipt.Image = Nothing
        'PrintReceiptHeader(picReceipt)       ' Print the receipt header.

        'MoveReceipt(picReceipt.CurrentY)
        'MoveReceipt(picReceipt.Location.Y)
        Show()                             ' Show the form.
        On Error Resume Next
        'SetFocus
        txtSku.Select()                     ' And give focus to the SKU entry box.
        On Error GoTo 0

        ' If there's no printer set up, get one.
        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1

        CashRegisterPrinterSelector.SetSelectedPrinter(CashRegisterPrinter)
        If CashRegisterPrinterSelector.GetSelectedPrinter Is Nothing Then
            imgLogo.Visible = False
            CashRegisterPrinterSelector.Visible = True
            chkSavePrinter.Visible = True
        End If

        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        'SetCustomer(0)
        GotCust = False
    End Sub

    Private Sub SetCustomer(ByVal Index As Integer)
        If Index = 0 Then
            lblCust.Text = ""
            lblCust.Tag = ""
            RefreshReceipt()
            Exit Sub
        End If

        Dim M As clsMailRec
        M = New clsMailRec
        If M.Load(frmCashRegisterAddress.MailIndex, "#Index") Then
            lblCust = M.First & " " & M.Last & "  " & DressAni(M.Tele)
            lblCust.Tag = Index
        Else
            lblCust.Text = ""
            lblCust.Tag = ""
        End If
        DisposeDA(M)

        RefreshReceipt()
    End Sub

    Private Sub RefreshReceipt()
        If SaleComplete Then Exit Sub
        'picReceipt.Cls
        picReceipt.Image = Nothing
        PrintReceiptHeader(picReceipt)       ' Print the receipt header.
        Dim I As Integer
        On Error GoTo ErrOut
        For I = LBound(SaleItems) To UBound(SaleItems)
            PrintReceiptLine(Printer, SaleItems(I).Quantity, SaleItems(I).Desc, SaleItems(I).Style, SaleItems(I).DisplayPrice)
        Next
Done:
        Exit Sub
ErrOut:
        Resume Done
    End Sub

    Private Function PrintReceiptLine(ByVal Dest As Object, ByVal Qty As Double, ByVal Desc As String, ByVal Item As String, ByVal Price As Decimal) As Boolean
        Dim Q As Integer, I As Integer, P As Integer

        If IsDymoPrinter(Dest) Then
            Q = DYMO_QtyCol
            I = DYMO_ItemCol
            P = DYMO_PriceCol
        Else
            Q = QtyCol
            I = ItemCol
            P = PriceCol
        End If

        Dest.FontSize = 10 '8
        If Trim(Desc) = "" Then Desc = "No description available"

        Do While Dest.TextWidth(Desc) > Dest.ScaleWidth - 100
            Desc = Microsoft.VisualBasic.Left(Desc, Len(Desc) - 1)
        Loop
        If Dest.CurrentY > Dest.ScaleHeight - 2 * Dest.TextHeight("X") Then
            On Error Resume Next
            Dest.Height = Dest.CurrentY + 3 * Dest.TextHeight("X")
            On Error GoTo 0
        End If

        If Item = "PAYMENT" Or Item = "CHANGE" Or Item = "SALES TAX" Or Item = "SUBTOTAL" Then
            PrintToPosition(Dest, Desc, I, VBRUN.AlignConstants.vbAlignLeft, False)
        ElseIf Item = "--- Adj ---" Then
            PrintToPosition(Dest, Desc, 50, VBRUN.AlignConstants.vbAlignLeft, True)
            Exit Function
        Else
            PrintToPosition(Dest, Desc, 50, VBRUN.AlignConstants.vbAlignLeft, True)
            If Item <> "DISCOUNT" Then
                PrintToPosition(Dest, CStr(Qty), Q, VBRUN.AlignConstants.vbAlignRight, False)
            End If
            PrintToPosition(Dest, Item, I, VBRUN.AlignConstants.vbAlignLeft, False)
        End If
        PrintToPosition(Dest, CurrencyFormat(Price), P, VBRUN.AlignConstants.vbAlignRight, True)
    End Function

    Private Sub MoveReceipt(ByVal TargetY As Integer)
        'If picReceipt.CurrentY <= picReceiptContainer.ScaleHeight Then
        If picReceipt.Location.Y <= picReceiptContainer.ClientRectangle.Height Then
            ' If the container is taller than the receipt, align with the container.
            'picReceipt.Move 0, 0, picReceiptContainer.ScaleWidth, picReceiptContainer.ScaleHeight
            picReceipt.Location = New Point(0, 0)
            picReceipt.Size = New Size(picReceiptContainer.ClientRectangle.Width, picReceiptContainer.ClientRectangle.Height)
            vsbReceipt.Visible = False
        Else
            ' Receipt is taller than the container, make sure TargetY is visible.
            ' If TargetY is on the last page, put that part at the bottom..
            If TargetY > picReceipt.ClientRectangle.Height - picReceiptContainer.ClientRectangle.Height Then
                TargetY = picReceipt.ClientRectangle.Height - picReceiptContainer.ClientRectangle.Height
            End If
            'picReceipt.Move 0, -TargetY
            picReceipt.Location = New Point(0, -TargetY)
            vsbReceipt.Visible = True

            Dim LineCount As Integer
            'LineCount = picReceipt.ScaleHeight / picReceipt.TextHeight("X")                      ' Number of text lines in the receipt.
            LineCount = picReceipt.ClientRectangle.Height / CreateGraphics.MeasureString("X", Me.Font).Height                      ' Number of text lines in the receipt.
            'LineCount = LineCount - picReceiptContainer.ScaleHeight / picReceipt.TextHeight("X") ' Minus number of visible lines..
            LineCount = LineCount - picReceiptContainer.ClientRectangle.Height / CreateGraphics.MeasureString("X", Me.Font).Height ' Minus number of visible lines..

            If vsbReceipt.Value > LineCount And LineCount > 0 Then vsbReceipt.Value = LineCount  ' In case the receipt shrinks?
            vsbReceipt.Maximum = LineCount ' Max max of 32767!
            'If CInt(TargetY / picReceipt.TextHeight("X")) > LineCount Then
            If CInt(TargetY / CreateGraphics.MeasureString("X", Me.Font).Height) > LineCount Then
                vsbReceipt.Value = LineCount
            Else
                vsbReceipt.Value = CInt(TargetY / CreateGraphics.MeasureString("X", Me.Font).Height)
            End If
        End If
    End Sub

    Private Function PrintReceiptHeader(ByVal Dest As Object) As Boolean
        Dim LogoW As Integer, LogoH As Integer

        Dest.CurrentX = 0
        Dest.CurrentY = 0
        Dest.Font.Name = "Arial"

        Dim PrintReceiptLogo As Boolean, PrintReceiptAddress As Boolean
        ' Logo on Tape removed: 20170105..  Has gone back and forth
        'PrintReceiptLogo = (imgLogo.Picture <> 0) And False   ' Logos don't print well.
        PrintReceiptLogo = (imgLogo.Image IsNot Nothing) And False   ' Logos don't print well.
        PrintReceiptAddress = Not PrintReceiptLogo

        ' this is used b/c jerry likes to demonstrate receipt printing on the regular printer..
        ' we manually feed these values to it
        ' but for dymo printers w/ continuous tape, they are smaller than the
        ' standard receipt printers, so...
        On Error Resume Next
        If IsDymoPrinter(Dest) Then

            'BFH20170201 Stanley's has been having trouble.. trying this
            If Not IsStanleys Then
                Dest.PaperSize = DYMO_PaperSize_ContinuousWide
                TapeScaleWidth = 2918 ' printer.scalewidth
            End If
            '    Dest.Height = TapePageLength
        Else
            TapeScaleWidth = 3540
        End If
        On Error GoTo 0


        If PrintReceiptLogo Then
            LogoW = TapeScaleWidth
            LogoH = 10000 ' something too large
            MaintainPictureRatio(imgLogo, LogoW, LogoH, False)
            Dest.PaintPicture(imgLogo.Image, 0, Dest.CurrentY, LogoW, LogoH)
            Dest.CurrentY = Dest.CurrentY + LogoH + 250

        End If

        If PrintReceiptAddress Then
            Dest.Font.Bold = True
            'Dest.FontSize = BestFontFit(Dest, StoreSettings.Name, 14, Dest.ScaleWidth - 100, 500)
            'PrintToPosition Dest, StoreSettings.Name, (Dest.ScaleWidth - Dest.ScaleLeft) / 2, vbAlignTop, True  ' AlignTop really means Center here.

            Dest.FontSize = 14
            'PrintToPosition Dest, StoreSettings.Name, 400, vbAlignLeft, True
            'PrintInBox Dest, StoreSettings.Name, 400, Dest.CurrentY, Dest.ScaleWidth - 800, Dest.TextHeight("X"), , vbCenter
            PrintInBox(Dest, StoreSettings.Name, 400, Dest.CurrentY, TapeScaleWidth - 1200, Dest.TextHeight("X"), , VBRUN.AlignmentConstants.vbCenter)

            Dest.Font.Size = 10
            Dest.Font.Bold = False
            Dest.Print(vbCrLf)
            tPr("frmCashRegister.PrintReceiptHeader/PrintReceiptAddress")

            PrintToPosition(Dest, StoreSettings.Address, 300, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            PrintToPosition(Dest, StoreSettings.City, 300, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            PrintToPosition(Dest, StoreSettings.Phone, 300, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            Dest.Print(vbCrLf) : Tp()
            tPr()
        End If

        If MailIndex <> 0 Then
            Dim cMR As New clsMailRec
            tPr("frmCashRegister.PrintReceiptHeader/PrintCustomerAddress")
            cMR.Load(MailIndex, "#Index") : Tp()
            PrintToPosition(Dest, "Sold To:", DYMO_QtyCol - 200, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            PrintToPosition(Dest, cMR.First & " " & cMR.Last, DYMO_QtyCol, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            PrintToPosition(Dest, DressAni(cMR.Tele), DYMO_QtyCol, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            PrintToPosition(Dest, "", DYMO_QtyCol, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            DisposeDA(Nothing) : Tp()
            tPr()
        End If

        Dest.Font.Size = 10
        If IsDymoPrinter(Dest) Then
            tPr("frmCashRegister.PrintReceiptHeader/ColumnHeaders")
            PrintToPosition(Dest, "QTY", DYMO_QtyCol, VBRUN.AlignConstants.vbAlignRight, False) : Tp()
            PrintToPosition(Dest, "ITEM", DYMO_ItemCol, VBRUN.AlignConstants.vbAlignLeft, False) : Tp()
            PrintToPosition(Dest, "PRICE", DYMO_PriceCol, VBRUN.AlignConstants.vbAlignRight, True) : Tp()
            tPr()
        Else
            PrintToPosition(Dest, "QTY", QtyCol, VBRUN.AlignConstants.vbAlignRight, False)
            PrintToPosition(Dest, "ITEM", ItemCol, VBRUN.AlignConstants.vbAlignLeft, False)
            PrintToPosition(Dest, "PRICE", PriceCol, VBRUN.AlignConstants.vbAlignRight, True)
        End If
    End Function

    Private Sub SetReturnMode(ByVal RMode As Boolean)
        If RMode = False Then
            ReturnMode = False
            lblEnterStyle.Text = "Enter Style Number:"
            cmdReturn.Text = "Return"
            'cmdReturn.ToolTipText = "Click to scan an item to be returned."
            ToolTip1.SetToolTip(cmdReturn, "Click to scan an item to be returned.")
        Else
            ReturnMode = True
            lblEnterStyle.Text = "Scan Returned Item:"
            cmdReturn.Text = "Scan Item"
            'cmdReturn.ToolTipText = "Click to scan an item to be purchased."
            ToolTip1.SetToolTip(cmdReturn, "Click to scan an item to be purchased.")
        End If
    End Sub

    Private Sub SetNonTaxable(ByVal mVal As Boolean)
        NonTaxable = mVal
        If NonTaxable Then
            cmdTax.Width = 1100
            cmdTax.Left = 595
            cmdTax.Text = "No Tax:"
            lblTax.Text = "0.00"
            'cmdTax.ToolTipText = "Sale is nontaxable.  Click here to make it taxable."
            ToolTip1.SetToolTip(cmdTax, "Sale is nontaxable.  Click here to make it taxable.")
            AddSalesTax(True)
        Else
            cmdTax.Text = "Tax:"
            cmdTax.Width = 855
            cmdTax.Left = 840
            lblTax.Text = CurrencyFormat(GetStoreTax1() * TaxableAmt)
            'cmdTax.ToolTipText = "Sale is taxable.  Click here to make it nontaxable."
            ToolTip1.SetToolTip(cmdTax, "Sale is taxable.  Click here to make it nontaxable.")
        End If
        lblDue.Text = CurrencyFormat(RunningTotal + lblTax.Text - lblTendered.Text)
        On Error Resume Next
        txtSku.Select()
    End Sub

    Private Function AddSalesTax(Optional ByVal Negative As Boolean = False) As Boolean
        Dim Qpt As Decimal

        Qpt = QueryPrintedTax()
        If Qpt = GetPrice(lblTax.Text) Then Exit Function

        AddSubtotal()   ' Print subtotals before every tax line..

        Dim TaxLine As clsSaleItem
        TaxLine = New clsSaleItem
        TaxLine.Desc = "SALES TAX"
        TaxLine.NonTaxable = True
        TaxLine.Price = GetPrice(lblTax.Text) - Qpt   ' Report any remaining tax, or tax refund.
        TaxLine.DisplayPrice = TaxLine.Price
        TaxLine.Quantity = 0
        TaxLine.Style = "SALES TAX"
        AddSaleItem(TaxLine)
        DisposeDA(TaxLine)
    End Function

    Private Sub ShowButtons(ByVal Mode As Integer)
        Select Case Mode
            Case 0
                fraSaleButtons.Visible = True
                fraPaymentButtons.Visible = False
            Case 1
                fraPaymentButtons.Visible = True
                fraSaleButtons.Visible = False
            Case Else
                ' Error, do nothing.
        End Select
    End Sub

    Private Function QueryPrintedTax() As Decimal
        Dim I As Integer

        QueryPrintedTax = 0
        'If IsEmpty(SaleItems) Then Exit Function
        If IsNothing(SaleItems) Then Exit Function
        If UBound(SaleItems) = -1 Then Exit Function
        For I = UBound(SaleItems) To LBound(SaleItems) Step -1
            Select Case SaleItems(I).Style
                Case "SALES TAX"
                    QueryPrintedTax = QueryPrintedTax + SaleItems(I).Price
                Case Else
            End Select
        Next
    End Function

    Private Function AddSubtotal() As Boolean
        Dim SubTotal As clsSaleItem

        SubTotal = New clsSaleItem
        SubTotal.Desc = "SUBTOTAL"
        SubTotal.NonTaxable = True
        SubTotal.Price = QuerySubtotal()
        SubTotal.DisplayPrice = SubTotal.Price
        SubTotal.Quantity = 0
        SubTotal.Style = "SUBTOTAL"
        AddSaleItem(SubTotal)
        DisposeDA(SubTotal)
    End Function

    Public Function QuerySubtotal() As Decimal
        Dim I As Integer

        QuerySubtotal = 0
        'If IsEmpty(SaleItems) Then Exit Function
        If IsNothing(SaleItems) Then Exit Function
        If UBound(SaleItems) = -1 Then Exit Function
        For I = UBound(SaleItems) To LBound(SaleItems) Step -1
            Select Case SaleItems(I).Style
                Case "SUBTOTAL"
        ' Subtotals aren't real money, discounts are adjusted into item price.
                Case "DISCOUNT"
                    QuerySubtotal = QuerySubtotal + SaleItems(I).Price
                Case "PAYMENT"
                    QuerySubtotal = QuerySubtotal - SaleItems(I).Price
                Case Else
                    QuerySubtotal = QuerySubtotal + SaleItems(I).Price
            End Select
        Next
    End Function

    Private Sub AddSaleItem(ByRef Itm As clsSaleItem)
        If SaleComplete Then Exit Sub  ' Can't add to complete sales.

        If IsNothing(SaleItems) Then
            ReDim SaleItems(0)
        ElseIf UBound(SaleItems) >= 0 Then
            ReDim Preserve SaleItems(UBound(SaleItems) - LBound(SaleItems) + 1)
        Else
            ReDim SaleItems(0)
        End If

        SaleItems(UBound(SaleItems)) = Itm
        If Itm.Style <> "PAYMENT" And Itm.Style <> "DISCOUNT" And Itm.Style <> "CHANGE" And Itm.Style <> "SUBTOTAL" Then
            cmdDiscount.Enabled = True
        Else
            cmdDiscount.Enabled = False
        End If

        ' Payment Price is >0 but treated as <0.
        If Itm.Style = "PAYMENT" Then
            lblTendered.Text = CurrencyFormat(lblTendered.Text + Itm.Price)
            'OpenCashDrawer
        ElseIf Itm.Style = "CHANGE" Then
            ' Take change off the amount due?
            lblTendered.Text = CurrencyFormat(lblTendered.Text + Itm.Price)
            'OpenCashDrawer
        ElseIf Itm.Style = "SUBTOTAL" Or Itm.Style = "SALES TAX" Then
            ' Do nothing.
        ElseIf Itm.Style = "DISCOUNT" Then
            RunningTotal = CurrencyFormat(RunningTotal + Itm.DisplayPrice)   ' Strip off partial cents.
            If Not Itm.NonTaxable Then TaxableAmt = TaxableAmt + Itm.DisplayPrice
        Else
            Itm.Price = Itm.Quantity * Itm.Price
            Itm.DisplayPrice = Itm.Price
            RunningTotal = CurrencyFormat(RunningTotal + Itm.Price)   ' Strip off partial cents.
            If Not Itm.NonTaxable Then TaxableAmt = TaxableAmt + Itm.Price
        End If

        PrintReceiptLine(picReceipt, Itm.Quantity, Itm.Desc, Itm.Style, Itm.DisplayPrice)
        MoveReceipt(picReceipt.Location.Y)

        lblTotal.Text = CurrencyFormat(RunningTotal)
        If NonTaxable Then
            lblTax.Text = "0.00"
        Else
            ' This doesn't handle nontaxable items.
            lblTax.Text = CurrencyFormat(GetStoreTax1() * TaxableAmt)
        End If
        lblDue.Text = CurrencyFormat(RunningTotal + lblTax.Text - lblTendered.Text)

        cmdDone.Enabled = (lblDue.Text <= 0)
    End Sub

    Private Sub cmdComm_Click(sender As Object, e As EventArgs) Handles cmdComm.Click
        Select Case cmdComm.Tag
            Case ""   ' "", 3840,720,375,375, "&C"
                LoadSalesStaff
                cboSalesList.Visible = True
                txtSku.Visible = False
                cmdComm.Text = "&Select"
                'cmdComm.Default = True
                Me.AcceptButton = cmdComm
                'cmdComm.Move 1800, 720, 855, 375
                cmdComm.Location = New Point(180, 72)
                cmdComm.Size = New Size(85, 37)
                lblEnterStyle.Visible = False
                cmdComm.Tag = "1"
            Case "1"  ' "1", 1800, 720, 855, 375, "&Select"
                cboSalesList.Visible = False
                txtSku.Visible = True
                'cmdComm.Caption = "&C"
                cmdComm.Text = "&C"
                'cmdComm.Default = False
                Me.AcceptButton = Nothing
                'cmdComm.Move 3840, 720, 375, 375
                cmdComm.Location = New Point(384, 72)
                cmdComm.Size = New Size(37, 37)
                lblEnterStyle.Visible = True
                cmdComm.Tag = ""
        End Select
    End Sub

    Private Sub LoadSalesStaff()
        Dim Sm As Object, EE As Integer
        If cboSalesList.Tag <> "" Then Exit Sub
        Sm = GetSalesmanDatabase(StoresSld, True)

        cboSalesList.Items.Clear()

        For EE = LBound(Sm, 1) To UBound(Sm, 1)
            'cboSalesList.AddItem Sm(EE, 1), EE  ' - 1
            'cboSalesList.itemData(cboSalesList.NewIndex) = Sm(EE, 2)
            cboSalesList.Items.Insert(EE, New ItemDataClass(Sm(EE, 1), Sm(EE, 2)))
        Next
        'cboSalesList.AddItem "NO COMMISSION", 0
        'cboSalesList.itemData(cboSalesList.NewIndex) = -1
        'cboSalesList.ListIndex = 0
        cboSalesList.Items.Insert(0, New ItemDataClass("NO COMMISSION", -1))
        cboSalesList.SelectedIndex = 0
        cboSalesList.Tag = "."
    End Sub

    Private Sub frmCashRegister_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox "frmCashReg: ->Form_load"
        'frmCashRegister.HelpContextID = 42500
        ' bfh20051202 - i want to use fracust, but Jerry doesn't like that way..
        ' it still holds the data tho
        fraCust.Visible = False

        'cmdDev.Visible = IsDevelopment()
        'MsgBox "frmCashReg: Form_load->"

    End Sub

    Private Sub cmdTax_Click(sender As Object, e As EventArgs) Handles cmdTax.Click
        Dim g As Graphics = picReceipt.CreateGraphics
        Dim x As Integer = 20
        Dim y As Integer = 30
        g.DrawString("Hello.", New Font("Arial", 12), Brushes.Black, x, y)
    End Sub

    Private Sub cmdPayment_Click(sender As Object, e As EventArgs) Handles cmdPayment.Click
        Dim g As Graphics = picReceipt.CreateGraphics
        Dim x As Integer = 20
        Dim y As Integer = 30
        g.DrawString("Hello.", New Font("Arial", 12), Brushes.Black, x, y)
    End Sub

    Private Sub picReceipt_Paint(sender As Object, e As PaintEventArgs) Handles picReceipt.Paint
        Dim StringToDraw As String = "Hi there!! :-)"
        Dim MyBrush As New SolidBrush(Color.Black)
        Dim StringFont As New Font("Arial", 20)
        Dim PixelsAcross As Integer = 20
        Dim PixelsDown As Integer = 30
        'e.Graphics.DrawString(StringToDraw, StringFont, MyBrush, PixelsAcross, PixelsDown)
        '----------------------------------------------------------------------------------

        Dim LogoW As Integer, LogoH As Integer

        'Dest.CurrentX = 0
        'Dest.CurrentY = 0
        'Dest.Font.Name = "Arial"


        Dim PrintReceiptLogo As Boolean, PrintReceiptAddress As Boolean
        ' Logo on Tape removed: 20170105..  Has gone back and forth
        'PrintReceiptLogo = (imgLogo.Picture <> 0) And False   ' Logos don't print well.
        PrintReceiptLogo = (imgLogo.Image IsNot Nothing) And False   ' Logos don't print well.
        PrintReceiptAddress = Not PrintReceiptLogo

        ' this is used b/c jerry likes to demonstrate receipt printing on the regular printer..
        ' we manually feed these values to it
        ' but for dymo printers w/ continuous tape, they are smaller than the
        ' standard receipt printers, so...
        'On Error Resume Next
        'If IsDymoPrinter(Dest) Then

        '    'BFH20170201 Stanley's has been having trouble.. trying this
        '    If Not IsStanleys Then
        '        Dest.PaperSize = DYMO_PaperSize_ContinuousWide
        '        TapeScaleWidth = 2918 ' printer.scalewidth
        '    End If
        '    '    Dest.Height = TapePageLength
        'Else
        '    TapeScaleWidth = 3540
        'End If
        On Error GoTo 0

        If PrintReceiptLogo Then
            LogoW = TapeScaleWidth
            LogoH = 10000 ' something too large
            MaintainPictureRatio(imgLogo, LogoW, LogoH, False)
            'Dest.PaintPicture(imgLogo.Image, 0, Dest.CurrentY, LogoW, LogoH)
            'Dest.CurrentY = Dest.CurrentY + LogoH + 250

        End If

        If PrintReceiptAddress Then
            'Dest.Font.Bold = True
            'Dest.FontSize = BestFontFit(Dest, StoreSettings.Name, 14, Dest.ScaleWidth - 100, 500)
            'PrintToPosition Dest, StoreSettings.Name, (Dest.ScaleWidth - Dest.ScaleLeft) / 2, vbAlignTop, True  ' AlignTop really means Center here.

            'Dest.FontSize = 14
            'PrintToPosition Dest, StoreSettings.Name, 400, vbAlignLeft, True
            'PrintInBox Dest, StoreSettings.Name, 400, Dest.CurrentY, Dest.ScaleWidth - 800, Dest.TextHeight("X"), , vbCenter
            'PrintInBox(Dest, StoreSettings.Name, 400, Dest.CurrentY, TapeScaleWidth - 1200, Dest.TextHeight("X"), , VBRUN.AlignmentConstants.vbCenter)

            '<PrintInBox>
            Dim PrintText As String
            PrintText = StoreSettings.Name
            If PrintText <> "" Then
                'If FontSize = -1 Then FontSize = 300
                'PrintOb.FontSize = BestFontFit(PrintOb, PrintText, FontSize, Width, Height)

                'Select Case VAlign
                ''Case vbAlignTop
                '    Case AlignConstants.vbAlignTop
                '        PrintOb.CurrentY = Top
                ''Case vbAlignBottom
                '    Case AlignConstants.vbAlignBottom
                '        PrintOb.CurrentY = Top + Height - PrintOb.TextHeight(PrintText)
                '    Case Else ' center
                '        PrintOb.CurrentY = Top + (Height - PrintOb.TextHeight(PrintText)) / 2
                'End Select

                'Dim El As Object
                'For Each El In Split(PrintText, vbCrLf)
                '    Select Case HAlign
                '    'Case vbAlignRight, 1
                '        Case AlignConstants.vbAlignRight, 1
                '            PrintOb.CurrentX = Left + Width - PrintOb.TextWidth(El)
                '    'Case vbAlignLeft, 0
                '        Case AlignConstants.vbAlignLeft, 0
                '            PrintOb.CurrentX = Left
                '        Case Else 'center
                '            PrintOb.CurrentX = Left + (Width - PrintOb.TextWidth(El)) / 2
                '    End Select
                '    PrintOb.Print(El)
                'PrintInBox(Dest, StoreSettings.Name, 400, Dest.CurrentY, TapeScaleWidth - 1200, Dest.TextHeight("X"), , VBRUN.AlignmentConstants.vbCenter)
                e.Graphics.DrawString(PrintText, StringFont, MyBrush, 20, 30)
                'Next
                '    PrintOb.Print PrintText
            End If

            'If BorderStyle <> 0 Then
            '    PrintOb.Line(Left, Top - Left + Width, Top, BorderStyle)
            '    PrintOb.Line(Left, Top - Left, Top + Height, BorderStyle)
            '    PrintOb.Line(Left + Width, Top - Left + Width, Top + Height, BorderStyle)
            '    PrintOb.Line(Left, Top + Height - Left + Width, Top + Height, BorderStyle)
            'End If
            '</PrintInBox>

            'Dest.Font.Size = 10
            'Dest.Font.Bold = False
            'Dest.Print(vbCrLf)
            tPr("frmCashRegister.PrintReceiptHeader/PrintReceiptAddress")

            'PrintToPosition(Dest, StoreSettings.Address, 300, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            'PrintToPosition(Dest, StoreSettings.City, 300, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            'PrintToPosition(Dest, StoreSettings.Phone, 300, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            'Dest.Print(vbCrLf) : Tp()
            tPr()
        End If

        'If MailIndex <> 0 Then
        '    Dim cMR As New clsMailRec
        '    tPr("frmCashRegister.PrintReceiptHeader/PrintCustomerAddress")
        '    cMR.Load(MailIndex, "#Index") : Tp()
        '    PrintToPosition(Dest, "Sold To:", DYMO_QtyCol - 200, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
        '    PrintToPosition(Dest, cMR.First & " " & cMR.Last, DYMO_QtyCol, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
        '    PrintToPosition(Dest, DressAni(cMR.Tele), DYMO_QtyCol, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
        '    PrintToPosition(Dest, "", DYMO_QtyCol, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
        '    DisposeDA(Nothing) : Tp()
        '    tPr()
        'End If

        'Dest.Font.Size = 10
        'If IsDymoPrinter(Dest) Then
        '    tPr("frmCashRegister.PrintReceiptHeader/ColumnHeaders")
        '    PrintToPosition(Dest, "QTY", DYMO_QtyCol, VBRUN.AlignConstants.vbAlignRight, False) : Tp()
        '    PrintToPosition(Dest, "ITEM", DYMO_ItemCol, VBRUN.AlignConstants.vbAlignLeft, False) : Tp()
        '    PrintToPosition(Dest, "PRICE", DYMO_PriceCol, VBRUN.AlignConstants.vbAlignRight, True) : Tp()
        '    tPr()
        'Else
        '    PrintToPosition(Dest, "QTY", QtyCol, VBRUN.AlignConstants.vbAlignRight, False)
        '    PrintToPosition(Dest, "ITEM", ItemCol, VBRUN.AlignConstants.vbAlignLeft, False)
        '    PrintToPosition(Dest, "PRICE", PriceCol, VBRUN.AlignConstants.vbAlignRight, True)
        'End If
    End Sub
End Class