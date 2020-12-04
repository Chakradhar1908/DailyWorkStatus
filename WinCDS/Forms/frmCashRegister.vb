Imports System.ComponentModel
Imports System.Drawing
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
    'Private Const QtyCol As Integer = 800
    Private Const QtyCol As Integer = 20
    'Private Const ItemCol As Integer = 1000
    Private Const ItemCol As Integer = 55
    'Private Const PriceCol As Integer = 3500
    Private Const PriceCol As Integer = 210
    'Private Const DYMO_QtyCol As Integer = 500
    Private Const DYMO_QtyCol As Integer = 10
    'Private Const DYMO_ItemCol As Integer = 700
    Private Const DYMO_ItemCol As Integer = 30
    'Private Const DYMO_PriceCol As Integer = 2900
    Private Const DYMO_PriceCol As Integer = 200
    Dim MailCheckFormLoaded As Boolean
    Dim MultiRows As Boolean
    Dim TopValue(0) As Integer, TopValue2(0) As Integer
    Dim Y As Integer, YY As Integer

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
        On Error GoTo 0
        CashRegisterPrinterSelector.SetSelectedPrinter(CashRegisterPrinter)
        If CashRegisterPrinterSelector.GetSelectedPrinter Is Nothing Then
            imgLogo.Visible = False
            CashRegisterPrinterSelector.Visible = True
            chkSavePrinter.Visible = True
        End If
        SetCustomer(0)
        Show()                             ' Show the form.
        On Error Resume Next
        'SetFocus
        txtSku.Select()                     ' And give focus to the SKU entry box.
        'On Error GoTo 0

        ' If there's no printer set up, get one.
        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        'CashRegisterPrinterSelector.SetSelectedPrinter(CashRegisterPrinter) --------> NOTE: Moved this line to above Show() method. Reason: Code exeuction sequence is different from vb6.0 code. Continuing it here will give wrong output.
        'If CashRegisterPrinterSelector.GetSelectedPrinter Is Nothing Then
        '    imgLogo.Visible = False
        '    CashRegisterPrinterSelector.Visible = True
        '    chkSavePrinter.Visible = True
        'End If
        'If gD Then MsgBox "frmCashReg:  " & F: F = F + 1
        'SetCustomer(0) --------> NOTE: Moved this line to above Show() method. Reason: Code exeuction sequence is different from vb6.0 code. Continuing it here will give wrong output.
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
            lblCust.Text = M.First & " " & M.Last & "  " & DressAni(M.Tele)
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

        '<CT>
        'Note: Commented the below code and replacd it with Application.DoEvents(), cause this code will be directly written in Paint event.
        'PrintReceiptHeader(picReceipt)       ' Print the receipt header.
        '</CT>
        'Dim I As Integer
        'On Error GoTo ErrOut
        'For I = LBound(SaleItems) To UBound(SaleItems)
        '    PrintReceiptLine(Printer, SaleItems(I).Quantity, SaleItems(I).Desc, SaleItems(I).Style, SaleItems(I).DisplayPrice)
        'Next
        'Application.DoEvents() 'This line will redirects to picReceipt picturebox's Paint event.
        'Done:
        '        Exit Sub
        'ErrOut:
        '        Resume Done
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
            cmdTax.Width = 77
            cmdTax.Left = 40
            cmdTax.Text = "No Tax:"
            lblTax.Text = "0.00"
            'cmdTax.ToolTipText = "Sale is nontaxable.  Click here to make it taxable."
            ToolTip1.SetToolTip(cmdTax, "Sale is nontaxable.  Click here to make it taxable.")
            AddSalesTax(True)
        Else
            cmdTax.Text = "Tax:"
            cmdTax.Width = 50
            cmdTax.Left = 63
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

        'Application.DoEvents()

        'PrintReceiptLine(picReceipt, Itm.Quantity, Itm.Desc, Itm.Style, Itm.DisplayPrice)
        'MoveReceipt(picReceipt.Location.Y)
        MultiRows = True
        picReceipt_Paint(New Object, New PaintEventArgs(picReceipt.CreateGraphics, New Rectangle))
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
                LoadSalesStaff()
                cboSalesList.Visible = True
                txtSku.Visible = False
                cmdComm.Text = "&Select"
                'cmdComm.Default = True
                Me.AcceptButton = cmdComm
                'cmdComm.Move 1800, 720, 855, 375
                cmdComm.Location = New Point(146, 45)
                cmdComm.Size = New Size(60, 26)
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
                cmdComm.Location = New Point(269, 45)
                cmdComm.Size = New Size(39, 26)
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
            'cboSalesList.Items.Insert(EE, New ItemDataClass(Sm(EE, 1), Sm(EE, 2)))
            cboSalesList.Items.Insert(EE, New ItemDataClass(Sm(EE, 0), Sm(EE, 1)))
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
        cmdDev.Visible = IsDevelopment()
        'MsgBox "frmCashReg: Form_load->"
        'TopValue(0) = 220
        'TopValue2(0) = 235

        'Dim b As Bitmap = New Bitmap(600, 600)
        'Dim g As Graphics = Graphics.FromImage(b)
        'picReceipt.BackgroundImage = b
        'picReceipt.Size = b.Size
        PictureboxScroll()
    End Sub

    Private Sub cmdTax_Click(sender As Object, e As EventArgs) Handles cmdTax.Click
        If SaleComplete Then Exit Sub
        SetNonTaxable(Not NonTaxable)
    End Sub

    Private Sub cmdPayment_Click(sender As Object, e As EventArgs) Handles cmdPayment.Click
        If SaleComplete Then Exit Sub
        ' Hide charge/management frame, show payment frame.
        ShowButtons(1)
        txtSku.Select()
    End Sub

    Private Sub picReceipt_Paint(sender As Object, e As PaintEventArgs) Handles picReceipt.Paint
        'Dim StringToDraw As String = "Hi there!! :-)"
        Dim MyBrush As New SolidBrush(Color.Black)
        'Dim StringFont As New Font("Arial", 20)
        'Dim PixelsAcross As Integer = 20
        'Dim PixelsDown As Integer = 30
        'e.Graphics.DrawString(StringToDraw, StringFont, MyBrush, PixelsAcross, PixelsDown)

        Dim b As Bitmap = New Bitmap(600, 800)
        'Dim g As Graphics = Graphics.FromImage(b)
        e.Graphics.FillRectangle(Brushes.White, New Rectangle(0, 0, b.Width, b.Height))
        'picReceipt.BackgroundImage = b
        'picReceipt.Size = b.Size
        'pnlPicReceipt.Size = New Size(200, 200)
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
            'LogoH = 10000 ' something too large
            LogoH = 1000 ' something too large
            MaintainPictureRatio(imgLogo, LogoW, LogoH, False)
            'Dest.PaintPicture(imgLogo.Image, 0, Dest.CurrentY, LogoW, LogoH)
            'Note:The below five lines are replacement for the above commented line. (Dest.PaintPicture)
            picReceipt.Image = imgLogo.Image
            picReceipt.Left = 0
            picReceipt.Top = picReceipt.Top
            picReceipt.Width = 354
            picReceipt.Height = LogoH
            'Dest.CurrentY = Dest.CurrentY + LogoH + 250
            picReceipt.Top = picReceipt.Top + LogoH + 25
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
                'e.Graphics.DrawString(PrintText, StringFont, MyBrush, 20, 30)
                'e.Graphics.DrawString(PrintText, New Font("Arial", 14, FontStyle.Bold), MyBrush, 20, 30)
                'e.Graphics.DrawString(PrintText, New Font("Arial", 14, FontStyle.Bold), MyBrush, picReceipt.Width / 4, 5)
                e.Graphics.DrawString(PrintText, New Font("Arial", 13, FontStyle.Bold), MyBrush, 25, 5)
                'Next
                '   PrintOb.Print PrintText
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
            e.Graphics.DrawString(StoreSettings.Address, New Font("Arial", 10), MyBrush, 20, 50) : Tp()
            'PrintToPosition(Dest, StoreSettings.City, 300, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            e.Graphics.DrawString(StoreSettings.City, New Font("Arial", 10), MyBrush, 20, 65) : Tp()
            'PrintToPosition(Dest, StoreSettings.Phone, 300, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            e.Graphics.DrawString(StoreSettings.Phone, New Font("Arial", 10), MyBrush, 20, 80) : Tp()
            'Dest.Print(vbCrLf) : Tp()
            tPr()
        End If

        If MailIndex <> 0 Then
            Dim cMR As New clsMailRec
            tPr("frmCashRegister.PrintReceiptHeader/PrintCustomerAddress")
            cMR.Load(MailIndex, "#Index") : Tp()
            'PrintToPosition(Dest, "Sold To:", DYMO_QtyCol - 200, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            e.Graphics.DrawString("Sold To:", New Font("Arial", 10), MyBrush, DYMO_QtyCol - 5, 120) : Tp()
            'PrintToPosition(Dest, cMR.First & " " & cMR.Last, DYMO_QtyCol, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            e.Graphics.DrawString(cMR.First & " " & cMR.Last, New Font("Arial", 10), MyBrush, DYMO_QtyCol, 140) : Tp()
            'PrintToPosition(Dest, DressAni(cMR.Tele), DYMO_QtyCol, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            e.Graphics.DrawString(DressAni(cMR.Tele), New Font("Arial", 10), MyBrush, DYMO_QtyCol, 160) : Tp()
            'PrintToPosition(Dest, "", DYMO_QtyCol, VBRUN.AlignConstants.vbAlignLeft, True) : Tp()
            e.Graphics.DrawString("", New Font("Arial", 10), MyBrush, DYMO_QtyCol, 180) : Tp()
            DisposeDA(Nothing) : Tp()
            tPr()
        End If

        'Dest.Font.Size = 10
        'If IsDymoPrinter(Dest) Then
        If IsDymoPrinter(picReceipt) Then
            tPr("frmCashRegister.PrintReceiptHeader/ColumnHeaders")
            'PrintToPosition(Dest, "QTY", DYMO_QtyCol, VBRUN.AlignConstants.vbAlignRight, False) : Tp()
            e.Graphics.DrawString("QTY", New Font("Arial", 10), MyBrush, DYMO_QtyCol, 200) : Tp()
            'PrintToPosition(Dest, "ITEM", DYMO_ItemCol, VBRUN.AlignConstants.vbAlignLeft, False) : Tp()
            e.Graphics.DrawString("ITEM", New Font("Arial", 10), MyBrush, DYMO_ItemCol, 200) : Tp()
            'PrintToPosition(Dest, "PRICE", DYMO_PriceCol, VBRUN.AlignConstants.vbAlignRight, True) : Tp()
            e.Graphics.DrawString("PRICE", New Font("Arial", 10), MyBrush, DYMO_PriceCol, 200) : Tp()
            tPr()
        Else
            If MailIndex = 0 Then
                'PrintToPosition(Dest, "QTY", QtyCol, VBRUN.AlignConstants.vbAlignRight, False)
                e.Graphics.DrawString("QTY", New Font("Arial", 10), MyBrush, QtyCol, 140) : Tp()
                'PrintToPosition(Dest, "ITEM", ItemCol, VBRUN.AlignConstants.vbAlignLeft, False)
                e.Graphics.DrawString("ITEM", New Font("Arial", 10), MyBrush, ItemCol, 140) : Tp()
                'PrintToPosition(Dest, "PRICE", PriceCol, VBRUN.AlignConstants.vbAlignRight, True)
                e.Graphics.DrawString("PRICE", New Font("Arial", 10), MyBrush, PriceCol, 140) : Tp()
            Else
                'PrintToPosition(Dest, "QTY", QtyCol, VBRUN.AlignConstants.vbAlignRight, False)
                e.Graphics.DrawString("QTY", New Font("Arial", 10), MyBrush, QtyCol, 200) : Tp()
                'PrintToPosition(Dest, "ITEM", ItemCol, VBRUN.AlignConstants.vbAlignLeft, False)
                e.Graphics.DrawString("ITEM", New Font("Arial", 10), MyBrush, ItemCol, 200) : Tp()
                'PrintToPosition(Dest, "PRICE", PriceCol, VBRUN.AlignConstants.vbAlignRight, True)
                e.Graphics.DrawString("PRICE", New Font("Arial", 10), MyBrush, PriceCol, 200) : Tp()
            End If
        End If

        'If MultiRows = False Then Exit Sub
        Y = 220
        YY = 240
        'Note: The below code is from Private Sub RefreshReceipt()
        Dim I As Integer
        On Error GoTo ErrOut
        For I = LBound(SaleItems) To UBound(SaleItems)
            'PrintReceiptLine(Printer, SaleItems(I).Quantity, SaleItems(I).Desc, SaleItems(I).Style, SaleItems(I).DisplayPrice)
            'Private Function PrintReceiptLine(ByVal Dest As Object, ByVal Qty As Double, ByVal Desc As String, ByVal Item As String, ByVal Price As Decimal) As Boolean
            'Note: The below code is from Private Function PrintReceiptLine
            Dim Q As Integer, Ic As Integer, P As Integer, Desc As String, Item As String
            Dim Z As Integer, Zincremented As Boolean
            Dim S As StringFormat = New StringFormat
            S.FormatFlags = StringFormatFlags.DirectionRightToLeft

            'If IsDymoPrinter(Dest) Then
            If IsDymoPrinter(Printer) Then
                Q = DYMO_QtyCol
                Ic = DYMO_ItemCol
                P = DYMO_PriceCol
            Else
                Q = QtyCol
                Ic = ItemCol
                P = PriceCol
            End If

            'Dest.FontSize = 10 '8
            'If Trim(Desc) = "" Then Desc = "No description available"
            If Trim(SaleItems(I).Desc) = "" Then
                Desc = "No description available"
            Else
                Desc = Trim(SaleItems(I).Desc)
            End If

            'Do While Dest.TextWidth(Desc) > Dest.ScaleWidth - 100
            '    Desc = Microsoft.VisualBasic.Left(Desc, Len(Desc) - 1)
            'Loop
            'If Dest.CurrentY > Dest.ScaleHeight - 2 * Dest.TextHeight("X") Then
            '    On Error Resume Next
            '    Dest.Height = Dest.CurrentY + 3 * Dest.TextHeight("X")
            '    On Error GoTo 0
            'End If
            'Dim TopValue(0) As Integer, Y As Integer
            'Dim TopValue2(0) As Integer
            'TopValue(0) = 220

            'Y = TopValue(0)
            'YY = TopValue2(0)
            Item = SaleItems(I).Style
            If Item = "PAYMENT" Or Item = "CHANGE" Or Item = "SALES TAX" Or Item = "SUBTOTAL" Then
                'PrintToPosition(Dest, Desc, I, VBRUN.AlignConstants.vbAlignLeft, False)
                'e.Graphics.DrawString(Desc, New Font("Arial", 10), MyBrush, Ic, 220)
                'e.Graphics.DrawString(Desc, New Font("Arial", 10), MyBrush, Ic, Y)
                e.Graphics.DrawString(Desc, New Font("Arial", 10), MyBrush, Ic, Z)
                'Z = Z + 20
                Zincremented = True
            ElseIf Item = "--- Adj ---" Then
                'PrintToPosition(Dest, Desc, 50, VBRUN.AlignConstants.vbAlignLeft, True)
                'e.Graphics.DrawString(Desc, New Font("Arial", 10), MyBrush, 50, 220)
                e.Graphics.DrawString(Desc, New Font("Arial", 10), MyBrush, 5, Y)
                'Exit Function
                Exit Sub
            Else
                'PrintToPosition(Dest, Desc, 50, VBRUN.AlignConstants.vbAlignLeft, True)
                'e.Graphics.DrawString(Desc, New Font("Arial", 10), MyBrush, 50, 220)
                e.Graphics.DrawString(Desc, New Font("Arial", 10), MyBrush, 5, Y)
                If Item <> "DISCOUNT" Then
                    'PrintToPosition(Dest, CStr(Qty), Q, VBRUN.AlignConstants.vbAlignRight, False)
                    'e.Graphics.DrawString(SaleItems(I).Quantity, New Font("Arial", 10), MyBrush, Q, 235)
                    e.Graphics.DrawString(SaleItems(I).Quantity, New Font("Arial", 10), MyBrush, Q + 5, YY)
                End If
                'PrintToPosition(Dest, Item, I, VBRUN.AlignConstants.vbAlignLeft, False)
                'e.Graphics.DrawString(Item, New Font("Arial", 10), MyBrush, Ic, 235)
                e.Graphics.DrawString(Item, New Font("Arial", 10), MyBrush, Ic, YY)
            End If
            'PrintToPosition(Dest, CurrencyFormat(Price), P, VBRUN.AlignConstants.vbAlignRight, True)
            'e.Graphics.DrawString(CurrencyFormat(SaleItems(I).DisplayPrice), New Font("Arial", 10), MyBrush, P, 235)
            'e.Graphics.DrawString(CurrencyFormat(SaleItems(I).DisplayPrice), New Font("Arial", 10), MyBrush, P, YY)
            If Item = "PAYMENT" Or Item = "CHANGE" Or Item = "SALES TAX" Or Item = "SUBTOTAL" Then
                'e.Graphics.DrawString(CurrencyFormat(SaleItems(I).DisplayPrice), New Font("Arial", 10), MyBrush, P, Y)

                'Dim S As StringFormat = New StringFormat
                'S.FormatFlags = StringFormatFlags.DirectionRightToLeft
                e.Graphics.DrawString(CurrencyFormat(SaleItems(I).DisplayPrice), New Font("Arial", 10), MyBrush, P, Z, S)
                'Z = Z + 20
            Else
                e.Graphics.DrawString(CurrencyFormat(SaleItems(I).DisplayPrice), New Font("Arial", 10), MyBrush, P, YY)
            End If

            'End of PrintReceiptLine
            'Y = TopValue(0)
            'TopValue(0) = Y + 20
            Y = Y + 40
            If Zincremented = False Then
                Z = Y
            Else
                Z = Z + 18
            End If
            'YY = TopValue2(0)
            'TopValue2(0) = YY + 15
            YY = YY + 40
        Next
Done:
        'MultiRows = False
        Exit Sub
ErrOut:
        Resume Done

        'PrintReceiptLine(picReceipt, Itm.Quantity, Itm.Desc, Itm.Style, Itm.DisplayPrice)
        'Private Function PrintReceiptLine(ByVal Dest As Object, ByVal Qty As Double, ByVal Desc As String, ByVal Item As String, ByVal Price As Decimal) As Boolean
        Dim Q2 As Integer, I2 As Integer, P2 As Integer

        If IsDymoPrinter(picReceipt) Then
            Q2 = DYMO_QtyCol
            I2 = DYMO_ItemCol
            P2 = DYMO_PriceCol
        Else
            Q2 = QtyCol
            I2 = ItemCol
            P2 = PriceCol
        End If

        Dim Itm As clsSaleItem
        Dim Descc As String, Item2 As String, Qty2 As Double, Price2 As Decimal
        'Dest.FontSize = 10 '8
        'If Trim(Desc) = "" Then Desc = "No description available"
        If Trim(Itm.Desc) = "" Then
            Descc = "No description available"
        Else
            Descc = Trim(Itm.Desc)
        End If

        'Do While Dest.TextWidth(Desc) > Dest.ScaleWidth - 100
        '    Desc = Microsoft.VisualBasic.Left(Desc, Len(Desc) - 1)
        'Loop
        'If Dest.CurrentY > Dest.ScaleHeight - 2 * Dest.TextHeight("X") Then
        '    On Error Resume Next
        '    Dest.Height = Dest.CurrentY + 3 * Dest.TextHeight("X")
        '    On Error GoTo 0
        'End If

        Item2 = Itm.Style
        Qty2 = Itm.Quantity
        Price2 = Itm.DisplayPrice
        If Item2 = "PAYMENT" Or Item2 = "CHANGE" Or Item2 = "SALES TAX" Or Item2 = "SUBTOTAL" Then
            'PrintToPosition(Dest, Desc, I, VBRUN.AlignConstants.vbAlignLeft, False)
            e.Graphics.DrawString(Descc, New Font("Arial", 10), MyBrush, I2, 255)
        ElseIf Item2 = "--- Adj ---" Then
            'PrintToPosition(Dest, Desc, 50, VBRUN.AlignConstants.vbAlignLeft, True)
            e.Graphics.DrawString(Descc, New Font("Arial", 10), MyBrush, I2, 255)
            'Exit Function
            Exit Sub
        Else
            'PrintToPosition(Dest, Desc, 50, VBRUN.AlignConstants.vbAlignLeft, True)
            e.Graphics.DrawString(Descc, New Font("Arial", 10), MyBrush, I2, 255)
            If Item2 <> "DISCOUNT" Then
                'PrintToPosition(Dest, CStr(Qty), Q, VBRUN.AlignConstants.vbAlignRight, False)
                e.Graphics.DrawString(CStr(Qty2), New Font("Arial", 10), MyBrush, Q2, 270)
            End If
            'PrintToPosition(Dest, Item, I, VBRUN.AlignConstants.vbAlignLeft, False)
            e.Graphics.DrawString(Item2, New Font("Arial", 10), MyBrush, I2, 270)
        End If
        'PrintToPosition(Dest, CurrencyFormat(Price), P, VBRUN.AlignConstants.vbAlignRight, True)
        e.Graphics.DrawString(CurrencyFormat(Price2), New Font("Arial", 10), MyBrush, P2, 270)
    End Sub

    Private Sub cmdFND_Click(sender As Object, e As EventArgs) Handles cmdFND.Click
        cmdFND.Visible = False
        ProcessSku(True)
    End Sub

    Private Function ProcessSku(Optional ByVal Fnd As Boolean = False) As Boolean
        ' Sole processing function to add txtSku to the sale.
        If Processing Then Exit Function            ' This function can't run while the subform is up.
        If ReceiptPrinted Then Exit Function        ' Can't add to printed receipt.
        If txtSku.Text = "" Then Exit Function      ' Ignore stray enter if input empty

        Processing = True

        ProcessSku = AddSkuToSale(txtSku.Text, Fnd)
        If ProcessSku Then txtSku.Text = ""
        FocusSelect(txtSku)

        Processing = False
    End Function

    Private Function AddSkuToSale(ByVal SKU As String, Optional ByVal Fnd As Boolean = False) As Boolean
        ' Look up the Sku.  If it's not in the database, fail.
        Dim InvData As CInvRec
        Dim Found As Boolean

        If SaleComplete Then Exit Function

        InvData = New CInvRec
        Found = InvData.Load(SKU, "Style")
        If Found Or Fnd Then
            ' Item exists, prompt for quantity.
            Dim IQP As clsSaleItem  ' Class object.. we should be able to store this in an array to prepare for commit.
            If Found Then
                IQP = frmCashRegisterQuantity.GetQuantityAndPrice(InvData.Style, InvData.Desc, InvData.OnSale, NonTaxable)
                IQP.Status = "ST"
            Else
                Dim FNDPrice As Decimal, FNDVendor As String, FNDDesc As String
                If Not frmCashRegisterFND.GetInformation(SKU, FNDPrice, FNDVendor, FNDDesc) Then
                    DisposeDA(InvData)
                    Exit Function
                End If
                IQP = frmCashRegisterQuantity.GetQuantityAndPrice(SKU, FNDDesc, FNDPrice, NonTaxable)

                IQP.VendorNo = FormatVendorNo(Microsoft.VisualBasic.Left(FNDVendor, 3))
                IQP.Vendor = Trim(Mid(FNDVendor, 4))
                IQP.Status = "FND"
            End If
            DisposeDA(InvData)
            If IQP Is Nothing Then Exit Function

            ' Also need to get price, in case price was changed..
            If IQP.Quantity > 0 Then
                If ReturnMode Then IQP.Quantity = -IQP.Quantity
                SetReturnMode(False)

                AddSaleItem(IQP)   ' Add the item to the sale array, and the receipt.
                AddSkuToSale = True

                ' Beep?  Some barcode scanners have built-in sounds for completed scans.
                ' In any case, the cashier needs audio feedback or they'll be constantly
                ' watching the screen instead of scanning more items..

            Else
                AddSkuToSale = False
            End If
        Else
            ' Item not found.
            'BFH20170104 - Adding FND support..
            MessageBox.Show("This item could not be found in the database!" & vbCrLf2 & "To enter an Item Not in Inventory, click FND.", "Item Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    txtSku = ""
            cmdFND.Visible = True
            FocusControl(txtSku)
            AddSkuToSale = False
        End If
    End Function

    Private Sub cmdDev_Click(sender As Object, e As EventArgs) Handles cmdDev.Click
        Dim S As String
        S = SelectOption("Select DEV MODE Function", frmSelectOption.ESelOpts.SelOpt_List + frmSelectOption.ESelOpts.SelOpt_ToItem, "Test Print Receipt")

        Select Case S
            Case "Test Print Receipt" : TestPrintReceipt()
        End Select
    End Sub

    Private Function TestPrintReceipt() As Boolean
        Dim vSaleNo As String
        vSaleNo = InputBox("Sale No:", "Enter Value", "10349")
        If vSaleNo = "" Then Exit Function

        SaleNo = vSaleNo

        Dim RS As ADODB.Recordset, IQP As clsSaleItem
        RS = GetRecordsetBySQL("SELECT * FROM [GrossMargin] WHERE SaleNo='" & SaleNo & "' ORDER BY [MarginLine]", , GetDatabaseAtLocation)

        Do While Not RS.EOF
            IQP = New clsSaleItem
            IQP.Style = Trim(IfNullThenNilString(RS("Style").Value))
            IQP.Status = Trim(IfNullThenNilString(RS("Status").Value))
            IQP.Price = IfNullThenZeroCurrency(RS("SellPrice").Value)
            IQP.DisplayPrice = IfNullThenZeroCurrency(RS("SellPrice").Value)
            IQP.Quantity = IfNullThenZero(RS("Quantity").Value)
            IQP.Location = IfNullThenZero(RS("Location").Value)
            IQP.Desc = IfNullThenNilString(RS("Desc").Value)
            IQP.Vendor = IfNullThenNilString(RS("Vendor").Value)
            IQP.VendorNo = IfNullThenNilString(RS("VendorNo").Value)

            AddSaleItem(IQP)
            IQP = Nothing

            RS.MoveNext()
        Loop

        ReceiptPrinted = True
        SaleComplete = True
        'Processed = True

        cmdMainMenu.Enabled = True
        cmdPrint.Enabled = True   ' Allow reprints after the first printing run.
        cmdCancelSale.Text = "Next Sale"
        'cmdCancelSale.ToolTipText = "Click to begin a new sale.  This data has been saved."
        ToolTip1.SetToolTip(cmdCancelSale, "Click to begin a new sale.  This data has been saved.")
        cmdCancelSale.Enabled = True
        cmdDone.Enabled = False

        PrintReceipt()
        TestPrintReceipt = True
    End Function

    Public Sub PrintReceipt(Optional ByVal CCSign As Boolean = False)
        On Error GoTo PrintReceiptError
        If Not SaleComplete Then
            MessageBox.Show("You can't print a receipt until the sale is completed.", "Sale Not Complete")
            Exit Sub
        End If

        If CashRegisterPrinterSelector.GetSelectedPrinter Is Nothing Then
            If Not SetDymoPrinter() Then
                MessageBox.Show("You must select a printer before printing receipts!", "No Printer Selected")
                Exit Sub
            End If
        End If

        Dim OldPrinter As String
        OldPrinter = Printer.DeviceName
        SetPrinter(CashRegisterPrinter)
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor

        ' We'll have to loop through the items and print them as text.
        ' The receipt printer just can't handle graphics.

        ' It's not a runtime error if the printer is offline!  How odd.
        PrintReceiptHeader(Printer)
        Dim I As Integer
        For I = LBound(SaleItems) To UBound(SaleItems)
            PrintReceiptLine(Printer, SaleItems(I).Quantity, SaleItems(I).Desc, SaleItems(I).Style, SaleItems(I).DisplayPrice)
        Next
        PrintReceiptTrailer(Printer, CCSign)

        '  Printer.ScaleHeight = Printer.CurrentY
        Printer.EndDoc()
        ReceiptPrinted = True

        SetPrinter(OldPrinter)
        'MousePointer = vbNormal
        Me.Cursor = Cursors.Default
        txtSku.Select()
        Exit Sub

PrintReceiptError:
        Dim N As Integer, S As String
        N = Err.Number
        S = Err.Description
        Select Case N
            Case 482 : ErrNoPrinter()
            Case Else
                MessageBox.Show("Error printing receipt [" & N & "]" & vbCrLf & S, "Error Printing Receipt")
        End Select
        ReceiptPrinted = True

        SetPrinter(OldPrinter)
        'MousePointer = vbNormal
        Me.Cursor = Cursors.Default
        txtSku.Select()
    End Sub

    Private Function PrintReceiptTrailer(ByRef Dest As Object, Optional ByVal CCSign As Boolean = False) As Boolean
        Dim I As Integer
        Dest.FontSize = 10 '8
        If Dest.CurrentY > Dest.ScaleHeight - 1 * Dest.TextHeight("X") Then
            On Error Resume Next ' Sometimes,it objects tothe 'height' being set..  Probably basedd on paper type..
            Dest.Height = Dest.CurrentY + 1 * Dest.TextHeight("X")
            Err.Clear()
            On Error GoTo 0
        End If

        ' Print time and sale number.
        PrintToPosition(Dest, Format(Now, DateFormatString() & " hh:mm") & " " & SaleNo, 0, VBRUN.AlignConstants.vbAlignLeft, True)

        ' Print special message..
        Dest.Print("")

        If CCSign Then
            Dest.Print("I authorize the above transaction.")
            Dest.Print("")
            Dest.Print("")
            Dest.Print("X ________________________")
            Dest.Print("            STORE COPY")
        Else
            MainMenu.rtbn.File = CashRegisterMessageFile()
            MainMenu.rtbn.FileRead(True)
            If MainMenu.rtbn.RichTextBox.Text <> "" Then Dest.Print(MainMenu.rtbn.RichTextBox.Text)
        End If

        ' special handling for DYMO receipt printer
        If Not IsDymoPrinter(Dest) Then

            'bfh20050317 - needed 5 blank lines at the bottom of the receipt!!
            For I = 1 To 6
                Dest.Print(" ")
            Next

            If IsRobys() Then
            Else
                'bfh20050317 - attempt to send <ESC>d0 to signal the receipt printer to cut the paper..
                Dest.Print(Chr(27) & "d0")
            End If
        Else
            For I = 1 To 5
                Dest.Print(" ")
            Next
            Dest.Print(".")
        End If

        If TypeName(Dest) = "PictureBox" Then
            TapePageLength = Dest.CurrentY
        End If
    End Function

    Private Function AddCashJournal(ByVal Amount As Decimal, ByVal TransType As Integer) As Boolean
        AddNewCashJournalRecord(TransType, Amount, SaleNo, "CASH REGISTER", DateFormat(Today))
    End Function

    Private Function SelectLastItem() As clsSaleItem
        Dim I As Integer
        If IsNothing(SaleItems) Then Exit Function
        If UBound(SaleItems) = -1 Then Exit Function
        For I = UBound(SaleItems) To LBound(SaleItems) Step -1
            Select Case SaleItems(I).Style
                Case "PAYMENT", "DISCOUNT", "CHANGE", "SUBTOTAL", "SALES TAX"
                Case Else
                    SelectLastItem = SaleItems(I)
                    Exit Function
            End Select
        Next
        SelectLastItem = Nothing
    End Function

    Private Sub chkSavePrinter_CheckedChanged(sender As Object, e As EventArgs) Handles chkSavePrinter.CheckedChanged
        If chkSavePrinter.Checked = True Then
            If CashRegisterPrinterSelector.GetSelectedPrinter Is Nothing Then
                MessageBox.Show("You must select a printer first.", "WinCDS")
            Else
                'CashRegisterPrinter = CashRegisterPrinterSelector.GetSelectedPrinter.DeviceName
                CashRegisterPrinter(, CashRegisterPrinterSelector.GetSelectedPrinter.DeviceName)
                CashRegisterPrinterSelector.Visible = False
                chkSavePrinter.Visible = False
                chkSavePrinter.Checked = False   ' Only save once.
                imgLogo.Visible = True
                txtSku.Select()
            End If
        End If
    End Sub

    Private Function SalesCode() As String
        Dim X As Integer
        If cboSalesList.SelectedIndex < 0 Then Exit Function
        'X = cboSalesList.itemData(cboSalesList.ListIndex)
        X = CType(cboSalesList.Items(cboSalesList.SelectedIndex), ItemDataClass).ItemData
        'bfh20051201 - changed from "" to "99" for default sales code to reflect normal sales
        If X = -1 Then SalesCode = "99" Else SalesCode = Format(X, "00")
    End Function

    Private Sub cmdDiscount_Click(sender As Object, e As EventArgs) Handles cmdDiscount.Click
        ' Change the price of the last item entered.. by percent or amount.
        ' Log the change as part of the same Margin record, so reports still work..

        If SaleComplete Then Exit Sub

        Dim Dsc As clsSaleItem, PrevItem As clsSaleItem
        PrevItem = SelectLastItem()
        If PrevItem Is Nothing Then
            MessageBox.Show("You must add an item to the sale before applying a discount.", "WinCDS")
            Exit Sub
        End If

        If Not CheckAccess("Give Discounts", True, False, True) Then
            ' Not authorized to give discounts.
            txtSku.Select()
            Exit Sub
        End If

        Dsc = frmCashRegisterQuantity.GetQuantityAndPrice("DISCOUNT", "", 0, True)
        If Dsc Is Nothing Then Exit Sub
        If Dsc.Price > 0 Then
            ' Discounts always have 0 price, but may display a negative price.
            ' That negative price is added to the discounted item's actual price, but not display price.
            Select Case Dsc.Quantity
                Case 2
                    Dsc.DisplayPrice = -Format(Dsc.Price, "0.00")
                    Dsc.Price = Dsc.DisplayPrice
                    Dsc.Desc = "Flat Discount"
                Case 1
                    Dsc.DisplayPrice = -Format(QuerySubtotal() * Dsc.Price / 100, "0.00")
                    Dsc.Desc = Format(Dsc.Price, "0") & "% Sale Discount"
                    Dsc.Price = Dsc.DisplayPrice
                Case Else
                    Dsc.DisplayPrice = -Format(PrevItem.Price * Dsc.Price / 100, "0.00")
                    Dsc.Desc = Format(Dsc.Price, "0") & "% Discount"
                    Dsc.Price = 0
                    PrevItem.Price = PrevItem.Price + Dsc.DisplayPrice
            End Select
            Dsc.NonTaxable = PrevItem.NonTaxable        ' Retain tax status from previous item.
            ' Adjust previous sale item..
            AddSaleItem(Dsc)              ' This needs to add it to the receipt.
        End If
        'cmdDiscount.Enabled = False

        txtSku.Select()
    End Sub

    Private Sub cmdPayCash_Click(sender As Object, e As EventArgs) Handles cmdPayCash.Click
        If SaleComplete Then ShowButtons(0) : Exit Sub
        AddPayment(1)
        ShowButtons(0)
        On Error Resume Next
        txtSku.Select()
    End Sub

    Private Function AddPayment(ByVal PmtType As Integer) As Boolean
        Dim A As Decimal, Pmt As clsSaleItem
        Dim X As New clsSaleItem

        If SaleComplete Then Exit Function

        A = GetPrice(lblDue.Text)
        If A < 0 Then
            If IsIn("" & PmtType, "3", "4", "5", "6") Then
                Pmt = frmCashRegisterQuantity.DoReturn(-A, "3", "", "")
            Else
                Pmt = frmCashRegisterQuantity.DoReturn(-A, PmtType, "", "")
            End If
        Else
            Pmt = frmCashRegisterQuantity.GetQuantityAndPrice("PAYMENT", PmtType, A, True)
        End If

        If Pmt Is Nothing Then Exit Function

        Pmt.NonTaxable = True         ' Payments are never taxable, or they'd reduce tax.
        AddSalesTax()                   ' Print sales tax before the subtotal..
        AddSubtotal()                   ' Put a subtotal before every payment..
        AddSaleItem(Pmt)               ' This needs to add it to the receipt.
        If Pmt.Extra1 <> "" Then
            X.Style = "--- Adj ---"
            X.Desc = Pmt.Extra1
            AddSaleItem(X)
            X = Nothing
        End If
        If Pmt.Balance > 0 Then
            Dim R As Decimal
            R = Pmt.Balance
            Pmt.Clear()
            Pmt.Style = "NOTES"
            Pmt.Desc = "CARD BALANCE: " & CurrencyFormat(R)
            AddSaleItem(Pmt)
        End If

        DisposeDA(X)
        AddPayment = True
    End Function

    Private Sub vsbReceipt_ValueChanged(sender As Object, e As EventArgs) Handles vsbReceipt.ValueChanged
        'MoveReceipt(vsbReceipt.Value * picReceipt.TextHeight("X"))
    End Sub

    Private Function AddSalesJournal() As Boolean
        Dim W As Decimal, T As Decimal, Sm As String, nT As Decimal
        W = GetPrice(lblTotal.Text)
        T = GetPrice(lblTax.Text)
        Sm = Microsoft.VisualBasic.Left(GetLocalComputerName, 12)
        nT = RunningTotal - TaxableAmt ' non taxable
        AddNewAuditRecord(SaleNo, "NS CASH REGISTER", Today, W, T, 0, 0, 0, W, T, 1, Sm, nT)
        '  Dim Sj As SalesJournal
        '  With Sj
        '    .LeaseNo = SaleNo
        '    .Name1 = "CASH REGISTER"
        '    .TransDate = DateFormat(Date)
        '    .Written = lblTotal.Caption
        '    .TaxCharged1 = lblTax.Caption
        '' BFH20060331 - for delivered sales, ARCASHSALES needs to total out to 0.  Since this should be the only line, it must be zero here
        ''    .ARCASHSALES = Format(GetPrice(.Written) + GetPrice(.TaxCharged1), "$0.00")
        '    .ARCASHSALES = 0
        '    .Controll = "(" & .ARCASHSALES & ")"
        '    .UndSls = "$0.00"
        '    .DelSls = .Written
        '    .TaxRec1 = .TaxCharged1
        '    .TaxCode = "1"
        '    .Salesman = Left(GetLocalComputerName, 12)
        '  End With
        '  SalesJournal_SetStore StoresSld
        '  SalesJournal_AddRecord Sj
    End Function

    Private Sub GetCustomer()
        If Not StoreSettings.bUseCashRegisterAddress Then Exit Sub
        SetCustomer(0)

        GotCust = True
        MC = MailCheck
        'MC.optTelephone.Checked = True
        MailCheckSaleNoChecked = False
        MC.HidePriorSales = True
        'MC.Show vbModal
        MC.ShowDialog()
        'MC.Dispose()
        MC.HidePriorSales = False
        If Not MC Is Nothing Then
            'Unload MC
            MC.Close()
        End If

        MC = Nothing
    End Sub

    Private Sub MC_Cancelled(ByRef PreventUnload As Boolean, ByRef PreventMainMenu As Boolean) Handles MC.Cancelled
        'Unload MC
        MC.Close()
        SetCustomer(0)
        '  Unload Me
    End Sub

    Private Sub MC_CustomerFound(MailIndex As Integer, ByRef Cancel As Boolean) Handles MC.CustomerFound
        Cancel = True
        'Unload MC
        MC.Close()
        MC.Dispose()
        EditCustomer(MailIndex)
    End Sub

    Private Sub EditCustomer(ByVal Index As Integer)
        frmCashRegisterAddress.AddressType = 0
        frmCashRegisterAddress.MailIndex = Index
        'frmCashRegisterAddress.Show vbModal
        frmCashRegisterAddress.ShowDialog()
        SetCustomer(frmCashRegisterAddress.MailIndex)
    End Sub

    Private Sub MC_CustomerNotFound(ByRef Ignore As Boolean, ByRef DoUnload As Boolean) Handles MC.CustomerNotFound
        Ignore = True
        DoUnload = True
        'Unload MC
        MC.Close()
        EditCustomer(0)
    End Sub

    Private Sub PromptForStyle()
        'InvCkStyle.Show vbModal
        InvCkStyle.ShowDialog()
        If Not InvCkStyle.Canceled Then txtSku.Text = InvCkStyle.StyleCkIt
        'Unload InvCkStyle
        InvCkStyle.Close()
    End Sub

    Private Sub frmCashRegister_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If MailCheckFormLoaded = False Then
            If Not GotCust Then GetCustomer()
            MailCheckFormLoaded = True
        End If
    End Sub

    Private Sub cmdMainMenu_Click(sender As Object, e As EventArgs) Handles cmdMainMenu.Click
        Me.Close()
    End Sub

    Private Sub cmdPayCheck_Click(sender As Object, e As EventArgs) Handles cmdPayCheck.Click
        If SaleComplete Then ShowButtons(0) : Exit Sub
        AddPayment(2)
        ShowButtons(0)
        txtSku.Select()
    End Sub

    Private Sub cmdPayCredit_Click(sender As Object, e As EventArgs) Handles cmdPayCredit.Click
        If SaleComplete Then ShowButtons(0) : Exit Sub
        AddPayment(3)  ' Which credit type?  3/4/5/6=VISA/MC/DISC/AMEX
        ShowButtons(0)
        txtSku.Select()
    End Sub

    Private Sub cmdPayDebit_Click(sender As Object, e As EventArgs) Handles cmdPayDebit.Click
        If SaleComplete Then ShowButtons(0) : Exit Sub
        AddPayment(9)
        ShowButtons(0)
        txtSku.Select()
    End Sub

    Private Sub cmdPayStoreCard_Click(sender As Object, e As EventArgs) Handles cmdPayStoreCard.Click
        If SaleComplete Then ShowButtons(0) : Exit Sub
        AddPayment(12)
        ShowButtons(0)
        txtSku.Select()
    End Sub

    Private Sub cmdPayReturnToSale_Click(sender As Object, e As EventArgs) Handles cmdPayReturnToSale.Click
        ShowButtons(0)
        txtSku.Select()
    End Sub

    Private Sub cmdReturn_Click(sender As Object, e As EventArgs) Handles cmdReturn.Click
        If SaleComplete Then Exit Sub
        ' Make them scan returned item..
        SetReturnMode(Not ReturnMode)
        txtSku.Select()
    End Sub

    Private Sub cmdCancelSale_Click(sender As Object, e As EventArgs) Handles cmdCancelSale.Click
        ' Clear the sale..
        BeginSale()
        On Error Resume Next
        txtSku.Select()
        GetCustomer()
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        PrintReceipt()
    End Sub

    Private Sub cmdDone_Click(sender As Object, e As EventArgs) Handles cmdDone.Click
        Dim Commable As String, NeedsSignature As Boolean
        Dim Cst As Decimal, Frt As Decimal
        ' If the sale's not ready to be completed, give a warning.
        ' Or better yet, don't have this button enabled at all.
        ' A cash&carry sale is not ready if:
        '   No items
        '   Cash due
        '   Change due?
        If IsNothing(SaleItems) Then Exit Sub  ' Silent error, no items on sale.

        AddSalesTax() ' If tax hasn't been added, add it now.

        If Val(lblDue.Text) > 0 Then
            MessageBox.Show("There is money due.  Sales must be paid in full at time of purchase.", "WinCDS")
            Exit Sub
        End If

        If Val(lblDue.Text) < 0 Then
            ' Change is always given in cash.
            Dim Itm As clsSaleItem
            Itm = New clsSaleItem
            Itm.Desc = "CHANGE (CASH)"
            Itm.NonTaxable = True
            Itm.Price = GetPrice(lblDue.Text)
            Itm.DisplayPrice = Itm.Price
            Itm.Quantity = 1
            Itm.Style = "CHANGE"
            AddSaleItem(Itm)
            DisposeDA(Itm)
        End If

        Dim I As Integer, SaleName As String, SaleIndex As String
        Dim cMR As clsMailRec
        If MailIndex = 0 Then
            SaleIndex = "0"
            SaleName = "CASH REGISTER"
        Else
            SaleIndex = MailIndex
            cMR = New clsMailRec
            If cMR.Load(SaleIndex, "#Index") Then
                SaleName = cMR.Last
            Else
                SaleName = "CASH REGISTER [UNKNOWN]"  ' just in case
            End If
            DisposeDA(cMR)
        End If

        cmdDone.Enabled = False
        cmdCancelSale.Enabled = False
        cmdMainMenu.Enabled = False
        cmdPayment.Enabled = False
        cmdReturn.Enabled = False
        cmdDiscount.Enabled = False
        cmdPrint.Enabled = False

        ' Process the sale.  Prompt for next sale?
        ' Create holding record.
        ' The store really has to be in Auto-HoldingID mode.
        ' Add lines to GrossMargin.
        ' Decrement 2Data quantities.
        ' Add lines to Detail?
        ' Add lines to Cash+Sales journals.
        ' It would be very good to have a single function for "Add this item to this sale",
        ' with all detailed accounting taken care of within that.

        Dim Holding As cHolding
        Holding = New cHolding
        Holding.LeaseNo = GetLeaseNumber()                  ' Create a lease number.
        Holding.Deposit = GetPrice(lblTendered.Text)     ' Amount paid.
        Holding.Sale = GetPrice(lblTotal.Text) + GetPrice(lblTax.Text)      ' Total amount of sale, with tax
        Holding.NonTaxable = RunningTotal - TaxableAmt      ' Amount that's not taxable..
        Holding.LastPay = Today
        Holding.Salesman = SalesCode()                        ' Who's logged in?
        '  Commable = IIf(Holding.Salesman = "", "", "C")
        Holding.Status = "D"                                ' All Cash Register sales are delivered.
        Holding.Comm = "N"                                  ' Commission isn't paid
        Holding.Index = Val(SaleIndex)
        Holding.Save()
        SaleNo = Holding.LeaseNo

        For I = LBound(SaleItems) To UBound(SaleItems)
            Select Case SaleItems(I).Style
                Case "PAYMENT", "CHANGE"
                    ' Save as payment.
                    ' Deal with description...
                    SaveNewMarginRecord(Holding.LeaseNo, "PAYMENT", SaleItems(I).Desc & Space(5) & DateFormat(Today), SaleItems(I).Quantity, SaleItems(I).Price,
                      "", "", "", 0, 0, 0, "", "", "DEL", Holding.Salesman,
                      StoresSld, Today, Today, StoresSld, SaleName, Today, "",
                      SaleIndex, "100", 0, "", "", "", Nothing, SaleItems(I).TransID)
                    AddCashJournal(SaleItems(I).Price, SaleItems(I).Quantity)
                Case "DISCOUNT"
                    ' Save as a note, zero cost.
                    SaveNewMarginRecord(Holding.LeaseNo, "NOTES", "DISCOUNT (" & SaleItems(I).Price & ")", 0, SaleItems(I).Price,
                      "", "", "", 0, 0, 0, "", "", "DEL", Holding.Salesman,
                      StoresSld, Today, Today, StoresSld, SaleName, Today, "",
                      SaleIndex, "0", 0, "", "", "", Nothing, "")
                Case "SALES TAX"
                    SaveNewMarginRecord(Holding.LeaseNo, "TAX1", "SALES TAX", 1, SaleItems(I).Price,
                      "", "", "", 0, 0, 0, "", "", "DEL", Holding.Salesman,
                      0, Today, Today, StoresSld, SaleName, Today, "",
                      SaleIndex, "0", 0, "", "", "", Nothing, "")
                Case "SUBTOTAL"
                    SaveNewMarginRecord(Holding.LeaseNo, "SUB", "Sub Total =", 0, SaleItems(I).Price,
                      "", "", "", 0, 0, 0, "", "", "DEL", Holding.Salesman,
                      0, Today, Today, StoresSld, SaleName, Today, "",
                      SaleIndex, "0", 0, "", "", "", Nothing, "")
                Case "--- Adj ---"
                    SaveNewMarginRecord(Holding.LeaseNo, "--- Adj ---", SaleItems(I).Desc, 0, 0,
                      "", "", "", 0, 0, 0, "", "", "DEL", Holding.Salesman,
                      0, Today, Today, StoresSld, SaleName, Today, "",
                      SaleIndex, "0", 0, "", "", "", Nothing, "")
                    NeedsSignature = True
                Case Else
                    ' Actual item style..
                    Dim InvData As CInvRec, Found As Boolean
                    Dim Dpt As Integer, RN As Integer, GM As Double, ST As String
                    InvData = New CInvRec
                    Found = InvData.Load(SaleItems(I).Style, "Style")

                    If Found Or SaleItems(I).Status = "FND" Then
                        ' BFH20051219 ItemCost
                        '            Cst = InvData.Cost
                        If Found Then
                            Cst = GetItemCost(SaleItems(I).Style, StoresSld, , SaleItems(I).Quantity)
                            Frt = (InvData.Landed - InvData.Cost) * SaleItems(I).Quantity
                            Dpt = InvData.DeptNo
                            RN = InvData.RN
                            GM = CalculateGM(SaleItems(I).Price, InvData.Landed)
                            ST = "DELTW"
                        Else
                            Cst = 0
                            Frt = 0
                            Dpt = 0
                            RN = 0
                            GM = 100
                            ST = "DELFND"
                        End If

                        Dim Xmar As CGrossMargin, Xdet As CInventoryDetail
                        Xmar = SaveNewMarginRecord(Holding.LeaseNo, SaleItems(I).Style, SaleItems(I).Desc, SaleItems(I).Quantity, SaleItems(I).Price,
            SaleItems(I).Vendor, Dpt, SaleItems(I).VendorNo, Cst, Frt, RN, "", Commable, ST, Holding.Salesman,
            StoresSld, Today, Today, StoresSld, SaleName, Today, "",
            SaleIndex, GM, 0, "", "", "", Nothing, "")


                        If RN <> 0 Then ' No Detail on FND
                            ' These are handled by CreateDetailRecord.
                            '              UpdateQuarterlySales InvData, SaleItems(I).Quantity, Date
                            '              UpdateInventoryQuantity InvData, SaleItems(I).Quantity, StoresSld
                            Xdet = CreateDetailRecord(InvData, SaleNo, SaleName, SaleItems(I).Quantity, "DELTW", StoresSld, Today, 0, Xmar.MarginLine)
                            Xmar.Detail = Xdet.DetailID
                            Xmar.Save()

                            InvData.Save()
                        End If

                        DisposeDA(Xdet, Xmar)
                    Else
                        MessageBox.Show("Error saving item: " & SaleItems(I).Style & ".", "WinCDS")
                    End If
                    DisposeDA(InvData)
            End Select
        Next

        AddSalesJournal()            ' Save an audit record.

        DisposeDA(Holding)
        PrintReceiptTrailer(picReceipt) ' Add datestamp and any trailing notes..
        'MoveReceipt(picReceipt.CurrentY)
        MoveReceipt(picReceipt.Top)
        SaleComplete = True

        OpenCashDrawer()            ' Open the drawer to accept payment and make change.

        SalePackageUpdate(SaleNo)  ' update package-related fields.

        cmdMainMenu.Enabled = True
        cmdPrint.Enabled = True   ' Allow reprints after the first printing run.
        If SwipeCards() And NeedsSignature Then
            PrintReceipt(True)
        End If
        PrintReceipt(False)        ' And always print the sale right away.
        cmdCancelSale.Text = "Next Sale"
        'cmdCancelSale.ToolTipText = "Click to begin a new sale.  This data has been saved."
        ToolTip1.SetToolTip(cmdCancelSale, "Click to begin a new sale.  This data has been saved.")
        cmdCancelSale.Enabled = True
        cmdDone.Enabled = False
        txtSku.Select()
    End Sub

    Private Sub frmCashRegister_DoubleClick(sender As Object, e As EventArgs) Handles MyBase.DoubleClick
        If Not IsDevelopment() Then Exit Sub
        If fraSaleTotals.Tag = "" Then
            'ControlLoading(fraSaleTotals)
            fraSaleTotals.Tag = "1"
        Else
            'ControlLoadingRemove(fraSaleTotals)
            fraSaleTotals.Tag = ""
        End If
    End Sub

    Private Sub frmCashRegister_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'queryunload of vb6.0
        If Not cmdMainMenu.Enabled Then e.Cancel = True

        'form unload of vb6.0
        modProgramState.Order = ""
        MainMenu.Show()                        ' Always go back to the main menu after a sale..
    End Sub

    Private Sub imgLogo_DoubleClick(sender As Object, e As EventArgs) Handles imgLogo.DoubleClick
        imgLogo.Visible = False
        CashRegisterPrinterSelector.Visible = True
        chkSavePrinter.Visible = True
    End Sub

    Private Sub fraCust_Click(sender As Object, e As EventArgs) Handles fraCust.Click
        If cmdCancelSale.Text = "Cancel Sale" Then GetCustomer()
    End Sub

    Private Sub lblCust_Click(sender As Object, e As EventArgs) Handles lblCust.Click
        If cmdCancelSale.Text = "Cancel Sale" Then GetCustomer()
    End Sub

    Private Sub picReceipt_Enter(sender As Object, e As EventArgs)
        On Error Resume Next  ' Somehow this is possible when the price/qty form is loaded modal.
        txtSku.Select()
    End Sub

    Private Sub txtSku_DoubleClick(sender As Object, e As EventArgs) Handles txtSku.DoubleClick
        PromptForStyle()
    End Sub

    Private Sub txtSku_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSku.KeyPress
        'If KeyAscii = 13 Or KeyAscii = 9 Then
        '    ProcessSku()
        'Else
        '    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'End If

        If Asc(e.KeyChar) = 13 Or Asc(e.KeyChar) = 9 Then
            ProcessSku()
        Else
            e.KeyChar = UCase(e.KeyChar)
        End If
    End Sub

    Private Sub txtSku_Leave(sender As Object, e As EventArgs) Handles txtSku.Leave
        ' Some barcode scanners terminate a scan with VbTab.
        ' In that case, this event traps the scanned barcode.
        If SaleComplete Then Exit Sub     ' Do nothing if the sale is final.
        If Processing Then Exit Sub       ' Don't do this if doing something else
        If cmdFND.Visible Then Exit Sub   ' LostFocus would block Click
        ProcessSku()
    End Sub

    Private Sub PictureboxScroll()
        'pnlPicReceipt.SuspendLayout()
        'Me.SuspendLayout()
        pnlPicReceipt.AutoScroll = True
        'pnlPicReceipt.AutoScrollMinSize = New Size(600, 400)
        'pnlPicReceipt.AutoScrollMinSize = New Size(600, 800)
        pnlPicReceipt.AutoScrollMinSize = New Size(200, 800)
        pnlPicReceipt.Size = New Size(298, 384)

        picReceipt.Dock = DockStyle.Fill
        picReceipt.Location = New Point(0, 0)
        'picReceipt.Size = New Size(600, 400)
        picReceipt.Size = New Size(600, 800)
        picReceipt.TabStop = False
        'pnlPicReceipt.ResumeLayout(False)
        'Me.ResumeLayout(False)
    End Sub
End Class
