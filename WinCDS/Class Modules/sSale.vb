﻿Imports stdole
Imports VBRUN
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports System.Drawing
Imports System.Globalization
Public Class sSale
    Private mStore As Integer
    Private mSaleNo As String

    Public Tele As String
    Public Name As String
    Public MailIndex As Integer
    Public CashRegisterSale As Boolean

    Public CustType As Integer
    Public AdvertizingType As Integer
    Public TaxZone As Integer

    Public SaleDate As String
    Public DelDate As String
    Public PorD As String
    Public Status As String

    Public ItemCount As Integer
    Private Items() As clsSaleItem

    Public SalesCode As String
    Public SalesSplit As String

    Private mLab As Decimal
    Private mDel As Decimal
    Private mStain As Decimal

    Public StopStart As String
    Public StopEnd As String

    Private ProcessSalePOs As Collection

    Public Function LoadFromBillOSale() As Boolean
        Dim X As Integer, N As Integer
        Dim Y As Integer
        Dim Taxed As Boolean
        Dim TaxThisItem As Boolean
        '  If Not IsBillOSaleFree Then Exit Function

        Clear()

        With BillOSale
            Store = StoresSld
            SaleNo = .BillOfSale.Text
            MailIndex = .MailIndex
            Name = .CustomerLast.Text
            Tele = .CustomerPhone1.Text

            CustType = .cboCustType.SelectedIndex
            AdvertizingType = .cboAdvertisingType.SelectedIndex
            TaxZone = .cboTaxZone.SelectedIndex

            SaleDate = .dteSaleDate.Value
            DelDate = .lblDelDate.Text  ' .dteDelivery.Value
            PorD = .PorD
            StopStart = IfNullThenNilString(.dtpDelWindow.Value)
            StopEnd = IfNullThenNilString(.dtpDelWindow2.Value)
            Status = .SaleStatus.Text

            SalesCode = Trim(.vGetSalesCode)
            If SalesCode = "" Then SalesCode = "99"
            SalesSplit = .vGetSalesSplit
            If SalesSplit = "" Then SalesSplit = "100.0 0.0 0.0"
        End With

        With BillOSale
            N = .LastLineUsed
            For X = 0 To N
                If .QueryStyle(X) = "TAX1" Then TaxZone = 0 : Taxed = True
                If .QueryStyle(X) = "TAX2" Then TaxZone = .QueryQuan(X) + 1 : Taxed = True

                If Taxed Then
                    For Y = 0 To X - 1
                        'If IsItem(Item(Y + 1).Style) Then Items(Y + 1).NonTaxable = False
                        If IsItem(Item(Y + 1).Style) Then Items(Y).NonTaxable = False
                        'If Item(Y + 1).Style = "LAB" And StoreSettings.bLaborTaxable Then Items(Y + 1).NonTaxable = False
                        If Item(Y + 1).Style = "LAB" And StoreSettings.bLaborTaxable Then Items(Y).NonTaxable = False
                        'If Item(Y + 1).Style = "DEL" And StoreSettings.bDeliveryTaxable Then Items(Y + 1).NonTaxable = False
                        If Item(Y + 1).Style = "DEL" And StoreSettings.bDeliveryTaxable Then Items(Y).NonTaxable = False
                        'If Item(Y + 1).Style = "STAIN" Then Items(Y + 1).NonTaxable = False
                        If Item(Y + 1).Style = "STAIN" Then Items(Y).NonTaxable = False
                        'If Item(Y + 1).Style = "NOTES" Then Items(Y + 1).NonTaxable = False
                        If Item(Y + 1).Style = "NOTES" Then Items(Y).NonTaxable = False
                        'If Item(Y + 1).Style = "DISCOUNT" Then Items(Y + 1).NonTaxable = False
                        If Item(Y + 1).Style = "DISCOUNT" Then Items(Y).NonTaxable = False
                    Next
                    Taxed = False
                End If

                TaxThisItem = False
                'If .QueryStyle(X) = "DEL" And StoreSettings.DeliveryTaxable = SS_Taxable_ALWAYS Then TaxThisItem = True
                '  If .QueryStyle(X) = "LAB" And StoreSettings.LaborTaxable = SS_Taxable_ALWAYS Then TaxThisItem = True

                AddGenericItem(.QueryStyle(X), .QueryDesc(X), Val(.QueryQuan(X)), GetPrice(.QueryPrice(X)), GetPrice(.QueryPrice(X)), Val(.QueryLoc(X)), .QueryStatus(X), Not TaxThisItem, .QueryMfg(X), .QueryTransID(X))
            Next
        End With
    End Function

    Public ReadOnly Property SubTotal(Optional ByVal tType As String = "") As Decimal
        Get
            Dim I As Integer, Style As String, Status As String
            Dim Cost As Decimal, Txbl As Boolean, IsDel As Boolean
            Dim IsPT As cdsPayTypes

            IsPT = PayTypeIs(tType)
            '  Dim AddTax as decimal
            SubTotal = 0
            If ItemCount = 0 Then Exit Property
            tType = LCase(tType)
            For I = ItemCount To 1 Step -1
                Style = Items(I - 1).Style
                Status = Items(I - 1).Status
                IsDel = IsDelivered(Status)
                Select Case Style
                    Case "SUBTOTAL", "SUB", "--- Adj ---"
        ' Subtotals aren't real money, discounts are adjusted into item price.
                    Case "PAYMENT"
                        If IsIn(tType, "paid") Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        ElseIf IsIn(tType, "") Then
                            SubTotal = SubTotal - Items(I - 1).Price
                        ElseIf IsPT <> cdsPayTypes.cdsPT_NONE And IsPT = Items(I - 1).Quantity Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        End If
                    Case "SALESTAX", "TAX1", "TAX2"
                        If IsIn(tType, "", "gross", "tax", "tax1", "tax2") Then
                            If tType = "tax1" Then
                                If Style = "TAX1" Then SubTotal = SubTotal + Items(I - 1).Price
                            ElseIf tType = "tax2" Then
                                If Style = "TAX2" Then SubTotal = SubTotal + Items(I - 1).Price
                            Else
                                SubTotal = SubTotal + Items(I - 1).Price
                            End If
                        ElseIf IsIn(tType, "delivered") And IsDel Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        ElseIf IsIn(tType, "undelivered") And Not IsDel Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        End If
                    Case "LAB", "DEL", "STAIN"
                        '        BFH20120322 - the 'Nontaxable' must be set manually, not figured out here
                        '        Txbl = Switch(Style = "LAB", StoreSettings.bLaborTaxable, Style = "DEL", StoreSettings.bDeliveryTaxable, Style = "STAIN", True, True, False)
                        Txbl = Not Items(I - 1).NonTaxable
                        If IsIn(tType, "nontaxable", "taxable") Then
                            If tType = "taxable" And Txbl Or tType = "nontaxable" And Not Txbl Then
                                SubTotal = SubTotal + Items(I - 1).Price
                            End If
                        ElseIf IsIn(tType, "", "gross", "written") Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        ElseIf IsIn(tType, "stain", "lab", "del") Then
                            If LCase(Style) = tType Then SubTotal = SubTotal + Items(I - 1).Price
                        ElseIf IsIn(tType, "delivered") And IsDel Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        ElseIf IsIn(tType, "undelivered") And Not IsDel Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        End If
                    Case Else  ' including notes, discount, etc
                        If IsIn(tType, "", "gross", "written") Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        ElseIf IsIn(tType, "taxable") Then
                            If Not Items(I - 1).NonTaxable Then
                                SubTotal = SubTotal + Items(I - 1).Price
                            End If
                        ElseIf IsIn(tType, "nontaxable", "gross") Then
                            If Items(I - 1).NonTaxable Then
                                SubTotal = SubTotal + Items(I - 1).Price
                            End If
                        ElseIf IsIn(tType, "gm") Then
                            If Items(I - 1).Landed = 0 Then Items(I - 1).LoadPricing()
                            Cost = Cost + Items(I - 1).Landed
                            SubTotal = SubTotal + Items(I - 1).Price
                        ElseIf IsIn(tType, "layaway") Then
                            If Items(I - 1).Status = "LAW" Or Items(I - 1).Status = "SSLAW" Then
                                SubTotal = SubTotal + Items(I - 1).Price
                            End If
                        ElseIf IsIn(tType, "items") Then
                            If IsItem(Items(I - 1).Style) Then
                                SubTotal = SubTotal + Items(I - 1).Price
                            End If
                        ElseIf IsIn(tType, "delivered") And IsDel Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        ElseIf IsIn(tType, "undelivered") And Not IsDel Then
                            SubTotal = SubTotal + Items(I - 1).Price
                        End If
                End Select
NextItem:
            Next

            ' add tax not used...
            '  For I = 1 To ItemCount
            '    If items(I-1).Style = "DEL" And StoreSettings.DeliveryTaxable = SS_Taxable_ALWAYS Then
            '      AddTax = AddTax + items(I-1).Price * Val(StoreSettings.SalesTax)
            '    End If
            '    If items(I-1).Style = "LAB" And StoreSettings.LaborTaxable = SS_Taxable_ALWAYS Then
            '      AddTax = AddTax + items(I-1).Price * Val(StoreSettings.SalesTax)
            '    End If
            '    If IsIn(items(I-1).Style, "TAX1", "TAX2") Then
            '      AddTax = 0
            '    End If
            '  Next
            '
            '  If IsIn(tType, "tax", "tax1", "addtax") Then
            '    SubTotal = SubTotal + AddTax
            '  End If
            '
            '  If IsIn(tType, "gross", "written") Then SubTotal = SubTotal + AddTax

            If tType = "gm" Then
                Debug.Print("SubTotal=" & SubTotal, "Cost=" & Cost, "GM=" & CalculateGM(SubTotal, Cost))
                SubTotal = CalculateGM(SubTotal, Cost)
            End If
        End Get

    End Property

    Public Function ProcessSale(Optional ByVal SpecifiedSaleNo As String = "", Optional ByVal DoPrint As Boolean = True) As String ' return sale number, if success
        Dim Commable As String, NeedsSignature As Boolean
        Dim Cst As Decimal, Frt As Decimal
        Dim I As Integer, SaleName As String, SaleIndex As String, TDesc As String
        Dim DelStat As String

        Dim DDelDat As String, ShipDte As String
        Dim SaleDte As String
        Dim InvData As CInvRec

        Dim cMR As clsMailRec
        Dim Holding As cHolding

        Dim dV As String, dVN As String, dDpt As Integer, dLc As Integer

        If Not OkToProcess(TDesc, SpecifiedSaleNo) Then
            If TDesc <> "" Then MessageBox.Show(TDesc, "Sale not ready", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Function
        End If

        On Error GoTo ProcessSaleError

        ' If the sale's not ready to be completed, give a warning.
        ' Or better yet, don't have this button enabled at all.
        ' A cash&carry sale is not ready if:
        '   No items
        '   Cash due
        '   Change due?
        If IsProcessed Then Exit Function
        If ItemCount = 0 Then Exit Function

        If DeliverOnProcess Then
            DDelDat = Today
            ShipDte = Today
            DelStat = "DEL"
        Else
            DDelDat = DelDate
            ShipDte = ""
            DelStat = ""
        End If

        '  FinishSale

        If MailIndex = 0 Then
            SaleIndex = "0"
            If CashRegisterSale Then
                SaleName = "CASH REGISTER"
            Else
                SaleName = "CASH & CARRY"
            End If
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

        Name = SaleName

        ' Process the sale.
        ' Create holding record.
        ' The store really has to be in Auto-HoldingID mode.
        ' Add lines to GrossMargin.
        ' Decrement 2Data quantities.
        ' Add lines to Detail?
        ' Add lines to Cash+Sales journals.
        ' It would be very good to have a single function for "Add this item to this sale",
        ' with all detailed accounting taken care of within that.

        Holding = New cHolding
        Holding.LeaseNo = GetLeaseNumber(, SpecifiedSaleNo) ' Create a lease number.
        If IsFormLoaded("BillOSale") Then BillOSale.BillOfSale.Text = Holding.LeaseNo
        Holding.Deposit = SubTotal("paid")                  ' Amount paid.
        Holding.Sale = SubTotal("gross")                    ' Total amount of sale, with tax
        Holding.NonTaxable = SubTotal("nontaxable")         ' Amount that's not taxable..
        Holding.LastPay = Today                              ' Paid today.
        Holding.Salesman = SalesCode                        ' Who's logged in?
        Commable = IIf(Holding.Salesman = "", "", "C")
        Holding.Status = "V"                                ' Start it as void, then overwrite later
        Holding.Comm = "N"                                  ' Commission isn't paid
        Holding.Index = Val(SaleIndex)                      ' mailing index
        Holding.Save()
        ProcessSale = Holding.LeaseNo                       ' sale no is return value
        SaleNo = ProcessSale                                ' also stored in a property

        For I = 1 To ItemCount
            Application.DoEvents()

            With Items(I - 1)
                Select Case .Style
                    Case "PAYMENT", "CHANGE"    ' Save as payment.
                        ' Deal with description...
                        Dim DDDD As String
                        Dim DDDA As Decimal
                        DDDD = .Desc
                        If InStr(DDDD, DateFormat(Today)) = 0 And InStr(DDDD, Today) = 0 Then DDDD = DDDD & Space(5) & DateFormat(Today)
                        DDDA = .Price
                        If PayTypeIsOutsideFinance(.Quantity) And DDDA = SubTotal("gross") Then
                            DDDA = 0
                            .Price = 0
                            Holding.Deposit = 0
                            Holding.Save()
                        End If
                        AddNewMarginRecord(Holding.LeaseNo, "PAYMENT", DDDD, .Quantity, DDDA,
                        "", "", "", 0, 0, 0, "", "", DelStat, Holding.Salesman,
                        0, SaleDate, DDelDat, Store, SaleName, ShipDte, "",
                        SaleIndex, "0", 0, "", "", "", Nothing, .TransID, SalesSplit)
                        '          Dim Memo As String
                        'On Error Resume Next
                        '          Memo = Mid(split(.Desc, "/")(2), 6)
                        '          AddNewCashJournalRecord .Quantity, .Price, ProcessSale, Trim(SaleName & " " & Memo), DateFormat(SaleDate)
                        If Not PayTypeIsOutsideFinance(.Quantity) Then
                            'AddNewCashJournalRecord(.Quantity, .Price, ProcessSale, SaleName, DateFormat(SaleDate))
                            AddNewCashJournalRecord(.Quantity, .Price, ProcessSale, SaleName, Date.Parse(DateFormat(SaleDate), Globalization.CultureInfo.InvariantCulture))
                            'If IsDate(SaleDate) Then
                            '    AddNewCashJournalRecord(.Quantity, .Price, ProcessSale, SaleName, DateTime.Parse(Format(SaleDate, "MM/dd/yyyy")))
                            'Else
                            '    AddNewCashJournalRecord(.Quantity, .Price, ProcessSale, SaleName, Nothing)
                            'End If
                        End If
                    Case "NOTES"
                        Dim xGM As CGrossMargin
                        xGM = SaveNewMarginRecord(Holding.LeaseNo, "NOTES", .Desc, .Quantity, .Price,
            "", "", "", 0, 0, 0, PorD, "", DelStat, Holding.Salesman,
            Items(I - 1).Location, SaleDate, DDelDat, Store, SaleName, ShipDte, "", SaleIndex,
            , , , , , , , SalesSplit)
                        If .Vendor <> "" Then .MakePo(Me, xGM)
                        DisposeDA(xGM)
                    Case "DISCOUNT"             ' Save as a note, zero cost.
                        DiscountVendorAndDept(I, dV, dVN, dDpt, dLc)
                        If False Then
                            AddNewMarginRecord(Holding.LeaseNo, "NOTES", Format(- .Price / (Items(I - 1).Price - .Price), "0%") & " DISCOUNT (" & .Price & ")", 0, 0,
              dV, dDpt, dVN, 0, 0, 0, "", "", DelStat, Holding.Salesman,
              0, SaleDate, DDelDat, Store, SaleName, ShipDte, "", SaleIndex,
              , , , , , , , SalesSplit)
                        Else
                            AddNewMarginRecord(Holding.LeaseNo, "NOTES", .Desc, 0, .Price,
              dV, dDpt, dVN, 0, 0, 0, PorD, "", DelStat, Holding.Salesman,
              dLc, SaleDate, DDelDat, Store, SaleName, ShipDte, "", SaleIndex,
              , , , , , , , SalesSplit)
                        End If
                    Case "SALES TAX", "TAX1"
                        AddNewMarginRecord(Holding.LeaseNo, "TAX1", "SALES TAX", 1, .Price,
            "", "", "", 0, 0, 0, "", "", DelStat, Holding.Salesman,
            0, SaleDate, DDelDat, Store, SaleName, ShipDte, "", SaleIndex,
            , , , , , , , SalesSplit)
                        TaxZone = 0
                    Case "TAX2"
                        AddNewMarginRecord(Holding.LeaseNo, "TAX2", .Desc, Items(I - 1).Quantity, .Price,
            "", "", "", 0, 0, 0, "", "", DelStat, Holding.Salesman,
            Items(I - 1).Location, SaleDate, DDelDat, Store, SaleName, ShipDte, "", SaleIndex,
            , , , , , , , SalesSplit)
                        TaxZone = Items(I - 1).Quantity
                    Case "SUBTOTAL", "SUB"
                        AddNewMarginRecord(Holding.LeaseNo, "SUB", "Sub Total =", 0, .Price,
            "", "", "", 0, 0, 0, "", "", DelStat, Holding.Salesman,
            0, SaleDate, DDelDat, Store, SaleName, ShipDte, "", SaleIndex,
            , , , , , , , SalesSplit)
                    Case "LAB", "DEL", "STAIN"
                        If .Desc = "" Then
                            If .Style = "LAB" Then .Desc = "Labor Charges"
                            If .Style = "DEL" Then .Desc = "Delivery Charges"
                            If .Style = "STAIN" Then
                                If IsBFMyer Then
                                    .Desc = "SAFEWARE PROTECTION PLAN"
                                Else
                                    .Desc = "Stain Protection"
                                End If
                            End If
                        End If

                        If False Then   ' use this case to make S/L/D have dept/vend/loc.  Makes them show up on those reports sorted by those
                            DiscountVendorAndDept(I, dV, dVN, dDpt, dLc)
                            AddNewMarginRecord(Holding.LeaseNo, .Style, .Desc, 0, .Price,
              dV, dDpt, dVN, 0, 0, 0, PorD, "", DelStat, Holding.Salesman,
              dLc, SaleDate, DDelDat, Store, SaleName, ShipDte, "", SaleIndex,
              , , , , , , , SalesSplit)
                        Else
                            AddNewMarginRecord(Holding.LeaseNo, .Style, .Desc, 0, .Price,
              "", "", "", 0, 0, 0, PorD, "", DelStat, Holding.Salesman,
              0, SaleDate, DDelDat, Store, SaleName, ShipDte, "", SaleIndex,
              , , , , , , , SalesSplit)
                        End If

                    Case "--- Adj ---"
                        AddNewMarginRecord(Holding.LeaseNo, "--- Adj ---", .Desc, 0, 0,
            "", "", "", 0, 0, 0, "", "", DelStat, Holding.Salesman,
            0, SaleDate, DDelDat, Store, SaleName, ShipDte, "", SaleIndex,
            , , , , , , , SalesSplit)
                        NeedsSignature = True
                    Case Else ' Actual item style..
                        .AddItemGrossMargin(Me)
                End Select
            End With
        Next

        Holding.Status = NewSaleStatus(Holding.Status)
        Status = DescribeHoldingStatus(Holding.Status)
        Holding.ArNo = GetStoreFinanceArNo

        AddSalesJournal(ProcessSale)           ' Save an audit record.  Also add CASH record if necessary
        SalePackageUpdate(SaleNo)              ' update package pricing

        Holding.Save()

        DisposeDA(Holding)

        If DoPrint Then PrintInvoices(SaleNo)

        Exit Function

ProcessSaleError:
        ErrMsg("An error occurred while processing this sale." & vbCrLf & "Error [" & Err.Number & "]: " & Err.Description & vbCrLf & "sSale.ProcessSale")
    End Function

    Public Property SaleNo() As String
        Get
            SaleNo = mSaleNo
        End Get
        Set(value As String)
            mSaleNo = value
        End Set
    End Property

    Public Function PrintInvoice(Optional ByVal CopyID As String = COPY_CUSTOMER, Optional ByVal Copies As Integer = 1, Optional ByVal vLoadSaleNo As String = "") As Boolean
        Dim I As Integer, cHold As cHolding
        Dim Pages As Integer, Page As Integer

        On Error GoTo PrintInvoiceError

        If vLoadSaleNo <> "" Then LoadSaleNo(vLoadSaleNo)
        If Copies < 1 Then Exit Function
        cHold = New cHolding
        If Not cHold.Load(SaleNo, "LeaseNo") Then
            DisposeDA(cHold)
            Exit Function
        End If
        For I = 1 To Copies
            Pages = GetMaxPages(GetMaxItemIndex)
            For Page = 0 To Pages - 1
                PrintInvoiceCommon(CopyID, Page, Pages, cHold)
            Next
            Printer.EndDoc()
        Next

        DisposeDA(cHold)

        Exit Function

PrintInvoiceError:
        MessageBox.Show("A printer error has occured." & vbCrLf & Err.Description)
    End Function

    Public Sub Clear()
        mSaleNo = ""
        Tele = ""
        Name = ""
        MailIndex = -1
        CashRegisterSale = False

        SaleDate = Today
        DelDate = ""
        PorD = ""

        ItemCount = 0

        SalesCode = ""
        SalesSplit = ""

        CustType = 0
        AdvertizingType = 0
        TaxZone = 0
        LAB = 0
        DEL = 0
        STAIN = 0

        StopStart = ""
        StopEnd = ""

        ProcessSalePOs = Nothing
    End Sub

    Public Property Store() As Integer
        Get
            Store = mStore
            If Store <= 0 Then Store = StoresSld
        End Get
        Set(value As Integer)
            mStore = value
        End Set
    End Property

    Public ReadOnly Property Item(ByVal Index As Integer) As clsSaleItem
        Get
            Item = Nothing
            'If Index < LBound(Items) Or Index > UBound(Items) Then Exit Property
            If Index < LBound(Items) Or Index > (UBound(Items) + 1) Then Exit Property
            Item = Items(Index - 1)
        End Get
    End Property

    Public Function AddGenericItem(
ByVal Style As String, ByVal Desc As String, Optional ByVal Quantity As Double = 0,
Optional ByVal Price As Decimal = 0, Optional ByVal DisplayPrice As Decimal = 0,
Optional ByVal Location As Integer = 0, Optional ByVal Status As String = "",
Optional ByVal NonTaxable As Boolean = False, Optional ByVal SSVendor As String = "",
Optional ByVal TransID As String = "") As Boolean
        Dim I As clsSaleItem
        I = New clsSaleItem
        With I
            .Desc = UCase(Desc)
            .NonTaxable = NonTaxable
            .Price = Price
            .DisplayPrice = IIf(DisplayPrice = 0, .Price, DisplayPrice)
            .Quantity = Quantity
            .Style = Style
            If IsItem(.Style) Then .LoadVendor()
            .Status = Status
            .Location = Location
            .Vendor = SSVendor
            If .Vendor <> "" Then .VendorNo = GetVendorNoFromName(.Vendor)
            .TransID = TransID
        End With
        AddGenericItem = AddSaleItem(I)
        DisposeDA(I)
    End Function

    Public Function OkToProcess(ByRef Msg As String, Optional ByVal SpecifiedSaleNo As String = "") As Boolean
        If StoreSettings.bManualBillofSaleNo Then
            If SpecifiedSaleNo = "" Then
                Msg = "Please enter a Bill of Sale No."
                On Error Resume Next
                If IsFormLoaded("BillOSale") Then BillOSale.txtSaleNo.Select()
                Exit Function
            End If
            If LeaseNoExists(SpecifiedSaleNo) Then
                Msg = "Sale Number " & SpecifiedSaleNo & " Already Exists."
                Exit Function
            End If
        End If
        If Val(StoreSettings.SalesTax) <> 0 Then
            If SubTotal("tax") = 0 Then
                'If MsgBox("Sale tax not applied.  Apply Sales tax now?", vbQuestion + vbYesNo, "Apply Sales Tax") = vbYes Then
                If MessageBox.Show("Sale tax not applied.  Apply Sales tax now?", "Apply Sales Tax", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Exit Function
                End If
            End If
        End If
        If WantCheckDisposal() And SaleHasBedding() And Not SaleHasDisposal() Then
            'If MsgBox("Bedding is indicated on the sale, but no disposal has been charged." & vbCrLf2 & "Did you charge a disposal fee?", vbQuestion + vbOKCancel) = vbCancel Then
            If MessageBox.Show("Bedding is indicated on the sale, but no disposal has been charged." & vbCrLf2 & "Did you charge a disposal fee?", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
                Exit Function
            End If
        End If

        OkToProcess = True
    End Function

    Public ReadOnly Property IsProcessed() As Boolean
        Get
            IsProcessed = (SaleNo <> "")
        End Get
    End Property

    Public ReadOnly Property DeliverOnProcess() As Boolean
        Get
            If IsProcessed Then Exit Property
            If Count("items") = 0 Then Exit Property
            If Count("items") = Count("deltw") Then DeliverOnProcess = True
        End Get
    End Property

    Public ReadOnly Property Count(Optional ByVal tType As String = "") As Integer
        Get
            Dim I As Integer
            If ItemCount = 0 Then Exit Property
            tType = Trim(LCase(tType))

            If tType = "" Then Count = ItemCount : Exit Property

            For I = 1 To ItemCount
                With Items(I - 1)
                    Select Case Items(I - 1).Style
                        Case "STAIN", "LAB", "DEL"
                            If LCase(Items(I - 1).Style) = tType Then Count = Count + 1
                        Case "TAX1", "TAX2"
                            If IsIn(tType, "tax") Then Count = Count + 1
                        Case "PAYMENT"
                            If IsIn(tType, "payment") Then Count = Count + 1
                        Case "--- Adj ---"
                            If IsIn(tType, "adj") Then Count = Count + 1
                        Case Else
                            If IsIn(tType, "items") And IsItem(Items(I - 1).Style) Then Count = Count + 1
                            If LCase(Items(I - 1).Status) = tType Then Count = Count + 1
                    End Select
                End With
            Next
        End Get
    End Property

    Private Sub DiscountVendorAndDept(ByVal I As Integer, ByRef V As String, ByRef vN As String, ByRef Dept As Integer, ByRef vLoc As Integer)
        Dim J As Integer
        Dim SD As String

        On Error Resume Next
        For J = I To 1 Step -1
            If IsItem(QueryStyle(J)) And Not (WantCheckDisposal() And GetDeptNoFromStyle(QueryStyle(J)) = DisposalDepartment()) Then
                V = QueryMfg(J)
                vN = GetVendorNoFromName(V)
                SD = GetDeptFromStyleNo(QueryStyle(J))
                If SD = "" Then Exit Sub
                'Dept = CLng(SD)
                Dept = CInt(SD)
                vLoc = QueryLoc(J)
                Exit Sub
            End If
        Next
    End Sub

    Public Function NewSaleStatus(Optional ByVal OldStatus As String = "") As String
        NewSaleStatus = "O" ' Default

        If SubTotal("layaway") <> 0 Or IsIn(OldStatus, "1", "2", "3", "4") Then
            NewSaleStatus = "L"
        End If

        If IsUFO() Then ' Or (IsFriendlys() And OldStatus = "L") Then      ' Get layaway terms.
            'LaAwaySelect.Show vbModal
            LaAwaySelect.ShowDialog()
            If LaAwaySelect.opt30.Checked = True Then NewSaleStatus = "1"
            If LaAwaySelect.opt60.Checked = True Then NewSaleStatus = "2"
            If LaAwaySelect.opt90.Checked = True Then NewSaleStatus = "3"
            If LaAwaySelect.opt120.Checked = True Then NewSaleStatus = "4"
        End If

        If SubTotal = 0 Then
            If NoItemsOnSale And HasNonItemsOnSale(True, True, True) Then
                ' don't force delivery..
            ElseIf AllItemsAreDelivered Then
                NewSaleStatus = "D"
            End If
        End If

        If IsStoreFinanceSale Then
            NewSaleStatus = "S"
        End If

        If IsCreditSale Then
            NewSaleStatus = "E"
        End If
    End Function

    Public ReadOnly Property IsCreditSale() As Boolean
        Get
            Dim I As Integer
            Dim RV As Decimal
            For I = 1 To ItemCount
                If Item(I).Style = "PAYMENT" Then
                    If PayTypeIsOutsideFinance(Left(Item(I).Desc, 14)) Then
                        RV = Item(I).DisplayPrice
                        If RV = 0 Or RV = SubTotal Then    ' Either zero price or rest of total
                            IsCreditSale = True
                            Exit Property
                        End If
                    End If
                End If
            Next
        End Get

    End Property

    Public ReadOnly Property HasNonItemsOnSale(Optional ByVal STAIN As Boolean = True, Optional ByVal Delivery As Boolean = True, Optional ByVal Labor As Boolean = True) As Boolean
        Get
            Dim I As Integer
            For I = 1 To ItemCount
                If STAIN And Item(I).Style = "STAIN" Then HasNonItemsOnSale = True : Exit Property
                If Labor And Item(I).Style = "LAB" Then HasNonItemsOnSale = True : Exit Property
                If Delivery And Item(I).Style = "DEL" Then HasNonItemsOnSale = True : Exit Property
            Next

        End Get
    End Property

    Public ReadOnly Property AllItemsAreDelivered() As Boolean
        Get
            Dim I As Integer
            For I = 1 To ItemCount
                If IsItem(Item(I).Style) And Not IsDelivered(Item(I).Status) Then AllItemsAreDelivered = False : Exit Property
            Next
            AllItemsAreDelivered = True
        End Get

    End Property

    Public ReadOnly Property IsStoreFinanceSale() As Boolean
        Get
            Dim I As Integer
            For I = 1 To ItemCount
                If Item(I).Style = "NOTES" Then
                    If Left(Item(I).Desc, 13) = "STORE FINANCE" Then IsStoreFinanceSale = True : Exit Property
                End If
            Next
        End Get
    End Property

    Public ReadOnly Property GetStoreFinanceArNo() As String
        Get
            GetStoreFinanceArNo = ""
            Dim I As Integer
            For I = 1 To ItemCount
                If Item(I).Style = "NOTES" Then
                    If Left(Item(I).Desc, 13) = "STORE FINANCE" Then
                        GetStoreFinanceArNo = Mid(Item(I).Desc, InStr(1, Item(I).Desc, " Account #", vbTextCompare) + 10)
                        Exit Property
                    End If
                End If
            Next
        End Get
    End Property

    Private Function AddSalesJournal(ByVal SaleNo As String) As Boolean
        Dim HS As String, AN As String, Tp As Decimal

        HS = HoldingStatusRepresents(Status)
        Tp = -(SubTotal("paid") -
            SubTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance)) -
            SubTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance2)) -
            SubTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance3)) -
            SubTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance4)) -
            SubTotal(PayListItem(cdsPayTypes.cdsPT_OutsideFinance5)) -
            SubTotal(PayListItem(cdsPayTypes.cdsPT_StoreFinance)))

        Select Case HS
            Case "D"
                AN = "PT " & Name
                AddNewAuditRecord(SaleNo, AN, SaleDate, SubTotal("written"), SubTotal("tax"), 0, 0, 0, SubTotal("written"), SubTotal("tax"), TaxZone, SalesCode, SubTotal("nontaxable"), GetCashierName)
            Case "S"
                AN = "SF " & Name
                If StoreFinanceAsDelivered Then
                    'BFH20161026 - The following line would create a DELIVERED "Open Store Finance" sale.  We replaced it one that represents the sale as "Open" in the Audit table (esp for the Sales Tax reprot)
                    AddNewAuditRecord(SaleNo, AN, SaleDate, SubTotal("written"), SubTotal("tax"), 0, SubTotal("gross"), 0, SubTotal("written"), SubTotal("tax"), TaxZone, SalesCode, SubTotal("nontaxable"), GetCashierName)
                Else
                    AddNewAuditRecord(SaleNo, AN, SaleDate, SubTotal("written"), SubTotal("tax"), SubTotal("gross"), Tp, SubTotal("gross"), 0, 0, TaxZone, SalesCode, SubTotal("nontaxable"), GetCashierName)
                End If


' These were the lines to create the back order (receivable) balance..  We now do this when delivered
'      tP = SubTotal(PayListItem(cdsPT_OutsideFinance))
'      If tP <> 0 Then
'        AddNewCashJournalRecord "11300", tP, SaleNo, "", Date
'      End If
'      AddNewCashJournalRecord "11300", SubTotal(), SaleNo, "", Date

'BFH20161031
' "C" was "Credit" (outside finance), now is "Closed Credit".
' New Sales are now created as "Open Finance Sale" ("E")
'    Case "C"
'      an = "NS " & Name
'      AddNewAuditRecord SaleNo, an, SaleDate, SubTotal("written"), SubTotal("tax"), SubTotal("gross"), -SubTotal("gross"), SubTotal("gross"), 0, 0, TaxZone, SalesCode, SubTotal("nontaxable"), GetCashierName
'      AddNewCashJournalRecord "11300", SubTotal, SaleNo, "", Date
            Case "E"
                AN = "NS " & Name
                'BFH20161031
                ' We now have "Open Credit".  Most of these old ways will be done when the sale is delivered,
                ' Rather than at sale creation, which is no longer set up in the delivered state.
                '      AddNewAuditRecord SaleNo, an, SaleDate, SubTotal("written"), SubTotal("tax"), SubTotal("gross"), -SubTotal("gross"), SubTotal("gross"), 0, 0, TaxZone, SalesCode, SubTotal("nontaxable"), GetCashierName
                '      AddNewCashJournalRecord "11300", SubTotal, SaleNo, "", Date
                AddNewAuditRecord(SaleNo, AN, SaleDate, SubTotal("written"), SubTotal("tax"), SubTotal("gross"), Tp, SubTotal("gross"), 0, 0, TaxZone, SalesCode, SubTotal("nontaxable"), GetCashierName)
                '      AddNewCashJournalRecord "11300", SubTotal(PayListItem(cdsPT_OutsideFinance)), SaleNo, "", Date
            Case Else
                AN = "NS " & Name
                AddNewAuditRecord(SaleNo, AN, SaleDate, SubTotal("written"), SubTotal("tax"), SubTotal("gross"), Tp, SubTotal("gross"), 0, 0, TaxZone, SalesCode, SubTotal("nontaxable"), GetCashierName)
                '      tP = SubTotal(PayListItem(cdsPT_OutsideFinance))
                '      If tP <> 0 Then
                '        AddNewCashJournalRecord "11300", tP, SaleNo, "", Date
                '      End If
        End Select
    End Function

    Private Function PrintInvoices(ByVal vSaleNo As String)
        Dim Copies As Integer, CopyID As String, S As sSale
        For Copies = 1 To Val(StoreSettings.PrintCopies)
            If Copies <= 4 Then CopyID = StoreSettings.SalesCopyID(FitRange(0, Copies - 1, 3))
            S = New sSale
            S.PrintInvoice(CopyID, 1, vSaleNo)
            DisposeDA(S)
        Next
    End Function

    Public Function LoadSaleNo(ByVal vSaleNo As String, Optional ByVal vStore As Integer = 0) As Boolean
        Dim Hold As cHolding, Gross As CGrossMargin, cM As clsMailRec
        Dim Taxed As Boolean, Y As Integer, ClearingStart As Integer
        Dim TaxThisItem As Boolean

        If vStore = 0 Then vStore = StoresSld
        mStore = vStore
        Clear()
        mSaleNo = vSaleNo
        Hold = New cHolding

        Hold.DataAccess.DataBase = GetDatabaseAtLocation(vStore)
        If Not Hold.Load(SaleNo, "LeaseNo") Then
            DisposeDA(Hold)
            Exit Function
        End If
        Status = DescribeHoldingStatus(Hold.Status)

        MailIndex = Hold.Index
        SalesCode = Hold.Salesman

        cM = New clsMailRec
        If cM.Load(Hold.Index, "#index") Then
            CustType = cM.CustType
            AdvertizingType = Val(cM.MailType)
            Name = cM.Last
            Tele = cM.Tele
        End If
        DisposeDA(cM)

        Dim FromLoop As Boolean
        Gross = New CGrossMargin
        Gross.DataAccess.DataBase = GetDatabaseAtLocation(Store)
        With Gross
            .DataAccess.Records_OpenSQL("SELECT * FROM [GrossMargin] WHERE SaleNo='" & SaleNo & "' ORDER BY [MarginLine]")
            .DataAccess.Records_Available()
            '.Load SaleNo, "SaleNo"
            FromLoop = False
            .cDataAccess_GetRecordSet(.DataAccess.RS)
            ' set these on the sale itself, from each individual line item...  little redundant but effective
            PorD = Gross.PorD
            SaleDate = Gross.SellDte
            DelDate = Gross.DDelDat
            StopStart = Gross.StopStart
            StopEnd = Gross.StopEnd

            Do
                If FromLoop = True Then
                    .cDataAccess_GetRecordSet(.DataAccess.RS)
                End If
                If .Style = "TAX1" Then TaxZone = 0 : Taxed = True
                If .Style = "TAX2" Then TaxZone = .Quantity + 1 : Taxed = True
                If IsADJ(.Style) Then ClearingStart = ItemCount + 1

                If Taxed Then
                    For Y = ClearingStart To ItemCount - 1
                        'If IsItem(Item(Y + 1).Style) Then Items(Y + 1).NonTaxable = False
                        If IsItem(Item(Y + 1).Style) Then Items(Y).NonTaxable = False
                        'If Item(Y + 1).Style = "LAB" And StoreSettings.bLaborTaxable Then Items(Y + 1).NonTaxable = False
                        If Item(Y + 1).Style = "LAB" And StoreSettings.bLaborTaxable Then Items(Y).NonTaxable = False
                        'If Item(Y + 1).Style = "DEL" And StoreSettings.bDeliveryTaxable Then Items(Y + 1).NonTaxable = False
                        If Item(Y + 1).Style = "DEL" And StoreSettings.bDeliveryTaxable Then Items(Y).NonTaxable = False
                        'If Item(Y + 1).Style = "STAIN" Then Items(Y + 1).NonTaxable = False
                        If Item(Y + 1).Style = "STAIN" Then Items(Y).NonTaxable = False
                        'If Item(Y + 1).Style = "NOTES" Then Items(Y + 1).NonTaxable = False
                        If Item(Y + 1).Style = "NOTES" Then Items(Y).NonTaxable = False

                        If IsADJ(Item(Y + 1).Style) Then GoTo DoneClearing
                    Next
DoneClearing:
                    Taxed = False
                End If

                TaxThisItem = False
                If .Style = "DEL" And StoreSettings.bDeliveryTaxable Then TaxThisItem = True
                If .Style = "LAB" And StoreSettings.bLaborTaxable Then TaxThisItem = True

                AddGenericItem(.Style, .Desc, .Quantity, .SellPrice, .SellPrice, .Location, .Status, Not TaxThisItem, .Vendor)
                FromLoop = True
            Loop While Gross.DataAccess.Records_Available

        End With
        DisposeDA(Gross, Hold)

        'MainMenu.rtbStorePolicy.RichTextBox.TextRTF = BillOSale.rtbStorePolicy.RichTextBox.TextRTF
        MainMenu.rtbStorePolicy.RichTextBox.Rtf = BillOSale.rtbStorePolicy.RichTextBox.Rtf

        'MainMenu.rtbn.RichTextBox.TextRTF = BillOSale.rtb.RichTextBox.TextRTF
        MainMenu.rtbn.RichTextBox.Rtf = BillOSale.rtb.RichTextBox.Rtf

        LoadSaleNo = True
    End Function

    Private Function GetMaxPages(ByVal Items As Integer) As Integer
        GetMaxPages = (Items \ 17) + 1
    End Function

    Private Function GetMaxItemIndex() As Integer
        Dim I As Integer, X As Integer
        X = 0
        For I = 1 To ItemCount
            X = X + ((Len(QueryDesc(I)) - 1) \ 46 + 1)
        Next
        GetMaxItemIndex = X
        'GetMaxItemIndex = X - 1
    End Function

    Public Function QueryDesc(ByVal Index As Integer)
        'QueryDesc = Left(Items(Index).Desc, Setup_2Data_DescMaxLen)
        QueryDesc = Left(Items(Index - 1).Desc, Setup_2Data_DescMaxLen)
    End Function

    Public Property LAB() As Decimal
        Get
            LAB = mLab
        End Get
        Set(value As Decimal)
            mLab = IIf(value < 0, 0, value)
        End Set
    End Property

    Public Property DEL() As Decimal
        Get
            DEL = mDel
        End Get
        Set(value As Decimal)
            mDel = IIf(value < 0, 0, value)
        End Set
    End Property

    Public Property STAIN() As Decimal
        Get
            STAIN = mStain
        End Get
        Set(value As Decimal)
            mStain = IIf(value < 0, 0, value)
        End Set
    End Property

    Public Function AddSaleItem(ByVal Itm As clsSaleItem) As Boolean
        ItemCount = ItemCount + 1
        ReDim Preserve Items(0 To ItemCount - 1)
        Items(ItemCount - 1) = Itm
        AddSaleItem = True
    End Function

    Public Function SaleHasBedding() As Boolean
        Dim DeptChk As Integer

        If IsSleepCity Then DeptChk = 0
        If IsBarrs Then DeptChk = 6
        Dim SD As String
        Dim I As Integer, S As String, D As Integer
        For I = 1 To ItemCount
            S = Items(I - 1).Style
            SD = GetDeptFromStyleNo(S)
            If SD = "" Then Exit Function
            'D = CLng(SD)
            D = CInt(SD)
            If D = DeptChk Then
                SaleHasBedding = True
                Exit Function
            End If
        Next
    End Function

    Public Function SaleHasDisposal() As Boolean
        Dim DeptChk As Integer
        Dim SD As String

        DeptChk = DisposalDepartment()
        '  If IsSleepCity Then DeptChk = 10
        '  If IsBarrs Then DeptChk = 10
        Dim I As Integer, S As String, D As Integer
        For I = 1 To ItemCount
            S = Items(I - 1).Style
            SD = GetDeptFromStyleNo(S)
            If SD = "" Then Exit Function
            'D = CLng(SD)
            D = CInt(SD)
            If D = DeptChk Then
                SaleHasDisposal = True
                Exit Function
            End If
        Next
    End Function

    Public Function QueryStyle(ByVal Index As Integer)
        'QueryStyle = Left(Items(Index).Style, Setup_2Data_StyleMaxLen)
        QueryStyle = Left(Items(Index - 1).Style, Setup_2Data_StyleMaxLen)
    End Function

    Public Function QueryMfg(ByVal Index As Integer)
        'QueryMfg = Items(Index).Vendor
        QueryMfg = Items(Index - 1).Vendor
    End Function

    Public Function QueryLoc(ByVal Index As Integer)
        'QueryLoc = Items(Index).Location
        QueryLoc = Items(Index - 1).Location
    End Function

    Public ReadOnly Property NoItemsOnSale() As Boolean
        Get
            Dim I As Integer
            For I = 1 To ItemCount
                If IsItem(Item(I).Style) Then Exit Property
            Next
            NoItemsOnSale = True
        End Get
    End Property

    Public ReadOnly Property PoNo(SaleItem As clsSaleItem) As Integer
        Get
            Dim tPO As String
            If ProcessSalePOs Is Nothing Then ProcessSalePOs = New Collection
            On Error Resume Next
            If SaleItem.Status = "SO" Or SaleItem.Status = "SS" Or (SaleItem.Style = "NOTES" And SaleItem.Vendor <> "") Then
                tPO = ""
                tPO = ProcessSalePOs(SaleItem.Vendor)
                If tPO = "" Then
                    tPO = GetPoNo()
                    ProcessSalePOs.Add(tPO, SaleItem.Vendor)
                End If
                PoNo = tPO
            End If
        End Get
    End Property

    Private Sub PrintInvoiceCommon(ByVal CopyID As String, ByVal Page As Integer, ByVal Pages As Integer, ByVal Holding As cHolding)
        Dim Xx As String, W As Integer
        Dim BoxLeft As Integer, BoxWidth As Integer
        Dim Sp As String, SpInst() As String, SpLoop As Object
        Dim LoopRow As Integer, Item As Integer, ItemLine As Integer
        Dim MfgForm As String, StyleForm As String, DescForm As String, PriceForm As String
        Dim LocForm As String, StatusForm As String, Quanform As String
        Dim C As CInvRec
        'Dim Logo As StdPicture
        Dim Logo As Image
        Dim ML As clsMailRec, M2 As MailNew2

        Dim SS() As Object, Sales1 As String, Sales2 As String, Sales3 As String
        Dim SSS As Object

        On Error Resume Next
        SSS = Split(SalesCode, " ")
        Sales1 = ""
        Sales1 = SSS(0)
        If Sales1 <> "" Then Sales1 = TranslateSalesman(Sales1)
        Sales2 = ""
        Sales2 = SSS(1)
        If Sales2 <> "" Then Sales2 = TranslateSalesman(Sales2)
        Sales3 = ""
        Sales3 = SSS(2)
        If Sales3 <> "" Then Sales3 = TranslateSalesman(Sales3)
        On Error GoTo 0

        'Logo = LoadPictureStd(StoreLogoFile(Store))
        Logo = Image.FromFile(StoreLogoFile(Store))

        '-------
        'Dim pbox As New PictureBox
        'pbox.Size = New Size(400, 120)
        'pbox.SizeMode = PictureBoxSizeMode.StretchImage
        'pbox.Image = Image.FromFile(StoreLogoFile(Store))
        '--------

        ML = New clsMailRec
        ML.DataAccess.DataBase = GetDatabaseAtLocation(Store)
        ML.Load(MailIndex, "#Index")
        GetMailNew2ByIndex(MailIndex, M2, Store)

        On Error GoTo HandleErr
        Printer.FontName = "Arial"
        Printer.FontSize = 18
        Printer.DrawWidth = 2
        If Printer.FontName <> "Arial" Or Printer.FontSize <> 18 Then
            MessageBox.Show("The computer could not set the proper font." & vbCrLf &
            "The Bill of Sale is designed to print in the font 'Arial' with size 18." & vbCrLf &
            "Attempting to print in " & Printer.FontName & ", size " & Printer.FontSize & "." & vbCrLf &
            "This could cause misalignment in the printout.", "Unable to set font", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End If
        Printer.FontBold = True
        Printer.CurrentY = 100

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   Logo (center)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If IsNothingOrZero(Logo) Then      ' street address
        If IsNothing(Logo) Then
            Printer.CurrentX = (6400) - Printer.TextWidth(Trim(StoreSettings.Name)) / 2
            Printer.Print(StoreSettings.Name)
            Printer.CurrentX = (6400) - Printer.TextWidth(Trim(StoreSettings.Address)) / 2
            Printer.Print(StoreSettings.Address)
            Printer.CurrentX = (6400) - Printer.TextWidth(Trim(StoreSettings.City)) / 2
            Printer.Print(StoreSettings.City)
            Printer.CurrentX = (6400) - Printer.TextWidth(Trim(StoreSettings.Phone)) / 2
            Printer.Print(StoreSettings.Phone)
        Else                  ' logo
            Printer.CurrentX = 4000
            '      Printer.PaintPicture Logo, Printer.Width / 2 - 5775 / 2, 150, 5775, 1525 '1995
            Dim opW As Integer, opH As Integer
            opW = Logo.Width
            opH = Logo.Height
            'PictureFitDimensions(opW, opH, 5775, 1525, True)
            'PictureFitDimensions(opW, opH, 400, 120, True)
            'Printer.ScaleMode = ScaleModeConstants.vbTwips
            'Printer.PaintPicture(Logo, Printer.Width / 2 - opW / 2, 150 + (1525 - opH) / 2, opW, opH)
            'Printer.PaintPicture(pbox.Image, 4000, 200, 1300, 400, 8000, 1150)
            'Printer.PaintPicture(pbox.Image, 4000, 200, 4000, 1200, 8000, 1150)
            'Printer.PaintPicture(Logo, 3500, 150 + (1525 - opH) / 2, opW, opH)
            'Printer.PaintPicture(Image.FromFile(StoreLogoFile(Store)), Printer.Width / 2 - opW / 2, 150, opW, opH)
            'Printer.PaintPicture(Image.FromFile(StoreLogoFile(Store)), 4000, 200, 2000, 2000)
            'Printer.PaintPicture(Logo, Printer.Width / 2 - opW / 2, 150, opW, opH, 9500, 1800)
            'Printer.PaintPicture(pbox.Image, 3500, 150, pbox.Width, pbox.Height, 6000, 300, pbox.Width + 300, pbox.Height + 300)
            'Printer.PaintPicture(pbox.Image, 3500, 150, pbox.Width, pbox.Height)

            Printer.PaintPicture(Image.FromFile(StoreLogoFile(0)), 4000, 200, 5000, 5000, 1200, 1000, 35000, 35000)
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   Date side
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Rectangle box
        Printer.Line(1100, 700, 3100, 1500, , True)
        ' print line inside rectangle box
        Printer.Line(1100, 1100, 3100, 1100)

        Printer.CurrentX = 0
        Printer.CurrentY = 100

        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.Print("     Date:")

        ' current date
        Printer.FontSize = 14
        Printer.CurrentX = 1000
        Printer.CurrentY = 100

        'Dim d As Date
        If IsUFO() Or IsSleepingSystems() Then
            Printer.Print(DateFormat(SaleDate) & "  " & TimeFormat(DateAndTime.TimeOfDay))
            'd = Date.Parse(DateFormat(SaleDate), CultureInfo.InvariantCulture)
            'Printer.Print(d & "  " & TimeFormat(DateAndTime.TimeOfDay))
        Else
            Printer.Print(DateFormat(SaleDate))
            'd = Date.Parse(DateFormat(SaleDate), CultureInfo.InvariantCulture)
            'Printer.Print(d)
        End If
        Printer.Print()

        Printer.CurrentX = 0
        Printer.CurrentY = 880

        Printer.FontSize = 10
        Printer.Print("Delivery:")

        If PorD = "D" Then
            Printer.CurrentX = 875
            Printer.CurrentY = 900
            Printer.Print("X")
        End If

        Printer.FontSize = 14
        Printer.CurrentX = 1400
        Printer.CurrentY = 700

        ' Day of Week
        Printer.FontBold = True
        Printer.Print(Format(DelDate, "ddd"))
        Printer.FontBold = False

        Printer.CurrentX = 0
        Printer.CurrentY = 1130
        Printer.FontSize = 10
        Printer.Print(" Pick Up:")

        If PorD = "P" Then
            Printer.CurrentX = 885
            Printer.CurrentY = 1110
            Printer.Print("X")
        End If

        Printer.FontSize = 14
        Printer.CurrentX = 1400

        ' delivery date
        Printer.FontSize = 14
        Printer.CurrentY = 1200
        Printer.CurrentX = 1400
        Printer.Print(DelDate)

        Dim twA As String, twB As String, twS As String
        twA = "" & StopStart ' BillOSale.dtpDelWindow(0).Value
        twB = "" & StopEnd ' BillOSale.dtpDelWindow(1).Value
        If StoreSettings.bUseTimeWindows And (twA <> "" Or twB <> "") Then
            Printer.Line(400, 1450, 3100, 1850, QBColor(0), True)

            Printer.FontSize = 9
            Printer.CurrentX = 500
            Printer.CurrentY = 1550
            PrintInBox(Printer, DescribeTimeWindow(twA, twB), 600, 1550, 2500, 300)
            '      Printer.Print twS
        End If



        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   Lease No side
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Printer.CurrentY = 100
        Printer.CurrentX = 9800 '10000

        Printer.FontSize = 18
        Printer.FontBold = True
        Printer.Print(Trim(SaleNo))
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.CurrentX = 10000
        Printer.Print("  Sale No:")

        Printer.CurrentX = 9600 '10100
        Printer.Print(Status)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   Addresses
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Print frame for address
        Printer.Line(0, 2000, 5500, 4850, QBColor(0), True)
        Printer.Line(6000, 2000, 11000, 4850, QBColor(0), True)

        Printer.CurrentX = 200
        Printer.CurrentY = 2200
        Printer.FontSize = 6

        If Not ML.Business Then
            Printer.Print("First Name", TAB(78), "Last Name")
        Else 'company
            Printer.Print(" Company")
        End If

        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 200
        Printer.Print("Address")
        Printer.Print()
        Printer.Print()
        Printer.Print()

        Printer.CurrentX = 200
        Printer.Print("Additional Address")
        Printer.Print()
        Printer.Print()
        Printer.Print()

        Printer.CurrentX = 200
        Printer.Print("City / State", TAB(75), "Zip")
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 200
        'Printer.Print "Telephone1"; Tab(58); "Telephone2"
        Printer.Print(ML.PhoneLabel1, TAB(58), ML.PhoneLabel2)
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 200
        Printer.Print()
        Printer.CurrentX = 200
        Printer.CurrentY = 4950
        Printer.FontSize = 18

        Printer.Print("Special ")
        Printer.CurrentX = 200
        Printer.Print("Instructions: ")
        Printer.FontSize = 10

        ' Ship to
        Printer.CurrentX = 6200 : Printer.CurrentY = 2100
        Printer.FontSize = 14
        Printer.Print("                SHIP TO ADDRESS:")

        Printer.FontSize = 6
        Printer.CurrentX = 6200 : Printer.CurrentY = 2400
        Printer.Print("First", SPC(37), "Last/Company")

        Printer.CurrentX = 6200 : Printer.CurrentY = 2800
        Printer.Print("Address")
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 6200
        Printer.Print("City / State", SPC(58), "Zip")
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 6200
        Printer.Print(M2.PhoneLabel3)
        '    Printer.Print "Telephone3 "

        Printer.FontSize = 10      ' special inst
        Printer.CurrentY = 5000    ' special instructions

        'special Instructions on separate lines
        Sp = WrapLongTextByPrintWidth(Printer, ML.Special, Printer.ScaleWidth - 2700)
        SpInst = Split(Sp, vbCrLf)
        For Each SpLoop In SpInst
            Printer.CurrentX = 2700
            Printer.Print(IfNullThenNilString(SpLoop))
        Next

        Printer.CurrentX = 6200 : Printer.CurrentY = 4700
        Printer.FontSize = 8.4
        Printer.Print("Sales Staff: ")

        Printer.CurrentX = 7200
        Printer.CurrentY = 4650
        Printer.FontSize = 11
        Printer.Print(Sales1, SPC(2), Sales2, SPC(2), Sales3)

        ' Desc line for inventory
        Printer.FontSize = 8.04

        Printer.CurrentX = 200 : Printer.CurrentY = 5870

        If CopyID = COPY_CUSTOMER Then
            Printer.Print(TAB(85), "Loc", SPC(2), "Status", SPC(2), "Quantity", SPC(2), "Description", SPC(90), "Price")
            BoxLeft = 3800
            BoxWidth = 14800 - BoxLeft
            'BoxWidth = 11000 - BoxLeft
            'Printer.Line(0, 5800, 3700, 7250, QBColor(0), True)
        Else
            Printer.Print("Style Number", SPC(10), "Manufacturer", SPC(30), "Loc ", "Status", SPC(2), "Quantity", SPC(5), "Description", SPC(83), "Price")
            BoxLeft = 0
            BoxWidth = 11000
        End If

        'heading box
        Printer.Line(BoxLeft, 5800, BoxWidth, 6150, QBColor(0), True)

        ' Inventory line boxes.
        Printer.Line(BoxLeft, 6300, BoxWidth, 6650, QBColor(0), True)
        Printer.Line(BoxLeft, 6700, BoxWidth, 7050, QBColor(0), True)
        Printer.Line(BoxLeft, 7100, BoxWidth, 7450, QBColor(0), True)
        Printer.Line(BoxLeft, 7500, BoxWidth, 7850, QBColor(0), True)
        Printer.Line(BoxLeft, 7900, BoxWidth, 8250, QBColor(0), True)
        Printer.Line(BoxLeft, 8300, BoxWidth, 8650, QBColor(0), True)
        Printer.Line(BoxLeft, 8700, BoxWidth, 9050, QBColor(0), True)
        Printer.Line(BoxLeft, 9100, BoxWidth, 9450, QBColor(0), True)
        Printer.Line(BoxLeft, 9500, BoxWidth, 9850, QBColor(0), True)
        Printer.Line(BoxLeft, 9900, BoxWidth, 10250, QBColor(0), True)
        Printer.Line(BoxLeft, 10300, BoxWidth, 10650, QBColor(0), True)
        Printer.Line(BoxLeft, 10700, BoxWidth, 11050, QBColor(0), True)
        Printer.Line(BoxLeft, 11100, BoxWidth, 11450, QBColor(0), True)
        Printer.Line(BoxLeft, 11500, BoxWidth, 11850, QBColor(0), True)
        Printer.Line(BoxLeft, 11900, BoxWidth, 12250, QBColor(0), True)
        Printer.Line(BoxLeft, 12300, BoxWidth, 12650, QBColor(0), True)
        Printer.Line(BoxLeft, 12700, BoxWidth, 13050, QBColor(0), True)
        '  Printer.Line (BoxLeft, 13100)-Step(BoxWidth, 350), QBColor(0), B  ' Removed 20030808 to make room for personal info
        '  Printer.Line (BoxLeft, 13500)-Step(11BoxWidth375, 350), QBColor(0), B

        ' If (page + 1) = Pages Then

        'new box on left - customer policy
        Printer.DrawWidth = 7
        Printer.Line(0, 13100, 8500, 14900, QBColor(0), True)
        Printer.DrawWidth = 1

        Dim x As String = ""
        If CopyID = COPY_CUSTOMER Then
            'Dim x As String
            'MainMenu.rtbn.DoPrintFile(StorePolicyMessageFile, 100, 6300, 3500, 7000, True, False)
            MainMenu.rtbn.mRichTextBox.LoadFile(StorePolicyMessageFile, RichTextBoxStreamType.RichText)
            x = MainMenu.rtbn.mRichTextBox.Text
            Printer.Line(0, 5800, 3750, 13050,, True)

            Dim printline As String, i As Integer, currenty As Integer = 6300
            Dim FirstVblf As Boolean, SecondVblf As Boolean, ThirdVblf As Boolean

            Printer.CurrentX = 100
            Printer.CurrentY = 6300

            For i = 1 To Len(x)
                Printer.FontName = "Arial Narrow"
                Printer.FontBold = True
                If i = 1 Then
                    printline = Mid(x, i, 32)
                    Printer.CurrentX = 100
                    Printer.CurrentY = currenty
                    Printer.FontUnderline = True
                    Printer.Print(printline)
                    Printer.FontUnderline = False
                    i = 32
                ElseIf i = 33 Then
                    printline = Mid(x, i, 43)
                    Printer.CurrentX = 100
                    Printer.CurrentY = currenty
                    Printer.FontUnderline = True
                    Printer.Print(printline)
                    Printer.FontUnderline = False
                    i = 76
                Else
                    printline = Mid(x, i, 60)
                    If printline.Contains(vbLf) Then
                        If FirstVblf = False Then
                            'Printer.CurrentX = 0
                            'Printer.CurrentY = currenty
                            'Printer.Print("")
                            printline = Mid(x, i, 34)
                            Printer.CurrentX = 100
                            Printer.CurrentY = currenty
                            Printer.FontUnderline = True
                            Printer.Print(printline)
                            Printer.FontUnderline = False

                            i = i + 34
                            'Printer.CurrentX = 0
                            'currenty = currenty + 200
                            'Printer.CurrentY = currenty
                            'Printer.Print("")
                            printline = Mid(x, i, 34)
                            Printer.CurrentX = 100
                            currenty = currenty + 200
                            Printer.CurrentY = currenty
                            Printer.FontUnderline = True
                            Printer.Print(printline)
                            Printer.FontUnderline = False

                            FirstVblf = True
                            i = i + 34
                        ElseIf SecondVblf = False Then
                            'i = i + 34
                            'Printer.CurrentX = 0
                            'Printer.CurrentY = currenty
                            'Printer.Print()
                            'printline = Mid(x, i, 24)
                            printline = Mid(x, i, 23)
                            Printer.CurrentX = 100
                            currenty = currenty + 200
                            Printer.CurrentY = currenty
                            Printer.FontUnderline = True
                            Printer.Print(printline)
                            Printer.FontUnderline = False
                            i = i + 24
                            'Printer.CurrentX = 0
                            'Printer.CurrentY = currenty
                            'Printer.Print()
                            printline = Mid(x, i, 32)
                            Printer.CurrentX = 100
                            currenty = currenty + 200
                            Printer.CurrentY = currenty
                            Printer.FontUnderline = True
                            Printer.Print(printline)
                            Printer.FontUnderline = False

                            SecondVblf = True
                            i = i + 31
                        ElseIf ThirdVblf = False Then
                            'i = i + 32
                            'Printer.CurrentX = 0
                            'Printer.CurrentY = currenty
                            'Printer.Print()
                            printline = Mid(x, i, 36)
                            'printline = Mid(x, i, 25)
                            Printer.CurrentX = 100
                            'currenty = currenty + 200
                            Printer.CurrentY = currenty
                            'Printer.FontUnderline = True
                            Printer.Print(printline)
                            'Printer.FontUnderline = False
                            i = i + 37
                            'Printer.CurrentX = 0
                            'Printer.CurrentY = currenty
                            'Printer.Print()
                            printline = Mid(x, i, 27)
                            Printer.CurrentX = 100
                            currenty = currenty + 400
                            Printer.CurrentY = currenty
                            'Printer.FontUnderline = True
                            Printer.Print(printline)
                            'Printer.FontUnderline = False
                            ThirdVblf = True
                            i = i + 28
                        End If
                        Printer.FontSize = 8
                        'If printline.Contains(vbLf) Then
                        '    Printer.CurrentX = 100
                        '    Printer.CurrentY = currenty
                        '    Printer.Print(Left(printline, InStr(printline, "&") - 1))

                        '    currenty = currenty + 200
                        '    Printer.CurrentX = 100
                        '    Printer.CurrentY = currenty
                        '    Printer.Print(Right(printline, InStr(printline, "&") - 1))
                    Else

                        Printer.CurrentX = 100
                        Printer.CurrentY = currenty
                        Printer.Print(printline & "-")
                        Printer.FontSize = 7.5

                        If i = 1288 Then
                            i = i + 13
                        ElseIf i = 1302 Then
                            i = i + 60
                            printline = Mid(x, i, 47)
                            Printer.CurrentX = 100
                            currenty = currenty + 200
                            Printer.CurrentY = currenty
                            Printer.FontUnderline = False
                            Printer.Print(printline & "-")
                            i = i + 49
                        Else
                            i = i + 59
                        End If

                    End If
                End If

                currenty = currenty + 200
                'If printline.Contains(vbLf) Then
                '    currenty = currenty + 400
                'Else
                '    currenty = currenty + 200
                'End If

            Next
            'Dim filetext As String = ReadFile(StorePolicyMessageFile, 1, CountFileLines(StorePolicyMessageFile))
            'Dim TotalLines As Integer
            'i = 1
            'TotalLines = CountFileLines(StorePolicyMessageFile, True)
            'Do While (i <= TotalLines)
            '    Printer.CurrentX = 100
            '    Printer.CurrentY = currenty
            '    Printer.Print(ReadFile(StorePolicyMessageFile, i, 1))
            '    i = i + 1
            '    currenty = currenty + 200
            'Loop

        End If
        '----------------
        '-----12/19
        'If CopyID = "File Copy" Then
        x = ""
        MainMenu.rtbn.mRichTextBox.LoadFile(CustomerTermsMessageFile)
        x = MainMenu.rtbn.mRichTextBox.Text
        Printer.FontBold = True
        Printer.CurrentX = 100
        Printer.CurrentY = 13200
        Printer.FontSize = 10
        Printer.Print(x)
        Printer.FontSize = 8.04
        '------------------

        Printer.FontBold = False
        If (Page + 1) = Pages Then
            Printer.CurrentX = 200 : Printer.CurrentY = 14000
            'Bal Due BOX
            Printer.DrawWidth = 8
            Printer.Line(9000, 14300, 11000, 14900, QBColor(0), True)
            Printer.DrawWidth = 1

            If Not IsUFO() Then
                Printer.CurrentX = 8600 : Printer.CurrentY = 13800
                'Printer.Print(TAB(120), "_______________________________")
                'Printer.Print("__________________________")
                Printer.Line(8600, 13800, 11000, 13800)
                Printer.CurrentX = 8650 : Printer.CurrentY = 13810
                'Printer.Print(TAB(120), "   Buyer's Approval")
                Printer.FontSize = 8
                Printer.Print("   Buyer's Approval")
            End If

            Printer.CurrentX = 9900 : Printer.CurrentY = 14100
            'Printer.Print(TAB(143), " Balance Due: ")
            Printer.FontSize = 8
            Printer.Print(" Balance Due: ")
        End If

        ' Printer.CurrentX = 200: Printer.CurrentY = 14750  '14900  14800 14750  ok
        ' Printer.CurrentX = 200: Printer.CurrentY = 13100  '14900  14800 14750  ok
        '.FontSize = 10

        'Printer_Location(8700, 13200, 10)  'Tab(130)
        'Printer.Print(TAB(100), " ", IfNullThenNilString(CopyID), "   ")
        Printer.FontSize = 10
        Printer.CurrentX = 8700 : Printer.CurrentY = 13200
        Printer.Print(IfNullThenNilString(CopyID))
        Printer.CurrentX = 10300 : Printer.CurrentY = 13200
        Printer.Print("Page " & Page + 1 & "/" & Pages, " ")

        Printer.FontSize = 8.04

        '************* Fill In Sale Info *****************************
        Printer_Location(200, 2400, 14)

        If ML.Index = 0 Then
            Printer.Print("CASH & CARRY")
        ElseIf Not ML.Business Then
            Printer.Print(IfNullThenNilString(ML.First), TAB(25), IfNullThenNilString(ML.Last))
        Else 'company
            Printer.Print(IfNullThenNilString(ML.Last))
        End If

        Printer_Location(200, 2950, 12, Trim(IfNullThenNilString(ML.Address)))

        Printer_Location(200, 3500, 12, Trim(IfNullThenNilString(ML.AddAddress)))

        Printer_Location(200, 4050, 12)
        Printer.Print(Trim(IfNullThenNilString(ML.City)), TAB(40), Trim(IfNullThenNilString(ML.Zip)))

        Printer_Location(200, 4550, 12)
        Printer.Print(DressAni(CleanAni(IfNullThenNilString(ML.Tele))), TAB(25), DressAni(CleanAni(IfNullThenNilString(ML.Tele2))))

        Printer_Location(6200, 2500, 12)
        Printer.Print(Trim(IfNullThenNilString(M2.ShipToFirst)), TAB(20), Trim(IfNullThenNilString(M2.ShipToLast)))

        Printer_Location(6200, 3000, 12, Trim(IfNullThenNilString(M2.Address2)))

        Printer_Location(6200, 3500, 12)
        Printer.Print(Trim(IfNullThenNilString(M2.City2)), TAB(96), Trim(IfNullThenNilString(M2.Zip2)))

        Printer_Location(6200, 4100, 12, DressAni(CleanAni(IfNullThenNilString(M2.Tele3))))

        Printer_Location(6200, 4400, 10, DressEmail(CleanEmail(IfNullThenNilString(ML.Email))))

        Printer.CurrentX = 150 : Printer.CurrentY = 6350 : W = 6350

        ' 17 items per page.
        For LoopRow = Page * 17 To (Page + 1) * 17 - 1
            Printer.FontSize = 9
            GetLinePart(Page, LoopRow - (Page * 17), Item, ItemLine)
            If Item < 0 Then
                MfgForm = ""
                StyleForm = ""
                DescForm = ""
                PriceForm = ""
                LocForm = ""
                StatusForm = ""
                Quanform = ""
            Else

                If CopyID = COPY_CUSTOMER Then
                    MfgForm = ""
                    StyleForm = ""
                    DescForm = QueryDesc(Item)
                    DescForm = Mid(DescForm, ItemLine * 46 + 1, 46)

                    If IsWoodPeckers() Then
                        StyleForm = QueryStyle(Item)
                    ElseIf IsStudioD() Then
                        If IsItem(QueryStyle(LoopRow)) Then
                            C = New CInvRec
                            If C.Load(QueryStyle(LoopRow), "Style") Then DescForm = C.SKU
                            DisposeDA(C)
                        End If
                    End If
                Else
                    StyleForm = QueryStyle(Item)
                    MfgForm = QueryMfg(Item)
                    DescForm = QueryDesc(Item)
                    DescForm = Mid(DescForm, ItemLine * 46 + 1, 46)
                End If
                PriceForm = CurrencyFormat(QueryPrice(Item))
                LocForm = QueryLoc(Item)
                StatusForm = QueryStatus(Item)
                Quanform = QueryQuan(Item)
            End If


            Dim ItemLineZeroDescForm As Boolean
            Dim ttCY As Integer
            ' 6 character status causes the line to shift down!
            If ItemLine = 0 Then
                ttCY = Printer.CurrentY
                Printer.CurrentY = ttCY : Printer.Print(IfNullThenNilString(StyleForm))
                Printer.CurrentY = ttCY : Printer.Print(TAB(37), Left(IfNullThenNilString(MfgForm), 15))
                On Error Resume Next
                If CopyID = COPY_CUSTOMER Then
                    Printer.CurrentY = ttCY : Printer.Print(TAB(97), IfZeroThenNilString(LocForm))
                    Printer.CurrentY = ttCY : Printer.Print(TAB(105), Left(IfNullThenNilString(StatusForm), 6))
                    Printer.CurrentY = ttCY : Printer.Print(TAB(130), IfNullThenNilString(Quanform))
                    Printer.CurrentY = ttCY : Printer.Print(TAB(137), IfNullThenNilString(DescForm))
                    ItemLineZeroDescForm = True
                Else
                    Printer.CurrentY = ttCY : Printer.Print(TAB(80), IfZeroThenNilString(LocForm))
                    Printer.CurrentY = ttCY : Printer.Print(TAB(87), Left(IfNullThenNilString(StatusForm), 6))
                    'Printer.CurrentY = ttCY : Printer.Print(TAB(97), IfNullThenNilString(Quanform))
                    Printer.CurrentY = ttCY : Printer.Print(TAB(110), IfNullThenNilString(Quanform))
                    Printer.CurrentY = ttCY : Printer.Print(TAB(117), IfNullThenNilString(DescForm))
                    ItemLineZeroDescForm = True
                End If
            End If
            If ItemLineZeroDescForm = False Then
                Printer.Print(TAB(71), IfNullThenNilString(DescForm))
            End If

            If ItemLine = 0 Then
                'allow over-write
                If StyleForm = "NOTES" And GetPrice(PriceForm) = 0 Then
                    '          Printer.Print
                ElseIf StyleForm <> "" Or GetPrice(PriceForm) > 0 Or GetPrice(PriceForm) < 0 Then ' discount
                    'PrintToPosition(Printer, PriceForm, 11350, AlignConstants.vbAlignRight, False)
                    Printer.CurrentY = ttCY
                    PrintToPosition(Printer, PriceForm, 11000, AlignConstants.vbAlignRight, False)
                Else
                End If
            End If

            'Printer.Print()
            Printer.CurrentX = 150
            Printer.CurrentY = Printer.CurrentY + 200
        Next

        ''-----12/19
        'Dim x As String
        'MainMenu.rtbn.mRichTextBox.LoadFile(CustomerTermsMessageFile)
        'x = MainMenu.rtbn.mRichTextBox.Text
        'Printer.CurrentX = 100
        'Printer.CurrentY = 13200
        'Printer.Print(x)
        ''----------------

        If IsUFO() Then 'Or IsFriendlys() Then
            If Holding.Status = "L" Or Holding.Status = "1" Or Holding.Status = "2" Or Holding.Status = "3" Or Holding.Status = "4" Then

                'reprint
                If Val(Holding.Status) = 1 Then Xx = 30
                If Val(Holding.Status) = 2 Then Xx = 60
                If Val(Holding.Status) = 3 Then Xx = 90
                If Val(Holding.Status) = 4 Then Xx = 120

                Printer.FontSize = 20
                Printer.FontBold = True
                Printer.CurrentY = 13500

                If IsUFO() Then
                    Printer.CurrentY = 12800 '12900 '13200
                    Printer.FontSize = 15
                    Printer.Print(TAB(2), " LAYAWAY PAYMENTS MUST BE MADE")
                    Printer.Print(TAB(2), " EVERY 2 WEEKS!  ", Xx, " DAY LAYAWAY")
                    Printer.Print(TAB(2), " Merchandise received in good condition!")
                    Printer.FontSize = 10
                    Printer.Print()
                    Printer.Print(TAB(2), "   Agreed: _________________________________________")
                    Printer.Print(TAB(12), " I accept the UFO Furniture Warehouse policies.")
                End If
                '        If IsFriendlys() Then
                '          Printer.Print Tab(2); " LAYAWAY PAYMENTS MUST BE MADE"
                '          Printer.Print Tab(2); " EVERY MONTH!  "; Xx; " DAY LAYAWAY"
                '          Printer.Print Tab(2); " Agreed: _________________________"
                '        End If

                'Unload LaAwaySelect
                LaAwaySelect.Close()
            End If
        End If

        Printer.CurrentX = 9000
        Printer.CurrentY = 14370
        Printer.FontSize = 20
        Printer.FontBold = True

        Printer.FontSize = 20
        If (Page + 1) = Pages Then
            Printer.FontBold = True
            If CopyID = COPY_CUSTOMER And Holding.Status = "F" Then
                Printer.Print(TAB(49), AlignString(CurrencyFormat(0), 9, AlignConstants.vbAlignRight))
            Else
                'Printer.Print(TAB(49), AlignString(CurrencyFormat(IfNullThenZeroCurrency(Holding.Sale - Holding.Deposit)), 9, AlignConstants.vbAlignRight))
                Printer.Print(AlignString(CurrencyFormat(IfNullThenZeroCurrency(Holding.Sale - Holding.Deposit)), 9, AlignConstants.vbAlignRight))
            End If
            Printer.FontBold = False
        End If

        Printer.FontBold = False

        If CopyID = COPY_CUSTOMER Then
            ' Needs to fit in about 3500 wide, fits in 6500, not in 6000.
            'MainMenu.rtbn.DoPrintFile(StorePolicyMessageFile, 100, 6300, 3500, 7000, True, False)
        End If

        ' Where does the RTB need to stop?
        '    Printer.Line (0, 13100)-Step(8500, 1800), QBColor(0), B
        '    If IsUFO() Or IsFriendlys() And
        If IsUFO() And (Holding.Status = "L" Or Holding.Status = "1" Or Holding.Status = "2" Or Holding.Status = "3" Or Holding.Status = "4") Then
            ' Don't print the customer terms box.
        Else
            'MainMenu.rtbn.DoPrintFile(CustomerTermsMessageFile, 100, 13200, 8300, 14000, True)

            'Dim x As String = MainMenu.rtbn.mRichTextBox.Text
            'Printer.CurrentX = 100
            'Printer.CurrentY = 13200
            'Printer.Print(x)
        End If
        Printer.EndDoc()

        DisposeDA(ML)
        Exit Sub

HandleErr:
        If Not CheckStandardErrors("Print Invoice") Then
            MessageBox.Show("ERROR in PrintInvoiceCommon: " & Err.Description & ", " & Err.Source & ", " & " Error NO: " & Err.Number)
            Resume Next
        End If
        Exit Sub ' no printer error exits.
    End Sub

    Private Sub GetLinePart(ByVal Page As Integer, ByVal Line As Integer, ByRef Item As Integer, ByRef ItemLine As Integer)
        Dim T As Integer, U As Integer, N As Integer, X As Integer, P As Integer, F As Integer

        X = 0
        For T = 1 To ItemCount
            N = ((Len(QueryDesc(T)) - 1) \ 46 + 1)
            For U = 1 To N
                X = X + 1

                P = (X \ 17)
                F = (X Mod 17)
                If Page = P And F > Line Then
                    Item = T
                    ItemLine = U - 1
                    Exit Sub
                ElseIf Page = P - 1 And Line = 16 And F = 0 Then
                    Item = T
                    ItemLine = U - 1
                    Exit Sub
                End If
            Next
        Next
        Item = -1
        ItemLine = 0
        Exit Sub
        MessageBox.Show("Could not Match Item Line.  Page " & (Page + 1) & ",'Line " & (Line + 1), "Invoice Printing Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Public Function QueryPrice(ByVal Index As Integer) As String
        QueryPrice = Items(Index - 1).Price
    End Function

    Public Function QueryStatus(ByVal Index As Integer) As String
        QueryStatus = Items(Index - 1).Status
    End Function

    Public Function QueryQuan(ByVal Index As Integer) As String
        QueryQuan = Items(Index - 1).Quantity
    End Function
End Class