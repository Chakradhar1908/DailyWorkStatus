Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class MailCheck
    Public CustomerTele As String
    Public NameFound As String
    Public Customer As String
    Public MarginNo As Integer
    Public HidePriorSales As Boolean                         ' Option for showing prior sales numbers.
    Dim RecNo(MaxLines) As Object
    Public SaleNo As String
    Private Const FRM_W1 As Integer = 8100
    Private Const FRM_W2 As Integer = 3735
    Private Const FRM_H1 As Integer = 3360
    Dim CustomerLast As String
    Public ServiceCallNo As Integer
    Dim tMail As MailNew
    Public Event CustomerFound(ByVal MailIndex As Integer, ByRef Cancel As Boolean)
    Public Event CustomerNotFound(ByRef Ignore As Boolean, ByRef DoUnload As Boolean)
    Public Event SaleNotFound()
    Public Event SaleFound(ByVal Holding As cHolding, ByRef Cancel As Boolean)
    Public Index As String
    Public ReturnHolding As Boolean                          ' Option for customer lookups.
    Public OrigStatus As String
    Public GrossSale As Decimal
    Public CheckOut As String
    Public OldTele As String
    Public FirstRec As Integer
    Public LastRec As Integer
    Public X As Integer                         ' Wacky link to bos2 grid line.
    Dim Margin As CGrossMargin
    Public Lease As String
    Public SalesPerson As String
    Public ArCashSls As Decimal
    Public Controll As Decimal
    Public TaxCode As Integer
    Public SpecialIns As String
    Public OriginalPrint As String

    Public Sub GetMarginLine()
        MarginNo = RecNo(BillOSale.X)
    End Sub
    Private Sub SetOrigSize()
        Width = FRM_W1
        Height = FRM_H1
    End Sub

    Public Sub LookUpCustomer(ByVal Inp As String, ByRef SelectFromBox As Boolean, Optional ByVal mLine As String = "")
        ' If SelectFromBox, Inp==MailIndex.
        ' Else check options.
        Customer = ""
        CustomerTele = Trim(Inp)   ' I don't like this being global, but changing the whole program would hold up the release.

        SaleNo = "" ' Something isn't cleaning up!

        Dim RS As ADODB.Recordset
        Dim tHold As cHolding
        Dim FailMsg As String

        SetOrigSize()

        If SelectFromBox Then
            If Not IsNumeric(CustomerTele) Then
                MsgBox("Invalid selection. Please try again.", vbCritical, "Warning!")
                Exit Sub
            End If
            If OrderMode("E") = True And optName.Checked = True And mLine <> "" And Microsoft.VisualBasic.Left(mLine, 3) = "   " Then
                Dim SN As String, T As Integer
                mLine = Trim(mLine)
                T = InStr(mLine, " ")
                SN = Trim(Microsoft.VisualBasic.Left(mLine, T))

                tHold = New cHolding
                tHold.DataAccess.DataBase = GetDatabaseAtLocation()

                If tHold.Load(SN, "LeaseNo") Then
                    RS = GetMailRecordset(tHold.Index)
                    SaleNo = tHold.LeaseNo
                Else
                    DisposeDA(tHold, RS)
                    RS = GetMailRecordset("-1")
                End If
                FailMsg = "Can't locate a sale by this number."
            End If

            If Not RS Is Nothing Then
                If RS.EOF Then RS = Nothing
            End If
            If RS Is Nothing Then
                RS = GetMailRecordset(CustomerTele)  ' Load customer by MailIndex.
            End If
        Else
            If Trim(CustomerTele) = "" Then
                FailMsg = "Please enter something in the box."
                RS = GetMailRecordset("-1")
            Else
                If optTelephone.Checked = True Then
                    ' Look up by phone number.
                    RS = GetMailRecordsetByTele(CleanAni(CustomerTele))
                    FailMsg = "Can't match telephone number."
                ElseIf optName.Checked = True Then
                    ' Look up by Name - impossible!
                    CustomerLast = Inp
                    RS = GetMailRecordset("-1")
                    '        If MsgBox("Is this an existing customer?  If so, please select from the list.", vbExclamation + vbYesNo, "Warning") = vbNo Then
                    '        Else
                    '          FailMsg = "Can't match customer's last name."
                    '          Exit Sub
                    '        End If
                ElseIf optSaleNo.Checked = True Then
                    ' Look up by SaleNo
                    tHold = New cHolding
                    tHold.DataAccess.DataBase = GetDatabaseAtLocation()

                    If tHold.Load(CustomerTele, "LeaseNo") Then
                        RS = GetMailRecordset(tHold.Index)
                        SaleNo = tHold.LeaseNo
                    Else
                        DisposeDA(tHold, RS)
                        RS = GetMailRecordset("-1")
                    End If
                    FailMsg = "Can't locate a sale by this number."
                ElseIf optServiceCall.Checked = True Then
                    ' Look up by Service Call Number
                    RS = getMailRecordsetByServiceCall(Inp)
                    ServiceCallNo = Val(Inp)
                    FailMsg = "There is no service call with this number."
                Else
                    ' No option selected?
                    FailMsg = "Invalid lookup options."
                    DisposeDA(tHold, RS)
                    Exit Sub
                End If
            End If
        End If

        ' We have the Mail Index if we're going to get it, and may have a Sale Number.
        Dim Cancel As Boolean, C2 As Boolean

        ' In terms of execution time, .EOF is ALWAYS preferable to .RecordCount.
        If Not RS.EOF Then
            CopyMailRecordsetToMailNew(RS, tMail)
            Customer = "Old"
            RaiseEvent CustomerFound(CLng(tMail.Index), Cancel)
            If Cancel Then
                DisposeDA(tHold, RS)
                'Unload Me
                Me.Close()
                Exit Sub
            End If
            GetCust()     ' Load mail data into BillOSale.
        Else
            'Does Not Find Customer in mailing list
            If OrderMode("A") Or MailMode("ADD/Edit") Or ArMode("S", "A") Then
                ' If we're allowed to create a new account here, check that we want to.
                If InputBox.Text <> "" Then
                    If MsgBox("Name Not In Data Base:  Try Again?", vbYesNo + vbExclamation, "Name Not Found") = vbYes Then
                        DisposeDA(tHold, RS)
                        Exit Sub
                    End If
                End If
                Customer = "New"
                NameFound = ""
                modMail.MailRec = 0
                BillOSale.Index = 0
                BillOSale.MailRec = 0
                BillOSale.Index = 0
                BillOSale.cmdApplyBillOSale.Enabled = True
            ElseIf Trim(SaleNo) <> "" Then
                ' Cash & Carry customer
                tMail.First = ""
                tMail.Last = "CASH & CARRY"
                tMail.Address = ""
                tMail.AddAddress = ""
                tMail.City = ""
                tMail.Zip = ""
                tMail.Tele = ""
                tMail.Tele2 = ""
                ' BillOSale.Sales1 = MARGIN.Salesman  ' Too bad we don't have Margin here.
                tMail.Index = -1
                Index = ""
                tMail.Special = ""
                tMail.Type = ""
                tMail.CustType = ""
                tMail.CreditCard = ""
                tMail.ExpDate = ""
                tMail.Email = ""
                tMail.TaxZone = 0

                GetCust()
            Else
                ' If we require an existing account, fail.
                Cancel = False : C2 = False
                RaiseEvent CustomerNotFound(Cancel, C2)
                If Not Cancel Then
                    If Trim(FailMsg) <> "" Then MsgBox(FailMsg, vbExclamation, "Warning")
                    If Visible Then
                        InputBox.Select()
                        SelectContents(InputBox)
                    End If
                End If
                If C2 Then
                    'Unload Me
                    Me.Close()
                End If
                DisposeDA(tHold, RS)
                Exit Sub
            End If
        End If

        ' ***** The following forms only require customer information.

        If OrderMode("A", "S") Or MailMode("ADD/Edit") Or ArMode("S", "A") Then
            BillOSale.DeleteLine = ""

            If OrderMode("S") Then  'service
                ' This should be the only code we need here.
                DisposeDA(tHold, RS)
                'Unload Me
                Me.Close()
                Service.LoadCustomer(CLng(tMail.Index), True)
                ServiceCallNo = 0  ' The selected item is only the default once.
                MainMenu.Hide()
                Service.Show()
                Exit Sub
            End If

            If Customer <> "Old" Then
                If optTelephone.Checked = True Then
                    BillOSale.CustomerPhone1 = DressAni(CleanAni(InputBox.Text))
                ElseIf optName.Checked = True Then
                    BillOSale.CustomerLast.Text = Trim(InputBox.Text)
                End If
            End If

            DisposeDA(tHold, RS)
            'Unload Me
            Me.Close()

            Exit Sub
        End If

        ' ***** Everything else needs a specific customer and sale number.
        'If Customer = "" Then Exit Sub  ' This breaks Cash&Carry.

        ' If we can't locate a sale, fail.
        If tHold Is Nothing Then
            SaleNo = AddOnAcc.GetSaleNumber(tMail.Index, BillOSale)
            'Unload AddOnAcc
            AddOnAcc.Close()
            AddOnAcc = Nothing
            tHold = New cHolding
            tHold.DataAccess.DataBase = GetDatabaseAtLocation()

            If Not tHold.Load(CStr(SaleNo), "LeaseNo") Then
                DisposeDA(tHold, RS)
                RaiseEvent SaleNotFound()
                MsgBox("Can't locate a sale for this customer.", vbExclamation, "Error")
                Exit Sub
            Else
                ' We've just located a sale.
            End If
        Else
            ' We already had a sale.
        End If

        Cancel = False
        RaiseEvent SaleFound(tHold, Cancel)
        If Cancel Then
            ' Form is already unloaded.. fix that. :)
            '    Me.Show
            '    InputBox.SetFocus
            '    InputBox.SelStart = 0
            '    InputBox.SelLength = Len(InputBox.Text)
            DisposeDA(tHold, RS)
            Exit Sub
        End If

        g_Holding = tHold  ' Nasty evil global holding object! Kill it!

        If ReturnHolding Then
            ' This is the future of this function..
            ' Mailcheck will become a lookup form, and not set -anything- in other forms.
            ' It will instead return data, and the other forms will set their own fields.
            ' Returning data to the calling function does no good.  I have to *ack* use
            ' the global holding object, or come up with something similarly evil.
            DisposeDA(tHold, RS)
            ReturnHolding = False
            Exit Sub
        End If

        OrigStatus = tHold.Status
        BillOSale.BalDue.Text = CurrencyFormat(tHold.Sale - tHold.Deposit)

        'BFH20150410 Some of the adjusted sales were WRONG in the Holding table.
        ' this would have 'fixed' it, but it created a 'currency' off.  Further, adjusting
        ' this way changed the cash flow, which threw their accounting way off.
        ' This was discontinued, and we will simpy allow these sales to be delivered 'as-is',
        ' with a 'hidden' balance due remaining.  Hopefully, we can track down the error in the
        ' adjustments.

        '  Dim SD As sSale, Amt as decimal
        '  Set SD = GetSaleData(tHold.LeaseNo)
        '  If tHold.Status <> "V" Then
        '    If tHold.Sale <> SD.SubTotal("gross") Or tHold.Deposit <> SD.SubTotal("paid") Then
        '      MsgBox "Error in sale balance in holding table." & vbCrLf & "Adjusting balance due.", vbInformation, "Balance error"
        '      ExecuteRecordsetBySQL "UPDATE [Holding] SET [Sale]=" & SQLCurrency(SD.SubTotal("gross")) & ", [Deposit]=" & SQLCurrency(SD.SubTotal("paid")) & " WHERE [LeaseNo]='" & tHold.LeaseNo & "'", , GetDatabaseAtLocation()
        '      Amt = SD.SubTotal("gross") - tHold.Sale
        '  '    AddNewCashJournalRecord "10200", Amt, tHold.LeaseNo, "Automatic Adjustment", Date, "SYSTEM"
        '      AddNewAuditRecord tHold.LeaseNo, SD.Name, Date, Amt, 0, Amt, 0, Amt, 0, 0, 0, SD.SalesCode, 0, "SYSTEM"
        '      Debug.Print "Holding Error on Sale display:"
        '      Debug.Print "Adjustments:  Sale(" & tHold.Sale & " -> " & SD.SubTotal("gross") & "), Deposit(" & tHold.Deposit & " -> " & SD.SubTotal("paid") & ")"
        '      Debug.Print "CASH ENTRY:   " & CurrencyFormat(Amt)
        '      tHold.Sale = SD.SubTotal("written")
        '      tHold.Deposit = SD.SubTotal("paid")
        '      BillOSale.BalDue = CurrencyFormat(SD.SubTotal())
        '    End If
        '  End If
        '  Set SD = Nothing

        BillOSale.SetLeaseNo(tHold.LeaseNo)
        GrossSale = tHold.Sale
        GetStatus()
        GetOrder(True)

        If OrderMode("B") Then ' Deliver Sale
            ' can't del a voided sale
            Hide()

            If Trim(tHold.Status) = "V" Then
                MsgBox("You Cannot Deliver A Voided Sale!", vbExclamation)
                DisposeDA(tHold, RS)
                Exit Sub
            End If
            BillOSale.cmdApplyBillOSale.Enabled = True
            On Error Resume Next
            'Load InvDel

            InvDel.ShowModal(BillOSale)
            '    InvDel.Show 'vbModal, BillOSale
            DisposeDA(tHold, RS)
            'Unload Me
            Me.Close()
            Exit Sub
        End If

        If OrderMode("C") Then ' Void Sale
            BillOSale.Refresh()
            Hide()

            If Not tHold.Void Then
                MsgBox("The sale and/or PO could not be voided.", vbCritical, "Warning")
            End If

            If MsgBox("Any More Sales To Void?", vbYesNo + vbQuestion) = vbYes Then
                'Unload BillOSale
                BillOSale.Close()
                'Unload Me
                Me.Close()
                BillOSale.Show()
                BillOSale.BillOSale2_Show()
                optSaleNo.Checked = True
                'Me.Show vbModal, BillOSale
                Me.ShowDialog(BillOSale)
                DisposeDA(tHold, RS)
                Exit Sub
            End If

            DisposeDA(tHold, RS)
            'Unload Me
            Me.Close()
            'Unload BillOSale
            BillOSale.Close()
            MainMenu.Show()
            Exit Sub
        End If

        If OrderMode("D") Then
            ' Payment on account
            If BillOSale.SaleStatus.Text = "Void" Then
                DisposeDA(tHold, RS)
                MsgBox("This sale is voided.  You cannot make any payments to it.", vbExclamation)
                Exit Sub
            End If

            If BillOSale.SaleStatus.Text = SALE_STATUS_FINANCE Or BillOSale.SaleStatus.Text = SALE_STATUS_OPENFINANCE Then
                MsgBox("This is an Installment Sale.  Only COD payments, set up on the contract, should be made here.", vbExclamation)
            End If

            If CheckOut = "Cancel" Then
                DisposeDA(tHold, RS)
                Exit Sub
            End If
            'Load OrdPay
            'OrdPay.HelpContextID = 44000
            OrdPay.Show()
        End If

        If OrderMode("E") Then
            BillOSale.cmdApplyBillOSale.Enabled = True
        End If

        DisposeDA(tHold, RS)

        '  If Order = "Credit" Then
        '    ' Customer Adjustment
        '    If tHold.Status = "V" Then
        '      ' Order is void and can't be changed.
        '      MsgBox "This sale is voided and cannot be changed!", vbCritical
        '      Exit Sub
        '    ElseIf Trim(tHold.Status) = "D" Or Trim(tHold.Status) = "B" Or Trim(tHold.Status) = "C" Then
        '      'Delivered, credit or backordered
        '      MsgBox "This sale is delivered and must be voided to change!", vbCritical
        '      Exit Sub
        '    Else
        '      Unload Me
        '    End If
        '  End If
    End Sub
    Public Sub GetCust()
        'FINDS OLD CUSTOMER & CONTINUES ON
        ' (Loads customer info into BillOSale from Mail object)
        If OrderMode("S") Then Exit Sub

        BillOSale.CustomerFirst.Text = tMail.First
        BillOSale.CustomerLast.Text = tMail.Last
        BillOSale.SetBusiness(tMail.Business)
        BillOSale.CustomerAddress.Text = tMail.Address
        BillOSale.AddAddress.Text = tMail.AddAddress
        BillOSale.CustomerCity.Text = tMail.City
        BillOSale.CustomerZip.Text = tMail.Zip
        BillOSale.CustomerPhone1.Text = DressAni(CleanAni(tMail.Tele))
        BillOSale.cboPhone1.Text = tMail.PhoneLabel1
        OldTele = tMail.Tele

        BillOSale.CustomerPhone2 = DressAni(CleanAni(tMail.Tele2))
        BillOSale.cboPhone2.Text = tMail.PhoneLabel2
        BillOSale.txtSpecInst.Text = tMail.Special
        On Error Resume Next
        If IsIn(tMail.Type, "0", "1", "2", "3") Then
            BillOSale.cboCustType.SelectedIndex = Val(tMail.Type)
        Else
            tMail.Type = 0 'screw up to allow for '-' instead of 0
        End If

        'BFH20171222 - No longer forcing new adv selection each time...
        '   If Not OrderMode("A") Then  'added 11-11-04 for force advertising selection each time
        SelectComboBoxItemData(BillOSale.cboAdvertisingType, Val(tMail.CustType))
        '        .cboAdvertisingType.ListIndex = Val(Mail.CustType)
        '    End If

        BillOSale.Index = Trim(tMail.Index)
        BillOSale.Email.Text = Trim(tMail.Email)
        BillOSale.Email.Tag = ""

        BillOSale.lblGrossSalesCaption.Text = CurrencyFormat(GetGrossSales(tMail.Index))


        If Val(tMail.TaxZone) <= BillOSale.cboTaxZone.Items.Count Then
            BillOSale.cboTaxZone.SelectedIndex = Val(tMail.TaxZone) - 1
        Else
            BillOSale.cboTaxZone.SelectedIndex = -1
        End If

        On Error GoTo 0

        If Trim(BillOSale.cboPhone1.Text = "") Then BillOSale.cboPhone1.Text = "Telephone"
        If Trim(BillOSale.cboPhone2.Text = "") Then BillOSale.cboPhone2.Text = "Telephone 2"
        GetCust2()

        If OrderMode("C") Then
            'ZOrder 1  ' Do nothing.. stoopid form misbehaving.
        Else
            'Unload Me
            Me.Close()
        End If
    End Sub
    Private Sub GetStatus()
        BillOSale.SaleStatus.Text = DescribeHoldingStatus(g_Holding.Status)
    End Sub
    Public Sub GetOrder(Optional ByVal SkipCustInfo As Boolean = False)
        Dim Sales As String, SalesSplit As String
        'to put on sales

        ' This function appears to be causing the slow loading of BillOSale.
        ' Is the problem in the generic database handler, or in the grid it's
        ' being loaded into?
        ' To witness the sluggishness:
        '   Freshly load VB.
        '   Run the project.
        '   Go to Order->View Sale
        '   Enter: Sale #33798.
        ' Subsequent searches will be fast.
        ' If the program has run in the IDE before, it will be fast.
        ' It's possible the databases aren't being cleaned up properly,
        ' which would explain the performance difference. That could also
        ' be Access or Windows caching information from the earlier
        ' connection objects.

        ' How can we determine where the slowness is happening?
        ' VB doesn't have good timing or profiling tools, so it's impossible
        ' to know how much time is spent in each function.

        ' This profiling effort is being put on hold until the major debugging
        ' effort is concluded.  At that point, we will download the free 7-day
        ' trial version of Rational Quantify and attempt to speed up the whole
        ' program.
        ' (http://www.rational.com/tryit/quantify_nt/index.jsp)

        ' Reset for next sale
        'BillOSale.UGridIO1.Clear()      'BFH20120825 - ADDED GRID CLEAR CUZ PAYMENT WAS KEEPING PREVIOUS SALE LINES

        MarginNo = 0
        FirstRec = 0

        LastRec = 0
        X = 0

        Dim cTable As New CGrossMargin
        cTable.DataAccess.DataBase = GetDatabaseAtLocation()
        Dim cTa As CDataAccess : cTa = cTable.DataAccess()
        cTa.Records_OpenIndexAt(Index:=Trim(CStr(SaleNo)), OrderBy:="MarginLine")
        If cTa.Record_Count > MaxLines Then
            MsgBox("ERROR!!" & vbCrLf2 & "This sale has " & cTa.Record_Count & " lines." & vbCrLf2 & "The maximum number of lines for a sale is " & MaxLines & "." & vbCrLf2 & "Please contact " & AdminContactName & " at " & AdminContactPhone & " immediately with how you created this sale." & vbCrLf2 & "NOTE:  The full sale is not displayed.", vbCritical, "Unable to view sale")
        End If
        Do While cTa.Records_Available()
            If X >= MaxLines + 1 Then Exit Do
            Margin = cTable

            MarginNo = Margin.MarginLine
            If FirstRec = 0 Then
                FirstRec = Margin.MarginLine
                On Error Resume Next
                BillOSale.dteSaleDate.Value = DateFormat(Margin.SellDte)
                On Error GoTo 0
                If IsDate(Margin.DDelDat) Then
                    BillOSale.lblDelDate.Text = DateFormat(Margin.DDelDat)
                    BillOSale.dteDelivery.Value = Margin.DDelDat
                Else
                    BillOSale.lblDelDate.Text = ""
                End If
                If IsDate(Margin.DDelDat) Then BillOSale.lblDelWeekday.Text = GetDay(CDate(Margin.DDelDat))
            End If

            RecNo(X) = Margin.MarginLine

            Lease = SaleNo
            BillOSale.BillOfSale.Text = Margin.SaleNo
            Sales = Margin.Salesman
            SalesSplit = Margin.SalesSplit

            SalesPerson = Margin.Salesman

            BillOSale.BillOSale2_Show()
            BillOSale.X = X
            BillOSale.SetStyle(X, Margin.Style)
            BillOSale.SetMfg(X, Margin.Vendor)
            BillOSale.SetLoc(X, Margin.Location)
            BillOSale.SetStatus(X, Margin.Status)
            BillOSale.SetQuan(X, CStr(Margin.Quantity))
            BillOSale.SetDesc(X, Margin.Desc)
            BillOSale.SetPrice(X, CurrencyFormat(Margin.SellPrice))

            If Margin.PorD = "D" Then
                BillOSale.chkDelivery.Text = 1
                BillOSale.dtpDelWindow.Value = Margin.StopStart
                BillOSale.dtpDelWindow2.Value = Margin.StopEnd
            ElseIf Margin.PorD = "P" Then
                BillOSale.chkPickup.Checked = 1
                BillOSale.dtpDelWindow.Value = Margin.StopStart
                BillOSale.dtpDelWindow2.Value = Margin.StopEnd
            End If

            If Trim(Margin.Style) = "SUB" Then
                If Trim(BillOSale.QueryPrice(X)) = "" Then BillOSale.SetPrice(X, "0")
                ArCashSls@ = BillOSale.QueryPrice(X)
            End If
            If Trim(Margin.Style) = "PAYMENT" Then
                If Trim(BillOSale.QueryPrice(X)) = "" Then BillOSale.SetPrice(X, "0")
                Controll@ = Controll@ + BillOSale.QueryPrice(X)
            End If

            If Trim(Margin.Style) = "TAX1" Or Trim(Margin.Style) = "TAX2" Then
                TaxCode = IIf(Margin.Quantity = 0, 1, Margin.Quantity)
            End If


            LastRec = MarginNo
            X = X + 1
        Loop
        cTa.Records_Close()
        BillOSale.LastRecord = X

        BillOSale.GridRefresh()
        ' bfh20050816 - this was confusing..  put it at the bottom of large sales instead of the top..
        ' made it easy to not notice the slider was down on the grid a ways and 'lose' items
        '  BillOSale.GridMove X

        BillOSale.Sales1.Text = Sales
        If Not OrderMode("A") Then
            GetSales()
            BillOSale.LoadSplitsToBoxes(SalesSplit, IIf(BillOSale.Sales2.Text = "", 1, IIf(BillOSale.Sales3.Text = "", 2, 3)))
        End If

        If Not SkipCustInfo Then
            FindCust()
            GetCust()
        End If

        BillOSale.HandleRecentNotes()

        Exit Sub

HandleErr:
        MsgBox("ERROR in MailCheck.GetOrder: " & Err.Description & ", " & Err.Source)
        If Err.Number = 13 Then Resume Next
        Resume Next
    End Sub
    Public Sub GetCust2()
        On Error GoTo HandleErr
        If Not IsNumeric(tMail.Index) Then
            MsgBox("Can't open mail record.", vbCritical, "Error")
            Exit Sub
        End If
        Dim RS2 As ADODB.Recordset
        RS2 = getRecordsetByTableLabelIndexNumber("MailShipTo", "Index", CStr(tMail.Index))
        If (Not RS2.EOF) Then
            Dim Mail2 As MailNew2
            CopyMailRecordsetToMailNew2(RS2, Mail2)
            BillOSale.ShipToLast.Text = Mail2.ShipToLast
            BillOSale.ShipToFirst.Text = Mail2.ShipToFirst
            BillOSale.CustomerAddress2.Text = Mail2.Address2
            BillOSale.CustomerCity2.Text = Mail2.City2
            BillOSale.CustomerZip2.Text = Mail2.Zip2
            BillOSale.CustomerPhone3 = DressAni(CleanAni(Mail2.Tele3))
            BillOSale.cboPhone3.Text = Mail2.PhoneLabel3
        Else
            BillOSale.ShipToLast.Text = ""
            BillOSale.ShipToFirst.Text = ""
            BillOSale.CustomerAddress2.Text = ""
            BillOSale.CustomerCity2.Text = ""
            BillOSale.CustomerZip2.Text = ""
            BillOSale.CustomerPhone3.Text = ""
            BillOSale.cboPhone3.Text = ""
        End If
        If Trim(BillOSale.cboPhone3.Text = "") Then BillOSale.cboPhone3.Text = "Telephone 3"
        RS2.Close()
        RS2 = Nothing
        Exit Sub

        ' Doesn't find second address causes an error
HandleErr:
        MsgBox("ERROR in GetCust2 " & Err.Description & ", " & Err.Source)
        If Err.Number = 13 Then Resume Next
    End Sub
    Private Sub GetSales()
        Dim SalesArr() As String
        SalesArr = Split(Trim(BillOSale.Sales1.Text), " ")
        If UBound(SalesArr) >= 0 Then BillOSale.Sales1.Text = Trim(getSalesName(SalesArr(0))) Else BillOSale.Sales1.Text = ""
        If UBound(SalesArr) >= 1 Then BillOSale.Sales2.Text = Trim(getSalesName(SalesArr(1))) Else BillOSale.Sales2.Text = ""
        If UBound(SalesArr) >= 2 Then BillOSale.Sales3.Text = Trim(getSalesName(SalesArr(2))) Else BillOSale.Sales3.Text = ""
    End Sub
    Public Sub FindCust()
        On Error GoTo HandleErr
        Dim RS As ADODB.Recordset

        If Val(Index) > 0 Then
            RS = GetMailRecordset(Trim(Index))
            CopyMailRecordsetToMailNew(RS, tMail)

            If OrderMode("A") Then
                If Trim(tMail.Index) = Trim(CustomerTele) Then
                    GetCust()
                    Customer = "OLD"
                End If
            End If

            If OrderMode("S") Then
                Service.lblFirstName.Text = Trim(tMail.First)
                Service.lblLastName.Text = Trim(tMail.Last)
                Service.lblAddress.Text = Trim(tMail.Address)
                Service.lblAddress2.Text = Trim(tMail.AddAddress)
                Service.lblCity.Text = Trim(tMail.City)
                Service.lblZip.Text = tMail.Zip
                Service.lblTele.Text = tMail.Tele
                Service.lblTele2.Text = tMail.Tele2
                SpecialIns = tMail.Special
                Exit Sub
            End If
        End If

        If Val(Index) <= 0 Then
            ' Doesn't find customer
            tMail.First = ""
            tMail.Last = "CASH & CARRY"
            tMail.Address = ""
            tMail.City = ""
            tMail.Zip = ""
            tMail.Tele = ""
            tMail.Tele2 = ""
            tMail.Index = -1
            Index = ""
            tMail.Special = ""
            tMail.Type = ""
            tMail.CustType = ""
        End If


        Customer = "OLD"
        If OrderMode("A") Then
            If MsgBox("Not In Data Base:  Try Again?", vbYesNo) = vbYes Then
                InputBox.Select()
                Exit Sub
            End If
        End If
        Exit Sub

HandleErr:
        ' MsgBox "ERROR in MailCheck.Findcust: " & Err.Description & ", " & Err.Source
        Resume Next
    End Sub

End Class