Imports ADODB
Imports VBRUN
Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class MailCheck
    Public CustomerTele As String
    Public NameFound As String
    Public Customer As String
    Public MarginNo As Integer
    Public HidePriorSales As Boolean                         ' Option for showing prior sales numbers.
    Dim RecNo(MaxLines) As Object
    Public SaleNo As String
    'Private Const FRM_W1 As Integer = 8100
    Private Const FRM_W1 As Integer = 550
    'Private Const FRM_W2 As Integer = 3735
    'Private Const FRM_W2 As Integer = 545
    Private Const FRM_W2 As Integer = 250
    'Private Const FRM_H1 As Integer = 3360
    Private Const FRM_H1 As Integer = 198

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
    Private WithEvents objSourceTelephone As CDbTypeAhead
    Private WithEvents objSourceName As CDbTypeAhead
    Private WithEvents objSource As CDbTypeAhead
    Public Event Cancelled(ByRef PreventUnload As Boolean, ByRef PreventMainMenu As Boolean)
    Private LastIndex As String                              ' For printing BillOSale Numbers.
    Private MailCheckClose As Boolean

    'Public ItemdataValue As Object
    'Public SelectedItemValue As String

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
                MessageBox.Show("Invalid selection. Please try again.", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
                    If MessageBox.Show("Name Not In Data Base:  Try Again?", "Name Not Found", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
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
                    If Trim(FailMsg) <> "" Then MessageBox.Show(FailMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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
                    BillOSale.CustomerPhone1.Text = DressAni(CleanAni(InputBox.Text))
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
                MessageBox.Show("Can't locate a sale for this customer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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
                MessageBox.Show("You Cannot Deliver A Voided Sale!", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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
                MessageBox.Show("The sale and/or PO could not be voided.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            If MessageBox.Show("Any More Sales To Void?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
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
                MessageBox.Show("This sale is voided.  You cannot make any payments to it.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            If BillOSale.SaleStatus.Text = SALE_STATUS_FINANCE Or BillOSale.SaleStatus.Text = SALE_STATUS_OPENFINANCE Then
                MessageBox.Show("This is an Installment Sale. Only COD payments, set up on the contract, should be made here.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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

        BillOSale.CustomerPhone2.Text = DressAni(CleanAni(tMail.Tele2))
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

        BillOSale.lblGrossSales.Text = CurrencyFormat(GetGrossSales(tMail.Index))


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
            'MailCheckClose = True
            'Me.Close()
            Me.Hide()
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
        BillOSale.UGridIO1.Clear()      'BFH20120825 - ADDED GRID CLEAR CUZ PAYMENT WAS KEEPING PREVIOUS SALE LINES

        MarginNo = 0
        FirstRec = 0

        LastRec = 0
        X = 0

        Dim cTable As New CGrossMargin
        cTable.DataAccess.DataBase = GetDatabaseAtLocation()
        Dim cTa As CDataAccess : cTa = cTable.DataAccess()
        cTa.Records_OpenIndexAt(Index:=Trim(CStr(SaleNo)), OrderBy:="MarginLine")
        If cTa.Record_Count > MaxLines Then
            MessageBox.Show("ERROR!!" & vbCrLf2 & "This sale has " & cTa.Record_Count & " lines." & vbCrLf2 & "The maximum number of lines for a sale is " & MaxLines & "." & vbCrLf2 & "Please contact " & AdminContactName & " at " & AdminContactPhone & " immediately with how you created this sale." & vbCrLf2 & "NOTE:  The full sale is not displayed.", "Unable to view sale", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        Do While cTa.Records_Available()
            cTable.cDataAccess_GetRecordSet(cTa.RS)
            If X >= MaxLines + 1 Then Exit Do
            Margin = cTable

            MarginNo = Margin.MarginLine
            If FirstRec = 0 Then
                FirstRec = Margin.MarginLine
                On Error Resume Next
                'BillOSale.dteSaleDate.Value =DateFormat(Margin.SellDte)
                BillOSale.dteSaleDate.Value = Date.Parse(DateFormat(Margin.SellDte), Globalization.CultureInfo.InvariantCulture)
                On Error GoTo 0
                If IsDate(Margin.DDelDat) Then
                    BillOSale.lblDelDate.Text = DateFormat(Margin.DDelDat)
                    'BillOSale.dteDelivery.Value = Margin.DDelDat
                    BillOSale.dteDelivery.Value = Date.Parse(Margin.DDelDat, Globalization.CultureInfo.InvariantCulture)
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
        MessageBox.Show("ERROR in MailCheck.GetOrder: " & Err.Description & ", " & Err.Source)
        If Err.Number = 13 Then Resume Next
        Resume Next
    End Sub

    Public Sub GetCust2()
        On Error GoTo HandleErr
        If Not IsNumeric(tMail.Index) Then
            MessageBox.Show("Can't open mail record.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
            BillOSale.CustomerPhone3.Text = DressAni(CleanAni(Mail2.Tele3))
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
        MessageBox.Show("ERROR in GetCust2 " & Err.Description & ", " & Err.Source)
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
            If MessageBox.Show("Not In Data Base:  Try Again?", "", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                InputBox.Select()
                Exit Sub
            End If
        End If
        Exit Sub

HandleErr:
        ' MsgBox "ERROR in MailCheck.Findcust: " & Err.Description & ", " & Err.Source
        Resume Next
    End Sub

    Public Sub MailCheck_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Height = FRM_H1
        Width = FRM_W2


        'SetButtonImage(cmdOK)
        'SetButtonImage(cmdCancel)
        SetButtonImage(cmdOK, 2)
        SetButtonImage(cmdCancel, 3)
        '    SetCustomFrame Me, ncBasicDialog
        On Error Resume Next 'Supposed to prevent an error if no printer is installed
        'OriginalPrint = Printer.DeviceName

        objSourceTelephone = New_CDbTypeAhead _
        (Table:="Mail left join Holding on Mail.Index=Holding.Index" _
        , Field:="Tele" _
        , Value:=InputBox.Text _
        , Match:=1 _
        , MinLength:=6
        )
        objSourceName = New_CDbTypeAhead _
        (Table:="Mail LEFT JOIN [Holding] ON Mail.Index=Holding.Index" _
        , Field:="Last" _
        , Value:=InputBox.Text _
        , Match:=1 _
        , MinLength:=2 _
        , ExtraSort:="First, LeaseNo"
        )

        If MailCheckSaleNoChecked = False Then
            optTelephone.Checked = True
        ElseIf MailCheckSaleNoChecked = True Then
            optSaleNo.Checked = True
        End If
        Setup()
    End Sub

    Private Sub Setup()
        On Error Resume Next
        InputBox.Select()

        'Order = "A"     '------> It will be assigned in modMainMenu. Because modMainMenu code is not completed
        'temporarily assigned here to run the below select case Order code. After modMainMenu completed, 
        'remove this line Order = "A" from here.
        If OrderMode("A") Then
            Me.Text = "Check For Prior Sales"
        End If

        If OrderMode("S") Or ArMode("A") Then   'service / CreditApp
            Me.Text = "Find Customers"
            lblInput.Text = "Enter Telephone Number"
        End If

        If MailMode("ADD/Edit") Or ArMode("S") Then
            Me.Text = "Add & Edit Customers"
            BillOSale.Sales1.Enabled = False
            BillOSale.Sales2.Enabled = False
            BillOSale.Sales3.Enabled = False
            BillOSale.SalesSplit1.Visible = False
            BillOSale.SalesSplit2.Visible = False
            BillOSale.SalesSplit3.Visible = False
        End If

        If OrderMode("A") Or MailMode("ADD/Edit") Or ArMode("S") Then       ' new sale
            lblMatches.Text = "Dbl Click On Telephone Number To Insert"
            optSaleNo.Enabled = False
            'optTelephone.value = True
            Exit Sub
        End If

        If OrderMode("A") Or Not MailMode("ADD/Edit") Or ArMode("S") Then
            If (Not OrderMode("", "S")) Or (Not ArMode("", "A")) Then
                CheckOut = ""
                lblInput.Text = "Enter Sale Number"
                lblMatches.Text = "Dbl Click On Sale Number To Insert"
                BillOSale.cmdApplyBillOSale.Enabled = False
                BillOSale.cmdCancel.Enabled = False
            End If
        End If
        If MailMode("ADD/Edit") Or ArMode("S") Then
            lblInput.Text = "Enter Telephone Number"
            BillOSale.cmdApplyBillOSale.Enabled = True
            BillOSale.cmdCancel.Enabled = True
        End If

        If OrderMode("B") Then
            Text = "Enter Sale To Be Delivered"
            Exit Sub
        End If

        If OrderMode("C") Then
            Text = "Enter Sale To Be Voided"
            Exit Sub
        End If


        If OrderMode("D") Then
            Text = "Enter Sale For Payment On Account"
            Exit Sub
        End If

        If OrderMode("E") Then
            Text = "Enter Sale To View"
        End If

        If OrderMode("S") Then
            optServiceCall.Visible = True
        Else
            optServiceCall.Visible = False
        End If
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor
        LookUpCustomer(InputBox.Text, False)
        'MousePointer = vbNormal
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Dim Prev As Boolean, PrevMM As Boolean

        RaiseEvent Cancelled(Prev, PrevMM)
        If Prev Then Exit Sub

        If OrderMode("S") Then
            'Unload Service
            Service.Close()
        End If
        Hide()
        'Unload Me
        Me.Close()

        If PrevMM Then Exit Sub
        If OrderMode("Credit", "CashRegister") Then Exit Sub

        modProgramState.Order = ""
        modProgramState.Mail = ""
        modProgramState.ArSelect = ""
        'Unload BillOSale
        BillOSale.Close()
        MainMenu.Show()
    End Sub

    Private Sub MailCheck_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        'SetCustomFrame Me, ncBasicDialog -> This line is not required. It is to set font and colors using modNeoCaption module.
        On Error Resume Next
        If IsParkPlace Then optName.Checked = True
        InputBox.Select()
        ServiceCallNo = 0
    End Sub

    Private Sub MailCheck_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Private Sub MailCheck_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        'If MailCheckSaleNoChecked = False Then
        If Width < FRM_W2 Then Width = FRM_W2
        'End If
        'If ScaleWidth - lstMatches.Left - 120 > 0 Then lstMatches.Width = ScaleWidth - lstMatches.Left - 120
        'If Me.ClientSize.Width - lstMatches.Left - 120 > 0 Then lstMatches.Width = Me.ClientSize.Width - lstMatches.Left - 120
        If Me.ClientSize.Width - lstMatches.Left > 0 Then lstMatches.Width = Me.ClientSize.Width - lstMatches.Left - 10
    End Sub

    Private Sub MailCheck_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'This event is replacement for form unload and queryunload events of vb6.0

        'Queryunload code
        'If UnloadMode = vbFormControlMenu Then Cancel = True ' cmdCancel.value = True   ' Still having problems with BillOSale.
        'If MailCheckClose = False Then   'IMP NOTE: Added this If block to skip the next if block. Otherwise, e.cancel=true is executing and this form is not cloing.
        '    If e.CloseReason = CloseReason.UserClosing Then
        '        e.Cancel = True
        '    End If
        'End If

        'Unload code
        'RemoveCustomFrame Me -> This line is not required. It is to set font and colors of a form using modNeoCaption module.
        objSourceTelephone = Nothing
        objSourceName = Nothing
        objSource = Nothing
        Index = 0
        MailRec = 0
        CustomerTele = ""
        OldTele = ""
        ReturnHolding = False
    End Sub

    Private Sub InputBox_TextChanged(sender As Object, e As EventArgs) Handles InputBox.TextChanged
        On Error Resume Next
        If optSaleNo.Checked = True Then Exit Sub              ' Don't search by sale number.
        If optTelephone.Checked = True Then FormatAniTextBox(InputBox)
        'If Not (objSource Is Nothing) Then objSource.Refresh() -> Refresh is replaced with RefreshObject in the below line. Cause there is already public event with Refresh name in CDbTypeAhead. Two objects with same not accepted in vb.net.
        If Not (objSource Is Nothing) Then objSource.RefreshObject()
    End Sub

    Private Sub InputBox_Enter(sender As Object, e As EventArgs) Handles InputBox.Enter
        SelectContents(InputBox)
    End Sub

    Private Sub lstMatches_DoubleClick(sender As Object, e As EventArgs) Handles lstMatches.DoubleClick
        If lstMatches.SelectedIndex = -1 Then Exit Sub
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor
        'LookUpCustomer(lstMatches.itemData(lstMatches.ListIndex), True, lstMatches.List(lstMatches.ListIndex))
        'ItemdataValue = CType(lstMatches.SelectedItem, ItemDataClass).ItemData
        'SelectedItemValue = lstMatches.SelectedItem.ToString
        LookUpCustomer(CType(lstMatches.SelectedItem, ItemDataClass).ItemData, True, lstMatches.SelectedItem.ToString)
        'MousePointer = vbNormal
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub lstMatches_KeyDown(sender As Object, e As KeyEventArgs) Handles lstMatches.KeyDown
        'If KeyCode = 13 Then
        If e.KeyCode = Keys.Enter Then
            'lstMatches_DblClick ' bfh20051229
            lstMatches_DoubleClick(lstMatches, New EventArgs)
        End If
    End Sub

    Private Sub InputBox_KeyDown(sender As Object, e As KeyEventArgs) Handles InputBox.KeyDown
        Dim C As String
        'Debug.Print "InputBox_KeyDown KeyCode=" & KeyCode & ", Shift=" & Shift
        'C = Format(Shift, "000") & Format(KeyCode, "000")
        C = Format(e.Shift, "000") & Format(e.KeyCode, "000")
        'Debug.Print "InputBox_KeyDown " & C
        Select Case C
            Case "000013"    ' Return
                cmdOK_Click(cmdOK, New EventArgs)
            Case "002192"    ' Ctrl-ESC
                If Not IsDevelopment() Then Exit Sub
                On Error Resume Next
                optName.Checked = True
                InputBox.Text = "CARTER"
                lstMatches.SelectedIndex = 0
                'lstMatches_DblClick
                lstMatches_DoubleClick(lstMatches, New EventArgs)
                'BillOSale.cmdApplyBillOSale.Value = True
                BillOSale.cmdApplyBillOSale.PerformClick()
        End Select
    End Sub

    Private Sub optTelephone_Click(sender As Object, e As EventArgs) Handles optTelephone.Click
        If OrderMode("A") Then
            lblInput.Text = "Enter Telephone:"
            cmdOK.Enabled = True
        Else
            lblInput.Text = "Telephone:  Tab off Telephone Number"
            cmdOK.Enabled = True 'False
        End If
        If MailMode("ADD/Edit") Then ' Or armode("S") Then
            cmdOK.Enabled = True
        End If

        CustomerTele = InputBox.Text
        'Width = 3735
        'Width = 8000
        'Width = FRM_W2
        Width = 545
        FormatAniTextBox(InputBox)
        InputBox.TabStop = True
        InputBox.TabIndex = 1
        'Set objSource = objSourceTelephone:    objSource.mUpdated = False : objSource.Refresh() ->Refresh is replaced with RefreshObject in the below line.
        objSource = objSourceTelephone : objSource.mUpdated = False : objSource.RefreshObject()
        On Error Resume Next
        InputBox.Select()
    End Sub

    Private Sub optName_Click(sender As Object, e As EventArgs) Handles optName.Click
        'name
        Width = FRM_W1
        lblInput.Text = "Last Name:  Tab Off Name"
        CustomerLast = InputBox.Text
        'cmdOK.Enabled = False
        InputBox.TabStop = True
        InputBox.TabIndex = 1
        On Error Resume Next
        InputBox.Select()
        'set objSource = objSourceName : objSource.mUpdated = False : objSource.Refresh() -> Refresh is replaced with RefreshObject in the below line.
        objSource = objSourceName
        objSource.mUpdated = False
        objSource.RefreshObject()
    End Sub

    Private Sub optSaleNo_Click(sender As Object, e As EventArgs) Handles optSaleNo.Click
        'sale no
        Width = FRM_W2
        cmdOK.Enabled = True
        lblInput.Text = "Enter Sale Number"
        InputBox.Text = CleanAni(InputBox.Text)
        SaleNo = InputBox.Text
        InputBox.TabStop = True
        If Visible = True Then
            InputBox.Select()
        End If
        objSource = Nothing 'objSourceSaleNo:    objSource.mUpdated = False:  objSource.Refresh
    End Sub

    Private Sub optServiceCall_Click(sender As Object, e As EventArgs) Handles optServiceCall.Click
        On Error Resume Next
        objSource = Nothing
        Width = FRM_W2
        cmdOK.Enabled = True
        lblInput.Text = "Enter Service Call Number"
        InputBox.Text = CleanAni(InputBox.Text)
        SaleNo = InputBox.Text
        InputBox.TabStop = True
        InputBox.Select()
    End Sub

    Private Sub objSourceTelephone_BuildKeyedLine(RS As Recordset, ByRef returnLine As String, ByRef returnKey As Integer) Handles objSourceTelephone.BuildKeyedLine
        Dim Index As String
        Dim LastName As String
        Dim FirstName As String
        Dim Address As String
        Dim Telephone As String
        Dim LeaseNo As String = ""
        ' bfh20060117
        Dim HoldStatus As String = ""
        Dim Balance As Decimal

        Index = IfNullThenZero(RS("Mail.Index").Value)
        LastName = IfNullThenNilString(RS("Last").Value)
        FirstName = IfNullThenNilString(RS("First").Value)
        Address = IfNullThenNilString(RS("Address").Value)
        Telephone = IfNullThenNilString(RS("Tele").Value)

        On Error Resume Next  '-This error handler is for Leaseno, status, sale and deposit columns. Recordset is returning DBNull if these columns are blank in database. Ifnullthennilstring will not handle it.
        LeaseNo = IfNullThenNilString(RS("LeaseNo").Value)
        ' bfh20060117
        HoldStatus = IfNullThenNilString(RS("Status").Value)
        Balance = IfNullThenZeroCurrency(RS("Sale").Value) - IfNullThenZeroCurrency(RS("Deposit").Value)

        Dim Extra As String

        If HidePriorSales Then
            If LastIndex = Index Then
                ' do nothing
                returnLine = "(skip)"
            Else
                returnLine = DressAni(CleanAni(Trim(Telephone))) & vbTab & AlignString(Trim(LastName) & ",  " & Trim(FirstName), 25, AlignConstants.vbAlignLeft, False) & vbTab & Trim(Address)
                LastIndex = Index
            End If
        Else
            Extra = Space(3) & LeaseNo & Space(3) & HoldStatus & Space(3) & FormatCurrency(Balance)
            If LastIndex = Index Then
                returnLine = Extra
            Else
                If LeaseNo = "" Then
                    returnLine = DressAni(CleanAni(Trim(Telephone))) & vbTab & AlignString(Trim(LastName) & ",  " & Trim(FirstName), 25, AlignConstants.vbAlignLeft, False) & vbTab & Trim(Address)
                Else
                    returnLine = DressAni(CleanAni(Trim(Telephone))) & vbTab & AlignString(Trim(LastName) & ",  " & Trim(FirstName), 25, AlignConstants.vbAlignLeft, False) & vbTab & Trim(Address) & vbCrLf & Extra
                End If

                LastIndex = Index
            End If
        End If
        returnKey = Index
    End Sub

    Private Sub objSourceName_BuildKeyedLine(RS As Recordset, ByRef returnLine As String, ByRef returnKey As Integer) Handles objSourceName.BuildKeyedLine
        Dim Index As String : Index = IfNullThenZero(RS("Mail.Index").Value)
        Dim LastName As String : LastName = IfNullThenNilString(RS("Last").Value)
        Dim FirstName As String : FirstName = IfNullThenNilString(RS("First").Value)
        Dim Address As String : Address = IfNullThenNilString(RS("Address").Value)
        Dim Telephone As String : Telephone = IfNullThenNilString(RS("Tele").Value)

        On Error Resume Next '-This error handler is for Leaseno, status, sale and deposit columns. Recordset is returning DBNull if these columns are blank in database. Ifnullthennilstring will not handle it.
        Dim LeaseNo As String = "" : LeaseNo = IfNullThenNilString(RS("LeaseNo").Value)
        ' bfh20060117
        Dim HoldStatus As String = "" : HoldStatus = IfNullThenNilString(RS("Status").Value)
        Dim Balance As Decimal : Balance = IfNullThenZeroCurrency(RS("Sale").Value) - IfNullThenZeroCurrency(RS("Deposit").Value)

        Dim Extra As String

        If HidePriorSales Then
            If LastIndex = Index Then
                ' do nothing
                returnLine = "(skip)"
            Else
                returnLine = AlignString(Trim(LastName) & ",  " & Trim(FirstName), 25, AlignConstants.vbAlignLeft, False) & vbTab & DressAni(CleanAni(Trim(Telephone))) & vbTab & Trim(Address)
                LastIndex = Index
            End If
        Else
            Extra = Space(3) & LeaseNo & Space(3) & HoldStatus & Space(3) & FormatCurrency(Balance)
            If LastIndex = Index Then
                returnLine = Extra
            Else
                If LeaseNo = "" Then
                    returnLine = AlignString(Trim(LastName) & ",  " & Trim(FirstName), 25, AlignConstants.vbAlignLeft, False) & vbTab & DressAni(CleanAni(Trim(Telephone))) & vbTab & Trim(Address)
                Else
                    returnLine = AlignString(Trim(LastName) & ",  " & Trim(FirstName), 25, AlignConstants.vbAlignLeft, False) & vbTab & DressAni(CleanAni(Trim(Telephone))) & vbTab & Trim(Address) & vbCrLf & Extra
                End If
                LastIndex = Index
            End If
        End If
        returnKey = Index
    End Sub

    Private Sub objSource_Refresh() Handles objSource.Refresh
        LastIndex = -1
        If optTelephone.Checked = True Then
            objSource.ListToListBox(CleanAni(InputBox.Text), lstMatches)
        Else
            objSource.ListToListBox(InputBox.Text, lstMatches)
        End If
    End Sub

    Public Function DeveloperEx() As String
        DeveloperEx = "Ctrl-ESC" & vbCrLf & " JOHN CARTER"
    End Function

    Private Sub optTelephone_CheckedChanged(sender As Object, e As EventArgs) Handles optTelephone.CheckedChanged
        optTelephone_Click(optTelephone, New EventArgs)
    End Sub

    Private Sub optSaleNo_CheckedChanged(sender As Object, e As EventArgs) Handles optSaleNo.CheckedChanged
        optSaleNo_Click(optSaleNo, New EventArgs)
    End Sub

    Private Sub lstMatches_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstMatches.SelectedIndexChanged

    End Sub
End Class