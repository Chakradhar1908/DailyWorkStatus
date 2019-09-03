Imports Microsoft.VisualBasic.Compatibility.VB6
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class BillOSale
    Dim printer As New Printer
    Dim LeaseNo As String
    Dim Note As String
    Dim Notes As String

    Public DelDate As String
    Public TransDate As String
    Public Index As Integer
    Public SalesCode As String

    Public BOS2IsHidden As Boolean

    Private PollingSaleDate As Boolean

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public NewStyleLine As Integer         ' Line currently having style (etc) information entered.
    Public DeleteLine As Object         ' Used by Mailcheck.LookupCustomer
    Public KitLines As Integer             ' Number of lines in the current kit - for OrdStatus.
    Public MailRec As Integer
    Public Sale As Decimal
    Public RN As Integer
    Public NsRec1 As Integer, NsRec2 As Integer, NsRec3 As Integer, NsRec4 As Integer, NsRec5 As Integer, NsRec6 As Integer
    Private Ba(0 To Setup_MaxStores_DB - 1) As Integer
    Private OO(0 To Setup_MaxStores_DB - 1) As Integer
    Public PoSold As Single
    Public Rb As Single
    Public NonTaxable As Decimal
    Public PrintBill As Boolean
    Public Written As Decimal
    Public ArCashSls As Decimal
    Public Deposit As Decimal
    'Public Index                        ' Taken from MailCheck.Index, which is variant.
    Public Typpe As String              ' Used by BillOSale.cmdApplyBillOSale_Click.

    Private AddingItem As Boolean       ' While this is true, don't load OrdSelect on Style_GotFocus.
    Private mCurrentLine As Integer        ' Current line selected
    Dim mLastRecord As Integer             ' Used by Mailcheck.GetOrder
    Dim Detail As Integer
    Dim Marginn As New CGrossMargin
    Dim MarginNo As Integer
    Dim Name1 As String
    'Dim LeaseNo As String
    'Dim TransDate As String
    Dim TaxCharged1 As Decimal
    Dim Controll As Decimal
    Dim UndSls As Decimal
    Dim DelSls As Decimal
    Dim TaxRec1 As Decimal
    Dim TaxRec2 As Decimal
    Public SalesTax1 As Decimal
    Public SalesTax2 As Decimal
    Dim TaxCode As Integer           ' Used in GrossMargin and Audit.
    Public TaxRate, LastTaxRate   ' Looks like a string, but could be double/currency.
    Dim TotSale As Decimal
    Dim Mail As MailNew
    Dim Copies As Integer            ' Used in PrintInvoiceCommon to hide certain information from customers.
    'Dim Notes As Variant

    Public SaleHasCCTransactions As Boolean

    'Dim PoNo as integer             ' These variables help combine items into one purchase order.
    Dim LastSale As String
    Dim LastMfg As String

    Dim ProcessSalePOs As Collection

    Private WithEvents mDBNotes As CDbAccessGeneral

    Public IsInternetSale As Boolean
    Private LastQuant As Integer, mLastGridText As String, LGTCol As Integer, LGTRow As Integer

    Private NotesInfo As String
    Public InstallmentTotal As Decimal
    Dim offY As Integer, offX As Integer
    Dim SplitKits As TriState

    Private Sub cboPhone1_Enter(sender As Object, e As EventArgs) Handles cboPhone1.Enter
        ' This event is replacement for Gotfocus of vb6.0
        SelectContents(cboPhone1)
    End Sub

    Private Sub cboPhone2_Enter(sender As Object, e As EventArgs) Handles cboPhone2.Enter
        SelectContents(cboPhone2)
    End Sub

    Private Sub cboPhone3_Enter(sender As Object, e As EventArgs) Handles cboPhone3.Enter
        SelectContents(cboPhone3)
    End Sub

    Private Sub CustomerLast_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CustomerLast.KeyPress
        'If KeyAscii = Asc(",") Then KeyAscii = Asc(" ") :         ' change , to ;
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If e.KeyChar = "," Then
            'i = Asc(" ")
            e.KeyChar = " "
        Else
            'i = Asc(UCase(e.KeyChar))
            e.KeyChar = UCase(e.KeyChar)
        End If
        'i = Asc(UCase(e.KeyChar))
    End Sub

    Private Sub CustomerAddress_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CustomerAddress.KeyPress
        'Dim i As Integer
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'i = Asc(UCase(e.KeyChar))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Private Sub AddAddress_KeyPress(sender As Object, e As KeyPressEventArgs) Handles AddAddress.KeyPress
        'Dim i As Integer
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'i = Asc(UCase(e.KeyChar))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Private Sub CustomerCity_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CustomerCity.KeyPress
        'Dim i As Integer
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'i = Asc(UCase(e.KeyChar))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Private Sub CustomerFirst_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CustomerFirst.KeyPress
        'Dim i As Integer
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'i = Asc(UCase(e.KeyChar))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Public Sub GetSpeechInputMode(ByRef Result As Boolean, ByVal SIType As String, ByVal CtrlName As String)
        If SIType = "spell" And IsIn(CtrlName, "Email", "txtSaleNo") Then Result = True
    End Sub

    Private Sub CustomerAddress2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CustomerAddress2.KeyPress
        'Dim i As Integer
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'i = Asc(UCase(e.KeyChar))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Private Sub CustomerCity2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CustomerCity2.KeyPress
        'Dim i As Integer
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'i = Asc(UCase(e.KeyChar))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Private Sub CustomerPhone1_TextChanged(sender As Object, e As EventArgs) Handles CustomerPhone1.TextChanged
        FormatAniTextBox(CustomerPhone1)
    End Sub

    Private Sub CustomerPhone2_TextChanged(sender As Object, e As EventArgs) Handles CustomerPhone2.TextChanged
        FormatAniTextBox(CustomerPhone2)
    End Sub

    Private Sub CustomerPhone3_TextChanged(sender As Object, e As EventArgs) Handles CustomerPhone3.TextChanged
        FormatAniTextBox(CustomerPhone3)
    End Sub

    Private Sub cboPhone1_TextChanged(sender As Object, e As EventArgs) Handles cboPhone1.TextChanged
        If cboPhone1.Text = "" Then cboPhone1.Text = "Telephone"
        If Len(cboPhone1.Text) > 50 Then
            cboPhone1.Text = Trim(Microsoft.VisualBasic.Left(cboPhone1.Text, 50))
            cboPhone1.SelectionStart = Len(cboPhone1.Text)
        End If
    End Sub

    Private Sub cboPhone2_TextChanged(sender As Object, e As EventArgs) Handles cboPhone2.TextChanged
        If cboPhone2.Text = "" Then cboPhone2.Text = "Telephone"
        If Len(cboPhone2.Text) > 50 Then
            cboPhone2.Text = Trim(Microsoft.VisualBasic.Left(cboPhone2.Text, 50))
            cboPhone2.SelectionStart = Len(cboPhone2.Text)
        End If
    End Sub

    Private Sub cboPhone3_TextChanged(sender As Object, e As EventArgs) Handles cboPhone3.TextChanged
        If cboPhone3.Text = "" Then cboPhone3.Text = "Telephone"
        If Len(cboPhone3.Text) > 50 Then
            cboPhone3.Text = Trim(Microsoft.VisualBasic.Left(cboPhone3.Text, 50))
            cboPhone3.SelectionStart = Len(cboPhone3.Text)
        End If
    End Sub

    Private Sub ShipToFirst_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ShipToFirst.KeyPress
        'Dim i As Integer
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'i = Asc(UCase(e.KeyChar))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Private Sub ShipToLast_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ShipToLast.KeyPress
        'Dim i As Integer
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'i = Asc(UCase(e.KeyChar))
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Public Property MailIndex() As Integer
        Get
            MailIndex = Index
        End Get
        Set(value As Integer)
            Index = value
        End Set
    End Property

    Public Property PorD() As String
        Get
            'PorD = Switch(chkDelivery = 1, "D", chkPickup = 1, "P", True, "")
            PorD = ""
            If chkDelivery.CheckState = 1 Then
                PorD = "D"
            ElseIf chkPickup.CheckState = 1 Then
                PorD = "P"
            ElseIf True Then
                PorD = ""
            End If
        End Get
        Set(value As String)
            chkPickup = IIf(value = "P", 1, 0)
            chkDelivery = IIf(value = "D", 1, 0)
        End Set
    End Property

    Private Sub ShowTimeWindowBox(ByVal Show As Boolean, Optional ByVal Enabled As Boolean = False)

        'dtpDelWindow(0).Value = "7:00 am"
        'dtpDelWindow.Value = New DateTime(dtpDelWindow.Value.Year, dtpDelWindow.Value.Month, dtpDelWindow.Value.Day, 7, 0, 0)
        dtpDelWindow.Value = Date.Parse(dtpDelWindow.Value.Date) & " " & TimeValue("07:00:00 AM")
        'dtpDelWindow(0).Value = ""
        dtpDelWindow.Value = Date.FromOADate(0)
        'dtpDelWindow2.Value = "5:00 pm"
        'dtpDelWindow2.Value = New DateTime(dtpDelWindow2.Value.Year, dtpDelWindow2.Value.Month, dtpDelWindow2.Value.Day, 5, 0, 0)
        dtpDelWindow2.Value = Date.Parse(dtpDelWindow2.Value.Date) & " " & TimeValue("05:00:00 PM")
        'dtpDelWindow2.Value = ""
        dtpDelWindow2.Value = Date.FromOADate(0)

        fraTimeWindow.Visible = Show And (StoreSettings.bUseTimeWindows)
        dtpDelWindow.Enabled = OrderMode("A")
        dtpDelWindow2.Enabled = OrderMode("A")
    End Sub

    Private Sub SetDelWeekday(ByVal DelDate As Date)
        Select Case Weekday(DelDate, cdsFirstDayOfWeek)
            Case 1 : lblDelWeekday.Text = "SUN."
            Case 2 : lblDelWeekday.Text = "MON."
            Case 3 : lblDelWeekday.Text = "TUES."
            Case 4 : lblDelWeekday.Text = "WED."
            Case 5 : lblDelWeekday.Text = "THURS."
            Case 6 : lblDelWeekday.Text = "FRI."
            Case 7 : lblDelWeekday.Text = "SAT."
        End Select
    End Sub

    Private Sub Email_DoubleClick(sender As Object, e As EventArgs) Handles Email.DoubleClick
        'If MsgBox("Send email to " & Email.Text & "?" & vbCrLf & "You should only send email if this computer is setup for email.", vbYesNo, "Send email") = vbYes Then
        RunShellExecute("open", "mailto:" & Email.Text, 0&, 0&, SW_SHOWDEFAULT)
        'End If
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        If MailMode("Book") Then
            'Unload MailCheck
            MailCheck.Close()
            'Unload BillOSale
            Me.Close()
            MailBook.Show()
        ElseIf cmdMainMenu.Text = "Back" Then
            modProgramState.Mail = ""
            'Unload MailCheck
            MailCheck.Close()

            'Unload BillOSale
            Me.Close()

            'Unload InvDefault  ' This shouldn't be necessary, but is..
            InvDefault.Close()
        Else
            modProgramState.Mail = ""
            'Unload MailCheck
            MailCheck.Close()

            'Unload BillOSale
            Me.Close()

            'Unload InvDefault  ' This shouldn't be necessary, but is..
            InvDefault.Close()

            MainMenu.Show()
            modProgramState.Order = ""
            modProgramState.ArSelect = ""
        End If
    End Sub

    Private Sub UpdateForm()
        On Error Resume Next

        MoveControl(imgLogo, 3330, 60, 5775, 1995, True)
        'ResizeAndCenterPicture(imgLogo, StoreLogoPicture())

        'imgLogo.Picture = Nothing
        imgLogo.Image = Nothing
        'imgLogo.Picture = StoreLogoPicture()
        'imgLogo.Image = StoreLogoPicture()
        '--------------------
        Dim StoreNum As Integer
        If StoreNum = 0 Then StoreNum = StoresSld
        StoreNum = FitRange(1, StoreNum, Setup_MaxStores)
        'StoreLogoPicture = LoadPictureStd(StoreLogoFile(StoreNum))
        imgLogo.Image = Image.FromFile(StoreLogoFile(StoreNum))
        '-----------

        StoreName.Text = IIf(imgLogo.Image IsNot Nothing, "", StoreSettings.Name)
        StoreAddress.Text = IIf(imgLogo.Image IsNot Nothing, "", StoreSettings.Address)
        StoreCity.Text = IIf(imgLogo.Image IsNot Nothing, "", StoreSettings.City)
        StorePhone.Text = IIf(imgLogo.Image IsNot Nothing, "", StoreSettings.Phone)

        ' New, View
        If Not OrderMode("A", "B", "E") Then cmdApplyBillOSale.Enabled = False

        'Mailing list
        If MailMode("ADD/Edit") Or ArMode("S", "A") Then cmdApplyBillOSale.Enabled = True

        '---------> NOTE: 'Remove this line later after complete coding has been done.     <----------
        cmdApplyBillOSale.Enabled = True
    End Sub

    Private Sub DoPrintType(Optional ByVal DoShow As Boolean = False)
        'fraButtons.Width = IIf(DoShow, 2655, 1815)
        fraButtons.Width = IIf(DoShow, 200, 150)
        fraPrintType.Visible = DoShow
        If DefaultMailingLabelType() = 30323 Then
            opt30323.Checked = True
        Else
            opt30252.Checked = True
        End If
    End Sub

    Private Sub cmdShowBodyOfSale_Click(sender As Object, e As EventArgs) Handles cmdShowBodyOfSale.Click
        If OrderMode("A") Then
            'cmdApplyBillOSale.Value = True
            cmdApplyBillOSale.PerformClick()
        Else
            BillOSale2_Show()
        End If
    End Sub

    Private Sub cmdSoldTags_Click(sender As Object, e As EventArgs) Handles cmdSoldTags.Click
        Dim I As Integer, X As Integer, DoAll As Boolean

        If LeaseNo = "" Then
            MsgBox("SOLD tags only allowed for completed sales.")
            Exit Sub
        End If

TryAgain:
        '        For I = 0 To UGridIO1.LastRowUsed
        '            If IsNotIn(QueryStatus(I), "ST") Then GoTo NextItem
        '            If Not IsItem(QueryStyle(I)) Then GoTo NextItem
        '            If DoAll Or Not DescHasSoldTagPrinted(QueryDesc(I)) Then X = X + PrintSoldTags(QueryStyle(I), CustomerLast.Text, LeaseNo, 1)
        '            SetDesc(I, DescSetSoldTagPrinted(QueryDesc(I), LeaseNo, QueryStyle(I))) ' this gets the new value and updates it in the GM table if possible
        'NextItem:
        '        Next

        If Not DoAll And X = 0 Then
            If MsgBox("Re-print all SOLD tags?", vbYesNo + vbQuestion, "Confirm Reprint") = vbYes Then
                DoAll = True
                GoTo TryAgain
            End If
        End If

        MsgBox("Complete!")
    End Sub

    Public Function QueryStatus(ByVal RowNum As Integer) As String
        QueryStatus = QueryGridField(RowNum, BillColumns.eStatus)
    End Function

    Public Function QueryStyle(ByVal RowNum As Integer) As String
        QueryStyle = Microsoft.VisualBasic.Left(QueryGridField(RowNum, BillColumns.eStyle), Setup_2Data_StyleMaxLen)
    End Function

    Public Function QueryDesc(ByVal RowNum As Integer) As String
        QueryDesc = Microsoft.VisualBasic.Left(QueryGridField(RowNum, BillColumns.eDescription), Setup_2Data_DescMaxLen)
    End Function

    Public Sub SetDesc(ByVal RowNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        SetGridField(RowNum, BillColumns.eDescription, Microsoft.VisualBasic.Left(CellVal, Setup_2Data_DescMaxLen), NoDisplay)
    End Sub

    Private Function SetGridField(ByVal RowNum As Integer, ByVal ColNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        If NoDisplay Then
            UGridIO1.SetValue(RowNum, ColNum, CellVal)
        Else
            UGridIO1.SetValueDisplay(RowNum, ColNum, CellVal)
        End If
    End Function

    ' MJK Generic grid data access, to ease field access without changing X.
    Private Function QueryGridField(ByVal RowNum As Integer, ByVal ColNum As Integer) As String
        ' These methods don't get the current (changed) values!
        '  QueryGridField = UGridIO1.GetColumn(ColNum).CellText(RowNum)
        QueryGridField = UGridIO1.GetValue(RowNum, ColNum)
    End Function

    Private Sub DatepickerValuechanged(sender As Object, e As EventArgs) Handles dtpDelWindow.ValueChanged, dtpDelWindow2.ValueChanged
        'NOTE: THIS EVENT IS REPLACEMENT FOR Private Sub dtpDelWindow_Change(Index As Integer) OF VB 6.0.
        Dim D1 As Date, D2 As Date

        If IsDate(dtpDelWindow.Value) And IsDate(dtpDelWindow2.Value) Then
            D1 = TimeValue(dtpDelWindow.Value)
            'If DateAfter(D1, "11:00p", False, "n") Then
            If DateAfter(D1, TimeValue("11:00:00 PM"), False, DateInterval.Minute) Then
                'dtpDelWindow.Value = dtpDelWindow.Value.Date & " " & TimeValue("10:00:00 PM")
                dtpDelWindow.Value = Date.Parse(dtpDelWindow.Value.Date) & " " & TimeValue("10:00:00 PM")
                D1 = TimeValue(dtpDelWindow.Value)
            End If

            D2 = TimeValue(dtpDelWindow2.Value)
            'If Not DateAfter(D2, D1, False, "n") Then
            If Not DateAfter(D2, D1, False, DateInterval.Minute) Then
                dtpDelWindow2.Value = Date.Parse(dtpDelWindow2.Value.Date) & " " & TimeValue(DateAdd(DateInterval.Minute, 30, D1))
                'dtpDelWindow2.Value = dtpDelWindow2.Value.Date & " " & TimeValue(DateAdd(DateInterval.Minute, 30, D1))
            End If
        End If
    End Sub

    Public Sub SetBusiness(ByVal IsBusiness As Boolean)
        Dim fontsize As Font = CustomerLast.Font

        If IsBusiness Then
            optBusiness.Checked = True
            lblFirst.Visible = False
            lblLast.Visible = False
            CustomerFirst.Text = ""
            CustomerFirst.Visible = False
            CustomerLast.Font = New Font(fontsize.FontFamily, 1)

            'CustomerLast.Move(120, CustomerLast.Top, 5850, 330)
            'CustomerLast.Location = New Point(120, CustomerLast.Top)
            CustomerLast.Location = New Point(8, CustomerLast.Top)
            'CustomerLast.Size = New Size(5850, 330)
            CustomerLast.Size = New Size(420, 26)
            CustomerLast.Font = New Font(fontsize.FontFamily, 12)
        Else
            optIndividual.Checked = True
            lblFirst.Visible = True
            lblLast.Visible = True
            CustomerFirst.Visible = True
            CustomerLast.Font = New Font(fontsize.FontFamily, 1)
            'CustomerLast.Move 3240, CustomerLast.Top, 2715, 330
            'CustomerLast.Location = New Point(3240, CustomerLast.Top)
            CustomerLast.Location = New Point(222, CustomerLast.Top)
            'CustomerLast.Size = New Size(2715, 330)
            CustomerLast.Size = New Size(203, 26)
            CustomerLast.Font = New Font(fontsize.FontFamily, 12)
        End If
    End Sub

    Private Sub BillOSale_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        ' This is the replacement for Form_Click() event of vb 6.0
        BillOSale2_Hide()
    End Sub

    Private Sub imgLogo_Click(sender As Object, e As EventArgs) Handles imgLogo.Click
        BillOSale2_Hide()
    End Sub

    Private Sub Email_Enter(sender As Object, e As EventArgs) Handles Email.Enter
        'this event is replacement for gotfocus of vb6.0
        BillOSale2_Hide()
    End Sub

    Private Sub imgCalendar_Click(sender As Object, e As EventArgs) Handles imgCalendar.Click
        ' Show the delivery calendar.
        If cmdMainMenu.Text <> "Back" Then   ' bfh20050816 - this is raised modal and so we can't modally show the calendar form..
            'MousePointer = vbHourglass
            Me.Cursor = Cursors.WaitCursor
            Calendar.LoadedByForm = True
            'MousePointer = vbDefault
            Me.Cursor = Cursors.Default
            'Calendar.Show vbModal, Me
            Calendar.ShowDialog(Me)
        End If
    End Sub

    Private Sub lblBalDueCaption_Click(sender As Object, e As EventArgs) Handles lblBalDueCaption.Click
        Recalculate()
    End Sub

    Public Sub Recalculate(Optional ByVal DontAdjustTaxes As Boolean = False)
        Dim Untaxed As Decimal, PriceError As Boolean

        ' recalculate: clear variables first
        BalDue.Text = 0
        BalDue.Text = 0
        SalesTax2 = 0
        SalesTax1 = 0
        OrdSelect.SalesTax2 = 0
        OrdSelect.SalesTax1 = 0
        Sale = 0
        Deposit = 0       ' Total deposit on sale
        Written = 0
        NonTaxable = 0

        ' Loop through each row, doing only what's necessary.
        Dim Xx As Integer, YY As Integer
        Dim oPrice As String, tPrice As Decimal, tStyle As String, Desc As String
        '  Dim AddTax As Currency

        For Xx = 0 To UGridIO1.LastRowUsed
            oPrice = UGridIO1.GetValue(Xx, BillColumns.ePrice)
            tStyle = Trim(UGridIO1.GetValue(Xx, BillColumns.eStyle))
            tPrice = GetPrice(oPrice, PriceError)
            Desc = Trim(UGridIO1.GetValue(Xx, BillColumns.eDescription))

            If PriceError Then
                MsgBox("There Is a Mistake On The Price Entered!", vbCritical)
                SetPrice(Xx, 0)
                UGridIO1.Row = Xx
                PriceFocus()
                ' Wouldn't it be better to completely recalculate, then bring focus back here?
                ' Irrelevant for now.
                ' Except that it's possible to remove sales tax this way, and process the sale....
                Exit Sub
            End If

            Select Case tStyle
                Case "PAYMENT"
                    Sale = Sale - tPrice
                    Deposit = Deposit + tPrice
                Case "TAX1"
                    If Not Trim(Desc) Like "SALES TAX DIFF.*" Then  ' dont' adjust adjustment sales tax diff
                        If Not DontAdjustTaxes And Not IsInternetSale Then
                            If IsPalazzo() And (Sale - NonTaxable) > 1600 And Xx = 2 Then
                                SalesTax1 = ((Sale - 1600) * 0.06) + 132
                            Else
                                SalesTax1 = CurrencyFormat(GetStoreTax1() * (Written - NonTaxable - SalesTax2 - SalesTax1))
                            End If
                        Else  ' don't recalculate the salestax for an internet sale... it was already done for us!
                            SalesTax1 = tPrice
                        End If
                    Else
                        SalesTax1 = tPrice
                    End If

                    tPrice = SalesTax1
                    Written = Written + tPrice
                    Sale = Sale + tPrice
                    Untaxed = 0
                Case "TAX2"
                    Dim Tax2 As Decimal
                    If Not Trim(Desc) Like "SALES TAX DIFF.*" Then  ' dont' adjust adjustment sales tax diff
                        If Not DontAdjustTaxes Then
                            GetTax2(Val(QueryQuan(Xx)))
                            Tax2 = CurrencyFormat(TaxRate * (Written - NonTaxable - SalesTax1 - SalesTax2))
                        Else
                            Tax2 = tPrice
                        End If
                    Else
                        Tax2 = tPrice
                    End If

                    tPrice = Tax2
                    Written = Written + tPrice
                    Sale = Sale + tPrice
                    SalesTax2 = SalesTax2 + Tax2
                    Untaxed = 0
                Case "SUB"    ' recalulate bal
                    tPrice = Written - Deposit
                Case "--- Adj ---" ' this is also really a subtotal line
                    tPrice = Written - Deposit
                Case Else
                    Dim HasTaxLine As Boolean
                    For YY = Xx To UGridIO1.LastRowUsed
                        If IsIn(UGridIO1.GetValue(YY, BillColumns.eStyle), "TAX1", "TAX2") Then
                            HasTaxLine = True
                        End If
                        If IsIn(UGridIO1.GetValue(YY, BillColumns.eStyle), "--- Adj ---") Then
                            Exit For
                        End If
                    Next
                    Sale = Sale + tPrice
                    Written = Written + tPrice
                    If IsItemNontaxable(tStyle, HasTaxLine) Then
                        NonTaxable = NonTaxable + tPrice
                    Else
                        Untaxed = Untaxed + tPrice   ' Untaxed is the amount not yet accounted for in taxes.
                    End If

                    '        ' for when LAB or DEL is taxed without a TAX1/2 line
                    '        If IsIn(tStyle, "LAB", "DEL") And Not HasTaxLine And Not IsItemNontaxable(tStyle, HasTaxLine) Then
                    '          AddTax = CurrencyFormat(GetStoreTax1 * tPrice)
                    '        End If
            End Select

            If tPrice = 0 Then
                If oPrice <> "" Then SetPrice(Xx, "")
            Else
                Dim uPrice As String
                uPrice = CurrencyFormat(tPrice)
                If uPrice <> oPrice Then SetPrice(Xx, uPrice)
            End If
        Next

        'SalesTax1 = SalesTax1 + AddTax
        'Written = Written + AddTax
        'Sale = Sale + AddTax

        NonTaxable = NonTaxable + Untaxed
        OrdSelect.SalesTax1 = CurrencyFormat(SalesTax1)
        OrdSelect.SalesTax2 = CurrencyFormat(SalesTax2)
        BalDue.Text = CurrencyFormat(Written - Deposit)
        If Sale <> Written - Deposit Then ErrMsg("Methods of determining Balance Due don't match.")
        DeleteLine = ""
    End Sub

    Public Sub SetPrice(ByVal RowNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        SetGridField(RowNum, BillColumns.ePrice, Format(CellVal, "###,###.00"), NoDisplay)
    End Sub

    Public Sub PriceFocus(Optional ByVal nRow As Integer = -1)
        GridFocus(BillColumns.ePrice, nRow) ' 6
    End Sub

    Public Sub GridFocus(ByVal ColNum As Integer, Optional ByVal RowNum As Integer = -1)
        On Error Resume Next
        UGridIO1.Loading = True
        If RowNum = -1 Then RowNum = CurrentLine
        If ColNum >= 0 And ColNum <= UGridIO1.MaxCols Then UGridIO1.Col = ColNum
        If RowNum >= 0 And RowNum <= UGridIO1.MaxRows Then UGridIO1.Row = RowNum
        UGridIO1.Loading = False
        UGridIO1.Select()
    End Sub

    Private Sub GetTax2(ByVal Quantity As Integer)
        TaxRate = QuerySalesTax2(Quantity - 1)
        LastTaxRate = TaxRate
        ActualTax()
    End Sub

    Private Sub ActualTax()
        Dim TaxListing As String, Tax As String
        Dim ItemCount As Integer

        TaxListing = TaxRate
        TaxRate = ""
        Tax = ""
        For ItemCount = 1 To Len(TaxListing) + 1
            Tax = Mid(TaxListing, ItemCount, 1)
            If Trim(Tax) = "" Then Exit Sub
            TaxRate = TaxRate + Tax
        Next
    End Sub

    Public Function SaleTotal(Optional ByVal tType As String = "") As Decimal
        Dim S As sSale
        S = New sSale

        S.LoadFromBillOSale()
        SaleTotal = S.SubTotal(tType)
        DisposeDA(S)
    End Function

    Public Function QueryQuan(ByVal RowNum As Integer) As Double
        QueryQuan = Val(QueryGridField(RowNum, BillColumns.eQuant))
    End Function

    Public Function SubTotal(ByVal TopRow As Integer, Optional ByVal TaxableOnly As Boolean = False) As Decimal
        Dim I As Integer, P As Decimal, S As String

        For I = 0 To TopRow
            S = IfNullThenNilString(QueryStyle(I))
            P = IfNullThenZeroCurrency(QueryPrice(I))
            If P <> 0 Then
                If S = "SUB" Then
                    ' Do nothing.
                ElseIf S = "PAYMENT" Then
                    SubTotal = SubTotal - P
                Else
                    If TaxableOnly And IsItemNontaxable(S) Then
                        ' Don't add the price.
                    Else
                        SubTotal = SubTotal + P
                    End If
                End If
            End If
        Next
    End Function

    Private Sub lblGrossSalesCaption_Click(sender As Object, e As EventArgs) Handles lblGrossSalesCaption.Click
        If SpeechActive() Then frmSpeech.TestCommand("zip", "CurrentFormLabels")
    End Sub

    Private Sub optIndividual_Click(sender As Object, e As EventArgs) Handles optIndividual.Click
        SetBusiness(False)
    End Sub

    Private Sub optBusiness_Click(sender As Object, e As EventArgs) Handles optBusiness.Click
        SetBusiness(True)
        On Error Resume Next
        CustomerLast.Select()
    End Sub

    Private Sub ShipToFirst_Enter(sender As Object, e As EventArgs) Handles ShipToFirst.Enter
        ShipToLast.TabStop = True
        CustomerAddress2.TabStop = True
        CustomerCity2.TabStop = True
        CustomerZip2.TabStop = True
        CustomerPhone3.TabStop = True
    End Sub

    Private Sub BillOSale_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' This is replacement for form_unload event of vb 6.0
        ' This event is replacement for form_queryunload event also.

        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
        End If

        On Error Resume Next
        DisposeDA(Marginn)
        'VerifyMailRecUnique MailRec, vbTrue  ' clear it if it was saved..
        MailRec = 0

        'Unload frmSalesList
        frmSalesList.Close()

        'Unload OrdSelect
        OrdSelect.Close()
        ' These shouldn't be necessary, but something sloppy is happening.
        'Unload InvDefault
        InvDefault.Close()
        '  Unload InvCkStyle

    End Sub

    Private Sub BillOSale_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        'Note: This event is replacement for Form_Activate event of vb 6.0

        If IsFormLoaded("OrdPay") Then Exit Sub

        On Error Resume Next
        'If txtSaleNo.Visible And txtSaleNo = "" And Not IsFormLoaded("InvCkStyle") And Not IsFormLoaded("OrdStatus") Then txtSaleNo.SetFocus
        ShipToFirst.TabStop = False
        ShipToLast.TabStop = False
        CustomerAddress2.TabStop = False
        CustomerCity2.TabStop = False
        CustomerZip2.TabStop = False
        CustomerPhone3.TabStop = False
        UpdateForm()
    End Sub

    Private Sub BillOSale_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UpdateForm()
        SetButtonImage(cmdApplyBillOSale, "ok")
        SetButtonImage(cmdCancel)
        ColorDatePicker(dteSaleDate)
        ColorDatePicker(dteDelivery)

        txtSaleNo.Visible = False
        'Me.Left = (Screen.Width - Me.Width) / 2
        Me.Left = (Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2
        'Me.Height = 8910 'reset to long version
        Me.Top = 0

        'This example sets the size of a form to 75 percent of screen size
        'and centers the form when it is loaded. To try this example,
        'paste the code into the Declarations section of a form.
        'Width = Screen.Width * 0.75  ' Set width of form.
        'Height = Screen.Height * 0.75  ' Set height of form.
        'Left = (Screen.Width - Width) / 2   ' Center form horizontally.
        'Top = (Screen.Height - Height) / 2   ' Center form vertically.

        dteSaleDate.Value = DateTime.Parse(DateFormat(Now))
        'dteSaleDate.Value = Now
        dteDelivery.Value = DateTime.Parse(DateFormat(Now))

        Order = "A"     '------> It will be assigned in modMainMenu. Because modMainMenu code is not completed
        'temporarily assigned here to run the below select case Order code. After modMainMenu completed, 
        'remove this line Order = "A" from here.

        Select Case Order
            Case "A" : Text = "New Sale"
            Case "C" : Text = "Void Order"
            Case "B" : Text = "Deliver Sale"
            Case "D" : Text = "Customer Payment"
        End Select

        If OrderMode("A") Then
            Arrange(True, False)
            LoadFakeGrid(True)
            If Not StoreSettings.bManualBillofSaleNo Then
                txtSaleNo.Visible = False
                dteDelivery.Visible = False
                lblDelDate.Text = ""
            Else
                txtSaleNo.Visible = True
            End If

            SaleStatus.Text = "New"
            dteSaleDate.Value = DateFormat(Now)
        End If

        If Not OrderMode("A") Then
            cmdApplyBillOSale.Enabled = False
            chkDelivery.Enabled = False  ' Disallow editing delivery date.
            chkPickup.Enabled = False
            dteSaleDate.Enabled = False
            Sales1.ReadOnly = True
            Sales2.ReadOnly = True
            Sales3.ReadOnly = True
            SalesSplit1.Enabled = False
            SalesSplit2.Enabled = False
            SalesSplit3.Enabled = False
        End If

        If MailMode("ADD/Edit", "Book") Or ArMode("S", "A") Then
            ' mailing list
            Arrange(False, False)
            cmdApplyBillOSale.Enabled = True
            cmdCancel.Enabled = True
        End If
        DoPrintType(True)

        If OrderMode("B", "E") Then ' Allow mail edit on Delivery or Order
            cmdApplyBillOSale.Enabled = True
        End If

        lblDelWeekday.Text = "None"

        dteDelivery.Visible = False
        If IsGrizzlys() Then
            cboCustType.Items.Add("Customer")
            cboCustType.Items.Add("Bear Club 10")
            cboCustType.Items.Add("Bear Club 20")
        Else
            cboCustType.Items.Add("Customer")
            cboCustType.Items.Add("Resale 1")
            cboCustType.Items.Add("Resale 2")
            cboCustType.Items.Add("Resale 3")
            cboCustType.Items.Add("Resale 4")
            cboCustType.Items.Add("Resale 5")
            cboCustType.Items.Add("Resale 6")
        End If

        LoadAdvTypesIntoComboBox(cboAdvertisingType, StoresSld)
        LoadSalesTax2IntoComboBox(cboTaxZone, StoresSld, True)

        If MailMode("ADD/Edit", "Book") Or ArMode("S") Then
            Text = "Customer Record"
            dteSaleDate.Enabled = False
            dteDelivery.Visible = False
            Sales1.Visible = False
            Sales2.Visible = False
            Sales3.Visible = False
            lblSaleNoCaption.Visible = False
            lblSales1.Visible = False
            lblSales2.Visible = False
            lblSales3.Visible = False
            lblDateCaption.Visible = False
            lblStatusCaption.Visible = False
            chkDelivery.Visible = False
            chkPickup.Visible = False
            dteSaleDate.Visible = False
            lblDelWeekday.Visible = False
            imgCalendar.Visible = False
            lblDelDate.Visible = False
            txtSaleNo.Visible = False
            SaleStatus.Visible = False
        End If

        LoadSalesSplitBoxes()
    End Sub

    Public Sub BillOSale2_Show()
        LoadBOS2()
        Arrange(BoS2:=True)
    End Sub

    Private Sub BillOSale_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        HoverPic()
        ResetLastLoginExpiry()
    End Sub

    Public Sub BillOSale2_Hide()
        Arrange(BoS2:=False)
    End Sub

    Private Sub LoadFakeGrid(Optional ByVal Visible As Boolean = False)
        With ugrFake
            .AddColumn(0, "Style Number", 100, True, False)
            .AddColumn(1, "Manufacturer", 150, False, False)
            .AddColumn(2, "Loc", 30, True, False)
            .AddColumn(3, "Status", 60, False, False)
            .AddColumn(4, "Quant.", 50, False, False, MSDataGridLib.AlignmentConstants.dbgRight)
            .AddColumn(5, "Description", 200, False, False)
            .AddColumn(6, "Price", 50, False, False, MSDataGridLib.AlignmentConstants.dbgRight)
            .AddColumn(7, "VendorNo", 0, True, False, , False)
            .MaxCols = 8
            .MaxRows = 20
            .Initialize()
            .GetDBGrid.AllowUpdate = False
            .Activated = True
            .Refresh()
        End With
        ugrFake.Visible = Visible
    End Sub

    Private Sub LoadBOS2()
        UGridIO1.Loading = True
        SetButtonImage(cmdProcessSale, "ok")
        SetButtonImage(cmdNextSale, "next")
        SetButtonImage(cmdMainMenu, "menu")
        SetButtonImage(cmdClear, "cancel")
        SetButtonImage(Notes_Open, "notes")

        If CustomerLast.Text = "CASH & CARRY" Then  'added 01-31-2003 to prevent a cash&Carry
            Index = 0 '""
            Marginn.Index = "" 'picking up last name and index
            Marginn.Name = ""
        End If

        OrdSelect.TaxApplied = ""
        With rtb
            .File = CustomerTermsMessageFile()
            .FileRead(False)
            .RichTextBox.Enabled = False
        End With

        With rtbStorePolicy
            .File = StorePolicyMessageFile()
            .FileRead(True)
            .RichTextBox.Enabled = False
        End With

        With UGridIO1
            .AddColumn(0, "Style Number", 100, True, False)
            .AddColumn(1, "Manufacturer", 200, False, False)
            .AddColumn(2, "Loc", 30, True, False)
            .AddColumn(3, "Status", 50, False, False)
            '.AddColumn(4, "Quant.", 50, False, False, MSDBGrid.AlignmentConstants.dbgRight)
            .AddColumn(4, "Quant.", 50, False, False, MSDataGridLib.AlignmentConstants.dbgRight)
            .AddColumn(5, "Description", 250, False, False)
            '.AddColumn(6, "Price", 70, False, False, MSDBGrid.AlignmentConstants.dbgRight)
            .AddColumn(6, "Price", 70, False, False, MSDataGridLib.AlignmentConstants.dbgRight)
            .AddColumn(7, "VendorNo", 0, True, False, , False)
            .AddColumn(8, "TransID", 0, True, False, , False)
            .MaxCols = 9
            .MaxRows = MaxLines
            .Initialize()

            With .GetDBGrid
                '.RowHeight = .Height / (items_per_page + 2) ' This handles the height of the individual rows. Which will indirectly effect the number of rows displayed.
                '.RowHeight = .Height / (18 + 2) ' This handles the height of the individual rows. Which will indirectly effect the number of rows displayed.
                .RowHeight = .Height / (18) ' This handles the height of the individual rows. Which will indirectly effect the number of rows displayed.
                .Height = .RowHeight * 20
            End With
            .Activated = True
            .Refresh()
            .GetDBGrid.AllowDelete = (OrderMode("A"))
            .GetDBGrid.AllowUpdate = (OrderMode("A"))
            .GetDBGrid.AllowRowSizing = False
            X = 0
            .Col = 0
            .Row = 0

        End With

        '  VerifyMailRecUnique MailRec, vbTrue  ' clear it if it was saved..
        MailRec = 0 : PrintBill = False : X = 0
        cmdProcessSale.Enabled = False

        If OrderMode("B") Then
            cmdNextSale.Enabled = False
        End If

        cmdSoldTags.Visible = IsDevelopment()

        If OrderMode("A") Then
            cmdPrint.Enabled = False
            cmdEmail.Enabled = False
            cmdSoldTags.Enabled = False
            Notes_Open.Enabled = False
        Else
            cmdClear.Enabled = False
            Notes_Open.Enabled = True
        End If

        If OrderMode("E") Then
            'scan
            ScanUp123.Enabled = True
            ScanDn.Enabled = True
        Else
            ScanUp123.Enabled = False
            ScanDn.Enabled = False
        End If
        UGridIO1.Loading = False

        EnableDiscountButtons()
        '  -- Disabled until we decide how we want it to go.
        '  If Order = "A" Then
        '    If MainMenu.LastLoginName = "EVERYBODY" Then MainMenu.LastLoginName = ""  ' This makes MainMenu not ask for a password..
        '    ChangePriceEnabled Not CheckAccess("Give Discounts", False, True, False), False
        '  Else
        '    cmdChangePrice.Visible = False
        '    cmdNoChangePrice.Visible = False
        '  End If
        'MessageBox.Show(UGridIO1.AxDataGrid1.Bookmark)
        'UGridIO1.AxDataGrid1.Row = 0
        'UGridIO1.AxDataGrid1.Col = 0
        'UGridIO1.AxDataGrid1.Text = "akekekw"
    End Sub

    Public Sub Arrange(Optional ByVal TALL As TriState = vbUseDefault, Optional ByVal BoS2 As TriState = vbUseDefault)
        Const FRM_SHORT_H = 6750
        'Const FRM_TALL_H = 9280
        Const FRM_TALL_H = 700

        'If BoS2 <> vbUseDefault Then                 ----------> Remove the if block comment later.
        '    fraBOS2.Visible = BoS2
        '    fraBOS2.Left = IIf(BoS2, 0, -15000)
        'End If

        If TALL <> vbUseDefault Then
            'Me.Height = IIf(TALL = vbTrue, FRM_TALL_H, FRM_SHORT_H)
            If TALL = vbTrue Then
                Me.Size = New Size(Me.Width, FRM_TALL_H)
            Else
                Me.Size = New Size(Me.Width, FRM_SHORT_H)
            End If
        End If
    End Sub

    Public Property X() As Integer
        Get
            X = CurrentLine
        End Get
        Set(value As Integer)
            CurrentLine = value
        End Set
    End Property

    Private Property Style() As String
        Get
            Style = QueryStyle(CurrentLine)
        End Get
        Set(value As String)
            SetStyle(CurrentLine, value, False)
        End Set
    End Property

    Private Sub EnableDiscountButtons(Optional ByVal vUnlock As Integer = -1)
        Select Case Order
            Case "A"
                If PrintBill Then
                    cmdChangePrice.Visible = False
                Else
                    cmdChangePrice.Visible = True
                End If
            Case Else
                cmdChangePrice.Visible = False
        End Select
    End Sub

    Private WriteOnly Property StyleSet() As String
        'Get
        '    Return Nothing
        'End Get
        Set(value As String)
            SetStyle(CurrentLine, value, True)
        End Set
    End Property

    Public Sub StyleFocus(Optional ByVal nRow As Integer = -1)
        GridFocus(BillColumns.eStyle, nRow) ' 0
    End Sub

    Private Property Mfg() As String
        Get
            Mfg = QueryMfg(CurrentLine)
        End Get
        Set(value As String)
            SetMfg(CurrentLine, value, False)
        End Set
    End Property

    Public Sub GridMove(ByVal MoveAmount As Object)
        'If (MoveAmount > 15) Then UGridIO1.MoveRowDown(Val(MoveAmount - 15))
    End Sub

    Private WriteOnly Property MfgSet() As String
        'Get
        '    Return Nothing
        'End Get
        Set(value As String)
            SetMfg(CurrentLine, value, True)
        End Set
    End Property

    Public Function GetGrid() As UGridIO
        'GetGrid = UGridIO1
    End Function

    Public Sub GridRefresh()
        'UGridIO1.Refresh(True)
    End Sub

    Public Sub MfgFocus(Optional ByVal nRow As Integer = -1)
        GridFocus(BillColumns.eManufacturer, nRow) ' 1
    End Sub

    Private ReadOnly Property PoNo() As Integer
        Get
            Dim tPO As String
            If ProcessSalePOs Is Nothing Then
                ProcessSalePOs = New Collection
            End If
            On Error Resume Next
            If Trim(Status) = "SO" Or Trim(Status) = "SS" Or Trim(Style) = "NOTES" And Trim(Mfg) <> "" Then
                tPO = ""
                tPO = ProcessSalePOs(Mfg)
                If tPO = "" Then
                    tPO = GetPoNo()
                    ProcessSalePOs.Add(tPO, Mfg)
                End If
                PoNo = tPO
            End If
        End Get

    End Property

    Private Sub UGridIO1_Change()
        'LastGridTextAlt = UGridIO1.Text

        'If UGridIO1.Col = BillColumns.eDescription Then
        '    '  Debug.Print "BOS TEXT: " & UGridIO1.Text
        '    FormatHelper(UGridIO1.Text)
        'End If
    End Sub

    Private WriteOnly Property LastGridTextAlt() As String
        Set(value As String)
            LGTRow = UGridIO1.Row
            LGTCol = UGridIO1.Col
            mLastGridText = value
        End Set
    End Property

    Private Sub Style_DblClick()
        If Style = "PAYMENT" Then ReversePayment() : Exit Sub

        Exit Sub

        If OrderMode("A") Then
            If MsgBox("Do you want to delete this line?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
            DeleteLine = "Y"

            ' cases to handle:
            '   0123456789012345678 mMaxLines
            '   ------------------- 0
            '   0------------------ 1
            '   01----------------- 2
            '   0123456789--------- 10
            '   012-456789--------- 10
            '   0124456789--------- 10
            '   0124056789--------- 10
            '   0124056789--------- 10
            '   0124056789--------- 10
            '   0123456789012345678 19
            Dim LoopRow As Integer
            Dim DeleteLast As Boolean
            DeleteLast = False
            '    If (X = (UGridIO1.MaxRows - 1)) Then
            '        DeleteLast = True
            '    Else
            '        X = X + 1
            '        If (Trim(QueryStyle(X)) = "") _
            'And (X <> 0) Then
            '            DeleteLast = True
            '        End If
            '    End If

            RowClear(X)
            'For LoopRow = X To UGridIO1.MaxRows - 2
            '    X = LoopRow
            '    If Trim(QueryStyle(LoopRow)) = "" Then
            '        RowCopy(LoopRow, LoopRow + 1)
            '        RowClear(LoopRow + 1)   '  clear next line
            '    End If
            'Next
            Recalculate()
            X = X + IIf(DeleteLast, -1, 0)
            If DeleteLast Then
                '!-11/16/98:AA: Price.SetFocus
                'Unload InvDefault
                InvDefault.Close()
            End If
            DeleteLine = ""
        End If
    End Sub

    Public Sub RowClear(ByVal RowToClear As Integer)
        SetStyle(RowToClear, "")
        SetMfg(RowToClear, "")
        SetLoc(RowToClear, "")
        SetStatus(RowToClear, "")
        SetQuan(RowToClear, "")
        SetDesc(RowToClear, "")
        SetPrice(RowToClear, "")
        SetMfgNo(RowToClear, "")
    End Sub

    Private Sub Mfg_Change(ByVal Index As Integer)
        Mfg = UCase(Mfg)
    End Sub

    Private Sub RowCopy(ByVal xTo As Integer, ByVal xFrom As Integer)
        SetStyle(xTo, QueryStyle(xFrom))
        SetMfg(xTo, QueryMfg(xFrom))
        SetLoc(xTo, QueryLoc(xFrom))
        SetStatus(xTo, QueryStatus(xFrom))
        SetQuan(xTo, QueryQuan(xFrom))
        SetDesc(xTo, QueryDesc(xFrom))
        SetPrice(xTo, QueryPrice(xFrom))
    End Sub

    Private Sub Status_DblClick()
        Dim tMargin As New CGrossMargin, R As String, L As Integer, NewPoNo As String, CL As Boolean

        If OrderMode("A") Then Exit Sub   ' Can't convert a new sale..

        Select Case Trim(Status)
            Case "ST"       'convert kit stock to so
                'Load OrdConvertLAW
                OrdConvertLAW.SetupForm(Trim(Style), Trim(Status), Loc)
                'OrdConvertLAW.Show vbModal, Me
                OrdConvertLAW.ShowDialog(Me)
                R = OrdConvertLAW.Result
                L = OrdConvertLAW.Location
                CL = OrdConvertLAW.Cancelled
                'Unload OrdConvertLAW
                OrdConvertLAW.Close()
                If CL Then DisposeDA(tMargin) : Exit Sub

                MailCheck.GetMarginLine()

                Select Case R
                    Case "ST" : ConvertSTToLAW()  ' Really LAW, but...
                    Case "SO"
                        tMargin.Load(MailCheck.MarginNo, "#MarginLine")
                        ConvertToSO()
                        NewPoNo = MakePo(tMargin)
                        SetPoNoOnDetailLine(tMargin.Detail, NewPoNo)
                        DescEnabled = True
                    Case "PO"
                        tMargin.Load(MailCheck.MarginNo, "#MarginLine")
                        ConvertSTToPO(tMargin)
                    Case "DELTW"
                        ConvertSTToDELTW()
                End Select
                If Val(L) <> Val(Loc) Then AdjustItemLocation(L, Loc)
            Case "PO"       ' convert PO back to ST
                Dim xRS As ADODB.Recordset, N As Integer
                MailCheck.GetMarginLine()
                tMargin.Load(MailCheck.MarginNo, "#MarginLine")
                xRS = GetRecordsetBySQL("SELECT * FROM [PO] where LeaseNo='" & tMargin.SaleNo & "' AND Style='" & tMargin.Style & "'", , GetDatabaseInventory)
                If Not xRS.EOF Then N = xRS("poid").Value
                DisposeDA(tMargin, xRS)

                'Load OrdConvertLAW

                OrdConvertLAW.SetupForm(Trim(Style), Trim(Status), Loc, N)
                'OrdConvertLAW.Show vbModal, Me
                OrdConvertLAW.ShowDialog(Me)
                R = OrdConvertLAW.Result
                L = OrdConvertLAW.Location
                CL = OrdConvertLAW.Cancelled
                'Unload OrdConvertLAW
                OrdConvertLAW.Close()
                If CL Then DisposeDA(tMargin) : Exit Sub

                MailCheck.GetMarginLine()

                Select Case R
                    Case "ST"
                        tMargin.Load(MailCheck.MarginNo, "#MarginLine")
                        ConvertPOToST(tMargin)
                End Select
                If Val(L) <> Val(Loc) Then AdjustItemLocation(L, Loc)
            Case "SS", "SO" ' Convert SS, SO to PO
                MailCheck.GetMarginLine()
                tMargin.Load(MailCheck.MarginNo, "#MarginLine")

                Dim RS As ADODB.Recordset
                RS = GetRecordsetBySQL("SELECT * FROM [PO] where LeaseNo='" & tMargin.SaleNo & "' AND Style='" & tMargin.Style & "'", , GetDatabaseInventory)
                If Not RS.EOF Then
                    '          EditPO.QuickViewPO RS("PoNo")
                    QuickShowPOForStyle(Style)
                    '        DescribePO , RS("PoID")
                Else
                    MsgBox("Could not find PoNo")
                End If

                DisposeDA(RS, tMargin)

    '      tMargin.Load MailCheck.MarginNo, "#MarginLine"
    '      ConvertSSToPO tMargin
    '
    '
    '      ' If this style is unclaimed on any POs, pop up a list and, if one is selected, change this item to that PO.
    '      ' Otherwise, allow the option to create a new PO?
    '
    '      tMargin.Load MailCheck.MarginNo, "#MarginLine"  ' Not right
    '      ConvertSTToPO tMargin
    '
    '      ShowPOs Style ' right idea, horrible side effects
    '
    '      Load OrdConvertLAW ' also not quite right
    '      OrdConvertLAW.SetupForm Trim(Style), Trim(Status), Loc
    '      OrdConvertLAW.Show vbModal, Me
    '      R = OrdConvertLAW.Result
    '      L = OrdConvertLAW.Location
    '      CL = OrdConvertLAW.Cancelled
    '      Unload OrdConvertLAW
    '      If CL Then DisposeDA tMargin: Exit Sub

            Case "LAW"
                'Load OrdConvertLAW

                OrdConvertLAW.SetupForm(Trim(Style), Trim(Status), Loc)
                'OrdConvertLAW.Show vbModal, Me
                OrdConvertLAW.ShowDialog(Me)
                R = OrdConvertLAW.Result
                L = OrdConvertLAW.Location
                CL = OrdConvertLAW.Cancelled
                'Unload OrdConvertLAW
                OrdConvertLAW.Close()
                If CL Then DisposeDA(tMargin) : Exit Sub

                MailCheck.GetMarginLine()

                Select Case R
                    Case "ST" : ConvertToStock()
                    Case "SO"
                        tMargin.Load(MailCheck.MarginNo, "#MarginLine")
                        ConvertToSO()
                        NewPoNo = MakePo(tMargin)
                        SetPoNoOnDetailLine(tMargin.Detail, NewPoNo)
                    Case "PO"
                        tMargin.Load(MailCheck.MarginNo, "#MarginLine")
                        ConvertToPO(tMargin)
                End Select
                If Val(L) <> Val(Loc) Then AdjustItemLocation(L, Loc)
            Case "SSLAW"
                If MsgBox("Convert to Special Order?", vbQuestion + vbYesNo) = vbYes Then
                    MailCheck.GetMarginLine()
                    tMargin.Load(CStr(MailCheck.MarginNo), "#MarginLine")

                    ConvertToSO()
                    NewPoNo = MakePo(tMargin)
                    SetPoNoOnDetailLine(tMargin.Detail, NewPoNo)
                End If
            Case "POREC", "SOREC"
                'Load OrdConvertLAW
                OrdConvertLAW.SetupForm(Trim(Style), Trim(Status), Loc)
                OrdConvertLAW.ShowDialog(Me)
                L = OrdConvertLAW.Location
                CL = OrdConvertLAW.Cancelled
                'Unload OrdConvertLAW
                OrdConvertLAW.Close()
                If CL Then DisposeDA(tMargin) : Exit Sub

                MailCheck.GetMarginLine()
                If Val(L) <> Val(Loc) Then AdjustItemLocation(L, Loc)
        End Select
        DisposeDA(tMargin)
    End Sub

    Private Function AdjustItemLocation(ByVal NewLoc As Integer, ByVal OldLoc As Integer)
        Dim S As String
        Dim cInv As CInvRec, Margin As CGrossMargin, Detail As CInventoryDetail

        AdjustItemLocation = Nothing
        S = Trim(Status)

        If IsIn(S, "ST", "DELTW") Then
            cInv = New CInvRec
            If Not cInv.Load(Style, "Style") Then
                MsgBox("Couldn't find this style in the database.", vbCritical, "Couldn't adjust item")
                DisposeDA(cInv)
                Exit Function
            Else
                cInv.AddLocationQuantity(OldLoc, Quan)
                cInv.AddLocationQuantity(NewLoc, -Quan)
                cInv.Save()
            End If
            DisposeDA(cInv)
        End If

        Margin = New CGrossMargin
        If Not Margin.Load(MailCheck.MarginNo, "#MarginLine") Then
            MsgBox("Couldn't load sale data.", vbCritical, "Couldn't adjust item")
            Exit Function
        Else
            Margin.Location = NewLoc
            Margin.Save()
        End If

        If IsIn(S, "ST", "SO", "DELTW") And Margin.Detail <> 0 Then
            Detail = GetDetail(Margin.Detail)
            If Detail Is Nothing Then
                MsgBox("Couldn't load detail record.", vbCritical, "Couldn't adjust item")
                DisposeDA(Detail)
            Else
                Detail.SetLocationQuantity(OldLoc, 0)
                Detail.SetLocationQuantity(NewLoc, Quan)
                Detail.Save()
            End If
            DisposeDA(Detail)
        End If

        DisposeDA(Margin)

        Loc = NewLoc
    End Function

    Private Function GetDetail(ByVal DetailID As Integer) As CInventoryDetail
        GetDetail = New CInventoryDetail
        If Not GetDetail.Load(DetailID, "#DetailID") Then DisposeDA(GetDetail)
    End Function

    Private Function ConvertPOToST(ByRef tMargin As CGrossMargin)
        Dim InvDetail As CInventoryDetail, InvRec As CInvRec
        Dim L As Integer

        On Error GoTo HandleErr

        If tMargin Is Nothing Then Exit Function

        If Trim(tMargin.Status) <> "PO" Then Exit Function
        L = Val(Loc)
        If L = 0 Then L = 1

        InvDetail = GetDetail(tMargin.Detail)
        If Not InvDetail Is Nothing Then      ' change the PO record to be a NS record
            InvDetail.Trans = "NS"
            InvDetail.SO1 = 0
            InvDetail.AmtS1 = tMargin.Quantity
            InvDetail.SetLocationQuantity(tMargin.Location, tMargin.Quantity)
            InvDetail.Save()
        End If
        InvRec = New CInvRec
        If InvRec.Load(Style, "Style") Then   ' set the PoSold in the inventory record [2data]
            InvRec.PoSold = InvRec.PoSold - Val(Quan)
            InvRec.Available = InvRec.Available - Val(Quan)
            InvRec.SetStock(L, InvRec.QueryStock(L) - Val(Quan))
            InvRec.Save()
        End If

        Status = "ST"                         ' set it on the screen

        tMargin.Status = "ST"                 ' set it in the margin record
        tMargin.Save()

        If g_Holding.Load(tMargin.SaleNo) Then  ' this will necessicate that this sale be in the open state
            g_Holding.Status = "O"
            g_Holding.Save()
        End If

        DisposeDA(InvDetail, InvRec)
        Exit Function

HandleErr:
        MsgBox("ERROR in Detail BillOSale.ConvertPOToST: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Resume Next
    End Function

    Private Function ConvertSTToPO(ByRef tMargin As CGrossMargin)
        Dim InvDetail As CInventoryDetail, InvRec As CInvRec
        Dim L As Integer

        On Error GoTo HandleErr

        If tMargin Is Nothing Then Exit Function

        If Trim(tMargin.Status) <> "ST" Then Exit Function
        L = Val(Loc)
        If L = 0 Then L = 1

        InvDetail = GetDetail(tMargin.Detail)
        If Not InvDetail Is Nothing Then      ' set it in the detail table to be a PO instead of, probably, NS
            InvDetail.Trans = "PO"
            InvDetail.AmtS1 = 0
            InvDetail.SetLocationQuantity(tMargin.Location, 0)
            InvDetail.SO1 = tMargin.Quantity
            InvDetail.Save()
        End If
        InvRec = New CInvRec
        If InvRec.Load(Style, "Style") Then   ' set the PoSold in the inventory record [2data]
            InvRec.PoSold = InvRec.PoSold + Val(Quan)
            InvRec.Available = InvRec.Available + Val(Quan)
            InvRec.SetStock(L, InvRec.QueryStock(L) + Val(Quan))
            InvRec.Save()
        End If

        Status = "PO"                         ' set it on the screen

        tMargin.Status = "PO"                 ' set it in the margin record
        tMargin.Save()

        If g_Holding.Load(tMargin.SaleNo) Then  ' this will necessitate that this sale be in the open state
            g_Holding.Status = "O"
            g_Holding.Save()
        End If

        DisposeDA(InvDetail, InvRec)
        Exit Function

HandleErr:
        MsgBox("ERROR in Detail BillOSale.ConvertSTToPO: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Resume Next
    End Function

    Private Function SetPoNoOnDetailLine(ByVal DetailID As Integer, ByVal NewPoNo As String) As Boolean
        Dim InvDetail As CInventoryDetail

        If NewPoNo = "" Then Exit Function
        InvDetail = GetDetail(DetailID)
        If Not InvDetail Is Nothing Then
            InvDetail.Name = InvDetail.Name & " " & NewPoNo
            InvDetail.Save()
        End If
        DisposeDA(InvDetail)
    End Function

    Private Function ConvertSSToPO(ByRef tMargin As CGrossMargin) As String
        ' Return PO Number
        Dim SQL As String, RS As ADODB.Recordset, SelList()
        Dim L As String
        Dim Sel As String, Res As Integer, N As Integer

        ConvertSSToPO = ""
        If Status <> "SS" Then Exit Function

        Exit Function  ' Not ready for prime time!

        SQL = ""
        SQL = SQL & "SELECT [PoNo],[PoDate],[Posted],[PrintPO],[DueDate],[Quantity] FROM [PO] "
        SQL = SQL & "WHERE [Style]='" & Style & "' "
        SQL = SQL & "AND [Name]='Stock' " ' BFH20090718
        SQL = SQL & "AND [Quantity]>=" & Quan & " " ' MJK20140106
        SQL = SQL & "AND [PrintPO]<>'V' AND [Posted]='' "
        If Loc() > 0 Then SQL = SQL & "AND Location=" & Loc() & " "
        SQL = SQL & "ORDER BY PoDate DESC"

        RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory)
        If RS.RecordCount = 0 Then
            RS = Nothing
            If MsgBox("Item " & Style & " is not on order. Create a new PO?", vbExclamation + vbYesNo, "No POs") = vbYes Then
                ' make a PO
                ConvertSSToPO = MakePo(tMargin) ' returns new PO#
                SetPoNoOnDetailLine(tMargin.Detail, ConvertSSToPO)
            End If
            Exit Function
        End If
        ReDim SelList(RS.RecordCount - 1)
        N = 0
        Do While Not RS.EOF
            If IfNullThenNilString(RS("PrintPO")) = "V" Or IfNullThenNilString(RS("PrintPO")) = "v" Then
                Status = "Void"
            ElseIf IfNullThenNilString(RS("Posted")) <> "" Then
                Status = "Received"
            Else
                Status = "Open"
            End If
            L = ArrangeString(RS("PoNo").Value, 10) & ArrangeString(DateFormat(RS("PoDate")), 12) & ArrangeString(RS("Quantity").Value, 4) & ArrangeString(Status, 6) & "Due:" & ArrangeString(DateFormat(IfNullThenNilString(RS("DueDate"))), 12)
            SelList(N) = L
            N = N + 1
            RS.MoveNext()
        Loop
        RS.Close()
        RS = Nothing

        Sel = SelectOptionArray("-- Select PO --", frmSelectOption.ESelOpts.SelOpt_List, SelList, "&Locate")
        If Sel <= 0 Then Exit Function
        Res = Microsoft.VisualBasic.Left(SelList(Val(Sel) - 1), 6)
        ConvertSSToPO = Res
    End Function

    Private Sub ConvertToSO(Optional ByVal NewPoNo As String = "")
        Dim InvDetail As CInventoryDetail
        Dim Margin As CGrossMargin
        Dim Inv As CInvRec

        On Error GoTo HandleErr

        Margin = New CGrossMargin
        If Not Margin.Load(MailCheck.MarginNo, "#MarginLine") Then
            MsgBox("Couldn't load sale data.", vbCritical)
            DisposeDA(Margin)
            Exit Sub
        End If

        Inv = New CInvRec
        Inv.Load(Margin.Style, "Style")

        If Trim(Margin.Status) = "LAW" Then
            ' so only
            InvDetail = GetDetail(Margin.Detail)
            If Not InvDetail Is Nothing Then
                InvDetail.LAW = 0                 ' BFH20050120 CHANGED FROM = "" TO = 0 because .LAW is Single, not string
                InvDetail.SO1 = Margin.Quantity
                If NewPoNo <> "" Then InvDetail.Name = InvDetail.Name & " " & NewPoNo
                InvDetail.Save()
            End If
            DisposeDA(InvDetail)
            Status = "SO"
            Margin.Status = "SO"
        ElseIf Trim(Margin.Status) = "SSLAW" Then
            Status = "SS"
            Margin.Status = "SS"
        ElseIf Trim(Margin.Status) = "ST" Then
            Inv.Available = Inv.Available + Val(Quan)
            Inv.SetStock(Val(Loc), Inv.QueryStock(Val(Loc)) + Val(Quan))
            Inv.Save()
            InvDetail = GetDetail(Margin.Detail)
            If Not InvDetail Is Nothing Then
                InvDetail.AmtS1 = 0
                '      InvDetail.SetLocationQuantity Val(Loc), 0
                InvDetail.SO1 = Margin.Quantity
                If NewPoNo <> "" Then InvDetail.Name = InvDetail.Name & " " & NewPoNo
                InvDetail.Save()
            End If
            DisposeDA(InvDetail)
            Status = "SO"
            Margin.Status = "SO"
        End If

        Margin.Save()

        If g_Holding.Load(Margin.SaleNo) Then
            g_Holding.Status = "O"
            g_Holding.Save()
        End If

        DisposeDA(Inv, Margin)
        Exit Sub

HandleErr:
        MsgBox("ERROR in Detail BillOSale.ConvertToSo: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Resume Next
    End Sub

    Private Sub ConvertToPO(ByRef tMargin As CGrossMargin)
        Dim InvDetail As CInventoryDetail, InvRec As CInvRec

        On Error GoTo HandleErr

        If tMargin Is Nothing Then Exit Sub

        If Trim(tMargin.Status) <> "LAW" Then Exit Sub

        InvDetail = GetDetail(tMargin.Detail)
        If Not InvDetail Is Nothing Then      ' set it in the detail table to be a PO instead of, probably, NS
            InvDetail.LAW = 0
            InvDetail.SO1 = tMargin.Quantity
            InvDetail.Trans = "PO"
            InvDetail.Save()
        End If
        InvRec = New CInvRec
        If InvRec.Load(Style, "Style") Then   ' set the PoSold in the inventory record [2data]
            InvRec.PoSold = InvRec.PoSold + Val(Quan)
            InvRec.Save()
        End If

        Status = "PO"                         ' set it on the screen

        tMargin.Status = "PO"                 ' set it in the margin record
        tMargin.Save()

        If g_Holding.Load(tMargin.SaleNo) Then  ' this will necessitate that this sale be in the open state
            g_Holding.Status = "O"
            g_Holding.Save()
        End If

        DisposeDA(InvDetail, InvRec)
        Exit Sub

HandleErr:
        MsgBox("ERROR in Detail BillOSale.ConvertToPO: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Resume Next
    End Sub

    Private Sub ConvertSTToDELTW()
        Dim Available As Integer
        Dim InvDetail As CInventoryDetail, Holding As cHolding
        Dim Margin As CGrossMargin, InvData As CInvRec

        On Error GoTo HandleErr

        Margin = New CGrossMargin
        If Not Margin.Load(MailCheck.MarginNo, "#MarginLine") Then
            DisposeDA(Margin)
            MsgBox("Error in ConvertSTToDELTW: Can't load Margin record #" & MailCheck.MarginNo, vbCritical, "Error")
            Exit Sub
        End If
        InvData = New CInvRec
        If Not InvData.Load(Margin.RN, "#Rn") Then
            DisposeDA(Margin, InvData)
            MsgBox("Error in ConvertSTToDELTW: Invalid Record Number.", vbCritical, "Error")
            Exit Sub
        End If

        Available = InvData.QueryStock(Margin.Location)

        If Available <= 1 Then
            If MsgBox("You are attempting to take the last item or have no stock leaving a negative balance!   Available: " & Available, vbInformation + vbOKCancel) = vbCancel Then
                DisposeDA(Margin, InvData)
                Exit Sub
            End If
        End If

        Status = "DELTW"
        Margin.Status = "DELTW"
        '  If Not IsDate(Margin.DDelDat) Then
        Margin.DDelDat = Today
        '  End If

        InvDetail = GetDetail(Margin.Detail)
        InvDetail.Trans = "DS"
        InvDetail.Save()

        InvData.OnHand = InvData.OnHand - Margin.Quantity
        If InvData.OnHand < 0 Then InvData.OnHand = 0
        InvData.Save()

        Margin.Save()

        DisposeDA(Margin, InvData, InvDetail)
        Exit Sub

HandleErr:
        MsgBox("ERROR in Detail BillOSale.ConvertSTToLAW: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Resume Next

    End Sub

    Private Sub ConvertSTToLAW()
        Dim Available As Integer
        Dim InvDetail As CInventoryDetail, Holding As cHolding
        Dim Margin As CGrossMargin, InvData As CInvRec

        On Error GoTo HandleErr

        Margin = New CGrossMargin
        If Not Margin.Load(MailCheck.MarginNo, "#MarginLine") Then
            DisposeDA(Margin)
            MsgBox("Error in ConvertSTToLAW: Can't load Margin record #" & MailCheck.MarginNo, vbCritical, "Error")
            Exit Sub
        End If
        InvData = New CInvRec
        If Not InvData.Load(Margin.RN, "#Rn") Then
            DisposeDA(Margin, InvData)
            MsgBox("Error in ConvertSTToLAW: Invalid Record Number.", vbCritical, "Error")
            Exit Sub
        End If

        Available = InvData.QueryStock(Margin.Location)

        If Available <= 1 Then
            If MsgBox("You are attempting to take the last item or have no stock leaving a negative balance!   Available: " & Available, vbInformation + vbOKCancel) = vbCancel Then
                DisposeDA(Margin, InvData)
                Exit Sub
            End If
        End If

        Status = "LAW"
        Margin.Status = "LAW"

        InvDetail = GetDetail(Margin.Detail)
        InvData.Available = InvData.Available + Margin.Quantity
        InvDetail.LAW = Margin.Quantity
        InvDetail.AmtS1 = 0
        InvDetail.Save()

        InvData.AddLocationQuantity(Margin.Location, Margin.Quantity)
        InvData.Save()

        Margin.Save()

        Holding = New cHolding
        If Holding.Load(Margin.SaleNo) Then
            Holding.Status = "L"
            Holding.Save()
        End If

        DisposeDA(Margin, InvData, InvDetail, Holding)
        Exit Sub

HandleErr:
        MsgBox("ERROR in Detail BillOSale.ConvertSTToLAW: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Resume Next
    End Sub

    Private Function MakePo(ByRef Margin As CGrossMargin) As String
        ' This doesn't seem to update InvData's OnOrder fields.
        Dim PO As New cPODetail

        On Error GoTo HandleErr
        With PO
            .PoNo = PoNo
            .SaleNo = Trim(BillOfSale.Text)
            ' bfh20050825
            ' this shouldn't be the original sale date, but rather the current date (which is when the po was made!)
            .PoDate = Today  'dteSaleDate
            .Name = CustomerLast.Text
            .Vendor = Mfg
            .InitialQuantity = GetPrice(Quan)
            .Quantity = GetPrice(Quan)
            .Style = Style
            .Desc = Trim(Desc)

            'If Trim(Margin.Status) = "SS" Then Margin.Cost = 0: Margin.ItemFreight = 0  ' This doesn't matter, Margin has already been saved.
            .Cost = Format(Margin.Cost, "Currency")
            '.Cost = Format(Margin.Cost, "decimal")

            .Location = Margin.Store  'location sold from 04-01-2002
            .SoldTo = "1"
            'BFH20060512 - Added b/c F1 wanted to make a SO sale in Store 2 or 3, selecting loc 1, and have shipto show loc 1
            If .Location <> Margin.Location And IsFurnOne() Then
                .ShipTo = "3"
                .ShiptoName = StoreSettings(Margin.Location).Name
                .ShipToAddress = StoreSettings(Margin.Location).Address
                .ShipToCity = StoreSettings(Margin.Location).City
                .ShipToTele = StoreSettings(Margin.Location).Phone
            Else
                .ShipTo = "2"  '04-01-2002 SHOULD BE DEFAULT LOCATION
            End If

            If StoreSettings.bPOSpecialInstr Then
                .Note1 = "1"
                .Note2 = "1"
            Else
                .Note1 = "0"
                .Note2 = "0"
            End If

            .Note3 = "0"
            .Note4 = "0"
            .PoNotes = ""
            .AckInv = ""
            .Posted = ""
            .PrintPo = ""
            .wCost = "1" ' Print w/Cost
            If StoreSettings.bPrintPoNoCost Then .wCost = "0"
            .RN = Margin.RN               'added margin.. 11-07-01
            .Detail = Margin.Detail       'added margin.. 11-07-01

            ' .MarginLine will be empty for Stock orders.
            If OrderMode("A") Then   'changed 11-07-01
                .MarginLine = MarginNo
            Else
                .MarginLine = MailCheck.MarginNo
            End If
        End With
        PO.Save()
        MakePo = PO.PoNo

        LastSale = Trim(BillOfSale.Text)
        LastMfg = Trim(Mfg)
        DisposeDA(PO)

        Exit Function

HandleErr:
        Resume Next
    End Function

    Private Function GetMaxItemIndex() As Integer
        Dim I As Integer, X As Integer

        X = 0
        'For I = 0 To UGridIO1.MaxRows - 1
        '    If Not IsInventoryItemComplete(I) Then Exit For
        '    X = X + ((Len(QueryDesc(I)) - 1) \ 46 + 1)
        'Next
        GetMaxItemIndex = X
    End Function

    Private Function GetMaxPages(ByVal Items As Integer) As Integer
        GetMaxPages = (Items \ 17) + 1
    End Function

    Private Sub GetLinePart(ByVal Page As Integer, ByVal Line As Integer, ByRef Item As Integer, ByRef ItemLine As Integer)
        Dim T As Integer, U As Integer
        Dim N As Integer, X As Integer, P As Integer, F As Integer

        X = 0
        'For T = 0 To UGridIO1.MaxRows - 1
        '    N = NumLineBreaks(QueryDesc(T))
        '    For U = 1 To N
        '        X = X + 1

        '        P = (X \ 17)
        '        F = (X Mod 17)
        '        If Page = P And F > Line Then
        '            Item = T
        '            ItemLine = U - 1
        '            Exit Sub
        '        ElseIf Page = P - 1 And Line = 16 And F = 0 Then
        '            Item = T
        '            ItemLine = U - 1
        '            Exit Sub
        '        End If
        '    Next
        'Next
        MsgBox("Could not Match Item Line.  Page " & (Page + 1) & ", Line " & (Line + 1), vbCritical, "Invoice Printing Error")
    End Sub

    Private Sub ConvertToStock()
        Dim Available As Integer
        Dim InvDetail As CInventoryDetail, Margin As CGrossMargin
        Dim InvData As CInvRec
        Dim Holding As cHolding

        On Error GoTo HandleErr

        Margin = New CGrossMargin
        If Not Margin.Load(MailCheck.MarginNo, "#MarginLine") Then
            MsgBox("Error in ConvertToStock: Can't load Margin record #" & MailCheck.MarginNo, vbCritical, "Error")
            DisposeDA(Margin)
            Exit Sub
        End If

        InvData = New CInvRec
        If Not InvData.Load(Margin.RN, "#Rn") Then
            MsgBox("Error in ConvertToStock: Invalid Record Number.", vbCritical, "Error")
            DisposeDA(Margin, InvData)
            Exit Sub
        End If

        Available = InvData.QueryStock(Margin.Location)

        If Available <= 1 Then
            If MsgBox("You are attempting to take the last item or have no stock leaving a negative balance!" & vbCrLf & "Available: " & Available, vbInformation + vbOKCancel) = vbCancel Then
                DisposeDA(Margin, InvData)
                Exit Sub
            End If
        End If

        Status = "ST"
        Margin.Status = "ST"

        InvDetail = GetDetail(Margin.Detail)
        InvData.Available = InvData.Available - Margin.Quantity
        InvDetail.LAW = 0
        InvDetail.AmtS1 = Margin.Quantity
        InvDetail.Save()
        DisposeDA(InvDetail)

        InvData.AddLocationQuantity(Margin.Location, -Margin.Quantity)
        InvData.Save()
        DisposeDA(InvData)

        Margin.Save()

        Holding = New cHolding
        If Holding.Load(Margin.SaleNo) Then
            Holding.Status = "O"
            Holding.Save()
        End If

        DisposeDA(Margin, Holding)  ' Margin used for loading holding
        Exit Sub

HandleErr:
        MsgBox("ERROR in Detail BillOSale.ConvertToStock: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Resume Next
    End Sub

    Public Property CurrentLine() As Integer
        Get
            CurrentLine = mCurrentLine
        End Get
        Set(value As Integer)
            mCurrentLine = value
        End Set
    End Property

    Public Sub CorrectMailDatabase()
        ' ALL SALES COME HERE TO CORRECT OR ADD NEW
        On Error GoTo HandleErr

        If MailCheck.NameFound = "" Then MailCheck.NameFound = True

        ' This is causing double Mail entries.
        ' The mail record is saved when we click OK on BillOSale,
        ' then again on cmdProcessSale.
        ' MailCheck is unloaded after loading   When this function reloads it, .Customer=New.
        ' To fix, we've made the tail of this function set MailCheck.Customer="Old" and NameFound=True.
        If Index > 0 Then GoTo ExitHere

        If (MailCheck.Customer = "New") Or (Index < 1 And (OrderMode("A", "B", "E"))) Or ((MailMode("ADD/Edit")) And (MailCheck.NameFound = False)) Then
            ' new customer

            '    VerifyMailRecUnique MailRec, vbTrue  ' clear it if it was saved..
            '
            '    Dim Extra as integer
            '    Extra = 1
            '    Do
            '      MailRec = MailTableRecordMax("Index") + Extra
            '      Extra = Extra + 1
            '    Loop While Not VerifyMailRecUnique(MailRec)
            '    VerifyMailRecUnique MailRec, vbFalse ' add our new record
            MailRec = MailTableRecordMax("Index") + 1
            Index = MailRec
            modMail.MailRec = MailRec    ' MJK 20030701 to handle add+edit without reload
        End If

ExitHere:

        ' If MailCheck.Customer = "Old" Then
        MailRec = Index ' modMail.MailRec
        '  Debug.Print "MailRec=" & MailRec
        ' End If
        ' If mailmode( "ADD/Edit") Then
        '     MailRec = Index
        ' End If

        ' REprint with corrections if any OR adds new customer to data Base
        '         Mail.Index = Trim(MailCheck.Index)
        Mail.Index = Trim(Index)
        Mail.Last = Trim(CustomerLast.Text)
        Mail.First = Trim(CustomerFirst.Text)
        Mail.Address = Trim(CustomerAddress.Text)
        Mail.AddAddress = AddAddress.Text
        Mail.City = Trim(CustomerCity.Text)
        Mail.Zip = Trim(CustomerZip.Text)
        Mail.Tele = CleanAni(CustomerPhone1.Text)
        Mail.Tele2 = CleanAni(CustomerPhone2.Text)
        Mail.PhoneLabel1 = cboPhone1.Text
        Mail.PhoneLabel2 = cboPhone2.Text
        Mail.Special = txtSpecInst.Text
        Mail.Type = cboCustType.SelectedIndex     'sets "-" in data base
        If cboAdvertisingType.SelectedIndex = -1 Then
            Mail.CustType = 0
        Else
            'Mail.CustType = cboAdvertisingType.itemData(cboAdvertisingType.SelectedIndex)
            'NOTE: THE ABOVE LINE WILL BE REPLACED WITH THE BELOW THREE LINES TO GET THE ITEMDATA VALUE USING ITEMDATACLASS
            'CUSTOM CLASS.
            Dim idc As ItemDataClass
            idc = cboAdvertisingType.Items(cboAdvertisingType.SelectedIndex)
            Mail.CustType = idc.ItemData
        End If

        Mail.Email = Email.Text
        If cboCustType.SelectedIndex = -1 Then Mail.Type = "0"
        Mail.Business = optBusiness.Checked
        Mail.TaxZone = cboTaxZone.SelectedIndex + 1
        ' Mail.Blank
        Dim RS As ADODB.Recordset
        Dim RS2 As ADODB.Recordset
        RS = getRecordsetByTableLabelIndexNumber("Mail", "Index", CStr(MailRec), True)
        SetMailRecordsetFromMailNew(RS, Mail)

        Typpe = Mail.Type
        ' Ship to address
        If Trim(CustomerAddress2.Text) <> "" Or Trim(ShipToFirst.Text) <> "" Or Trim(ShipToLast.Text) <> "" Or CleanAni(CustomerPhone3.Text) <> "" Then
            Dim Mail2 As MailNew2
            If Val(Mail.Index) <> 0 Then
                ' Save changes if any
                'Dim RS2 As ADODB.Recordset
                RS2 = getRecordsetByTableLabelIndexNumber("MailShipTo", "Index", CStr(Mail.Index), True)
                CopyMailRecordsetToMailNew2(RS2, Mail2)
                Mail2.Index = Mail.Index
                Mail2.ShipToLast = ShipToLast.Text
                Mail2.ShipToFirst = ShipToFirst.Text
                Mail2.Address2 = CustomerAddress2.Text
                Mail2.City2 = CustomerCity2.Text
                Mail2.Zip2 = CustomerZip2.Text
                Mail2.Tele3 = CleanAni(CustomerPhone3.Text)
                Mail2.PhoneLabel3 = Trim(cboPhone3.Text)
                'Mail2.Blank =
            Else
                ' Just added address #2
                Mail2.Index = Mail.Index
                Mail2.ShipToLast = ShipToLast.Text
                Mail2.ShipToFirst = ShipToFirst.Text
                Mail2.Address2 = CustomerAddress2.Text
                Mail2.City2 = CustomerCity2.Text
                Mail2.Zip2 = CustomerZip2.Text
                Mail2.Tele3 = CleanAni(CustomerPhone3.Text)
                Mail2.PhoneLabel3 = Trim(cboPhone3.Text)
                'Mail2.Blank =
            End If
            SetMailRecordsetFromMailNew2(RS2, Mail2)
            SetMailRecordsetByTableLabelIndex(RS2, "MailShipTo", "Index", CStr(Mail.Index))
        Else
            ExecuteRecordsetBySQL("DELETE * FROM MAILSHIPTO WHERE INDEX=" & MailRec, , GetDatabaseAtLocation())
        End If
        SetMailRecordset(RS)

        ' If we're updating an existing order, update corresponding records also.
        If UCase(MailCheck.Customer) = "OLD" Then
            If g_Holding.Index > 1 Then
                ' We're updating an existing mail record.
                ' This means we may have to update ArApp, InstallmentInfo, Service, Transactions.
            End If
            'Always update Holding, GM, Audit, Detail, PO.
            MailRecordUpdated(BillOfSale.Text, Mail)
        End If

        'check for installment customer
        'If Trim(MailCheck.OldTele) <> "" Then
        'If frmSetup .Installment = "Y" Then
        '  tid = Trim(MailCheck.OldTele)
        MailCheck.Customer = "Old"
        MailCheck.NameFound = True

        Exit Sub
HandleErr:
        '  MsgBox "Error in Mailing List Update [" & Err.Number & ":" & vbCrLf & Err.Description, vbExclamation
        Resume Next
    End Sub

    Public Sub QuickShowSaleTicket(ByVal SaleNo As String, Optional ByVal StoreNo As Integer = 0, Optional ByVal ReturnToOriginalStore As Boolean = False)
        Dim OldMMOrder As String, OldStoreNo As Integer

        If StoreNo = 0 Then StoreNo = StoresSld
        OldStoreNo = StoresSld

        If SaleNo <> "" Then
            If StoreNo <> StoresSld Then
                ' Sale was in a different store.  We have to switch to view it.
                StoresSld = StoreNo
                If Not ReturnToOriginalStore Then
                    MsgBox("This sale was made in store " & StoreNo & "." & vbCrLf &
               "Your current login has been changed to store " & StoreNo & "." & vbCrLf &
               "You may want to change it back before making new sales.",
               vbExclamation, "Current Store Changed")
                Else
                    MsgBox("This sale was made in store " & StoreNo & ", not in your current store (store " & OldStoreNo & ")" & vbCrLf &
               "Please note that you must log into the correct store to view this sale normally.",
               vbExclamation, "Sale store different than login store")
                End If
            End If

            ' This displays the Inventory Detail's matching customer record in a disabled BillOSale.  No edits allowed.
            OldMMOrder = Order
            Order = "E"
            MailCheck.optSaleNo.Checked = True
            MailCheck.InputBox.Text = SaleNo
            'MailCheck.cmdOK.Value = True
            MailCheck.cmdOK.PerformClick()
            Show()
            cmdApplyBillOSale.Enabled = False
            cmdCancel.Enabled = False

            BillOSale2_Show()
            cmdClear.Enabled = False
            cmdNextSale.Enabled = False
            cmdProcessSale.Enabled = False
            ScanDn.Enabled = False
            ScanUp123.Enabled = False
            'UGridIO1.GetDBGrid.AllowUpdate = False
            cmdMainMenu.Text = "Back"
            '  BFH20160130
            '  Because this whole feature was completely added on, and was not thought out...
            '  There are certain features which actually still needed the store #...
            '  If we check the caption of the MainMenu button to be "Back", as above,
            '  then we can rely on the tag to be the store number...
            '  Such as it is in old programs like this..  We simply "make it work".
            '  Currently the features requiring/using this are:
            '    - The Print mini-button.
            '    - The Email mini-button.
            cmdMainMenu.Tag = StoresSld
            Order = OldMMOrder
        End If

        If ReturnToOriginalStore And OldStoreNo <> StoresSld Then
            StoresSld = OldStoreNo
            '    main_StoreChange OldStoreNo
            '    frmSetup .LoadStore
        End If
    End Sub

    Private Sub MailRecordUpdated(ByVal SaleNo As String, ByRef Mail As MailNew)
        Dim tHol As New cHolding

        If Trim(SaleNo) = "" Then Exit Sub

        If tHol.Load(SaleNo, "LeaseNo") Then
            tHol.Index = Mail.Index
            tHol.Save()
        End If
        DisposeDA(tHol)

        Dim tGM As New CGrossMargin
        tGM.Load(SaleNo, "SaleNo")
        Do While Not tGM.DataAccess.Record_EOF
            tGM.Name = Trim(Mail.Last)
            tGM.Index = Mail.Index
            tGM.Save()
            tGM.DataAccess.Records_MoveNext()
        Loop
        DisposeDA(tGM)

        ExecuteRecordsetBySQL("UPDATE [Detail] SET Name='" & Mail.Last & "' WHERE SaleNo='" & SaleNo & "' AND Store=" & StoresSld, , GetDatabaseInventory)
        ' BFH20050503 - This too was wrong...  Can't look at anything in the invent DB just by sale no
        '  Dim tDet As New CInventoryDetail
        '  tDet.Load SaleNo, "SaleNo"
        '  Do While Not tDet.DataAccess.Record_EOF
        '    tDet.Name = Trim(Mail.Last)
        '    tDet.Save
        '    tDet.DataAccess.Records_MoveNext
        '  Loop
        '  Set tDet = Nothing

        ExecuteRecordsetBySQL("UPDATE [PO] SET Name='" & Mail.Last & "' WHERE LeaseNo='" & SaleNo & "' AND Location=" & StoresSld, , GetDatabaseInventory)
        ' BFH20050313 - This is wrong..
        '   Lease No is not unique in the PO table...
        '   PO is in the invent database, so LeaseNo is only valid with location!!!!
        '  Dim tPO As New cPODetail
        '  tPO.Load SaleNo, "LeaseNo"
        '  Do While Not tPO.DataAccess.Record_EOF
        '    tPO.Name = Trim(Mail.Last)
        '    tPO.Save
        '    tPO.DataAccess.Records_MoveNext
        '  Loop
        '  Set tPO = Nothing
    End Sub

    Private Sub LoadSalesSplitBoxes()
        Dim N As Integer
        Dim L As Object
        Dim A() As Object = {SalesSplit1, SalesSplit2, SalesSplit3}

        'For Each L In Array(SalesSplit1, SalesSplit2, SalesSplit3)
        For Each L In A
            L.items.Clear
            'L.AddItem("0%")
            L.items.add("0%")

            For N = 5 To 30 Step 5 : L.items.add(AlignString(N, 3) & "%") : Next
            'L.AddItem("33.33%")
            L.items.add("33.33%")
            For N = 35 To 45 Step 5 : L.items.add(AlignString(N, 3) & "%") : Next
            'L.AddItem("50%")
            L.items.add("50%")
            For N = 55 To 100 Step 5 : L.items.add(AlignString(N, 3) & "%") : Next
            'L.ListIndex = 0
            L.selectedindex = 0
        Next
    End Sub

    Private Sub UGridIO1_RowDelete(LastRow As Object) Handles UGridIO1.RowDelete
        If UGridIO1.Row < NewStyleLine Then
            NewStyleLine = NewStyleLine - 1
        End If
        ' ReCalculate  ' For some reason, calling this here adds a price to the second row of a kit.
        ' We'll recalculate in the AfterDelete event instead.
        If Trim(QueryStyle(X)) = "" Then StyleAddBegin(X) ' If there's nothing in the current style cell, force a selection.
    End Sub

    Private Sub StyleAddBegin(ByVal LineNo As Integer)
        '  NewStyleLine = LineNo  ' This should automatically increment on insert, decrement on delete, and reset on clear.
        '  Debug.Print "Starting style addition on " & NewStyleLine
        CheckSplitKits()
        OrdSelect.Show()
        AddingItem = True
    End Sub

    Private Sub UGridIO1RowColChange(LastRow As Object, LastCol As Object, newRow As Object, newCol As Object, ByRef Cancel As Boolean) Handles UGridIO1.RowColChange
        Dim Style As String

        If UGridIO1.Loading = True Then Exit Sub

        'If IsEmpty(LastRow) And IsEmpty(LastCol) And IsEmpty(newRow) And IsEmpty(newCol) Then Exit Sub
        If IsNothing(LastRow) And IsNothing(LastCol) And IsNothing(newRow) And IsNothing(newCol) Then Exit Sub
        'If newRow = -1 Then Exit Sub

        If newRow = -1 Then newRow = 0
        If Order <> "A" Then Exit Sub

        Style = QueryStyle(newRow)

        If Not IsNothing(LastRow) And Not IsNothing(newRow) Then
            If Val(LastRow) = Val(newRow) Then UGridIO1.ForceRowSave()
        End If

        X = UGridIO1.Row
        If Not IsNothing(newRow) Then If (LastRow = Str(newRow)) And (LastCol = newCol) Then Exit Sub
        'Debug.Print "UGridIO1_RowColChange", LastRow, LastCol, newRow, newCol, UGridIO1.LostFocusFlag
        'If Not IsNull(LastRow) Then
        Cancel = False

        Select Case LastCol
            Case BillColumns.eStyle
    '      Cancel = Not Style_LostFocus(CInt(LastRow))
    '    Case BillColumns.eDescription
    '      Cancel = Not Desc_LostFocus(CInt(LastRow))
    '    Case BillColumns.eLoc
    '      Cancel = Not Loc_LostFocus(CInt(LastRow))
            Case BillColumns.ePrice
                Cancel = Not Price_LostFocus(LastRow)
            Case BillColumns.eQuant
                Cancel = Not Quant_LostFocus(LastRow)
        End Select
        Cancel = False
        'If Cancel Then Exit Sub  ' This causes the popup form to not show when tabbing from price to style.
        ' End If


        Select Case UGridIO1.Col
            Case BillColumns.eStyle
                ' Problem: InvDefault may be unloaded at the time this is called.
                ' It is unknown whether it may ever be loaded at this time.
                ' On ReEnter, at any rate, OrdSelect is unloaded as quickly as it's shown.
                ' Or is it?  Could be just hidden behind another, more up-front, form.
                Dim ReEnt As Boolean
                If IsFormLoaded("InvDefault") Then
                    ReEnt = InvDefault.optReEnter.Checked
                Else
                    ReEnt = True
                End If
                '      If IsFormLoaded("InvDefault") Then
                If ReEnt Then
                    If Not PrintBill Then     ' prevents form showing after sale is processed
                        '          If Not IsInventoryItemComplete(NewStyleLine) Then
                        '            ' We had a partially entered line, then switched focus?
                        '            ' Problem: IsInventoryItemComplete only checks that we have a style number.
                        '          End If
                        If MaxLines - LinesUsed() <= LinePadRequired() Then
                            MsgBox("This sale has almost reached the maximum number of lines." & vbCrLf &
                                     "You can add tax if required and complete the sale." & vbCrLf &
                                     "Then you can enter a new sale for the balance of the merchandise.")
                        End If
                        CheckSplitKits()

                        If AddingItem Then
                            OrdSelect.ShowToBillOSale2()
                        Else
                            'Unload OrdSelect          ' fixes disabled Process button.
                            OrdSelect.Close()
                            StyleAddBegin(CurrentLine)
                        End If
                    End If
                Else
                End If
        '      End If
            Case BillColumns.eDescription
            Case BillColumns.eLoc
            Case BillColumns.ePrice
        '      EnableGridPrice '  -- Disabled until we decide how we want it to go.
            Case BillColumns.eQuant
                LastQuant = Val(UGridIO1.GetValue(newRow, newCol))
        End Select

        'Commented out MJK20030918 to prevent item overwrites.
        '  If newRow <> NewStyleLine Then
        '    If OrdSelect.RowChangeOK Then NewStyleLine = newRow
        'End If
        LastGridTextAlt = ""

        If UGridIO1.GetDBGrid.Row <= 3 Then
            If UGridIO1.Row >= 19 Then
                If UGridIO1.Row >= UGridIO1.LastRowUsed Then
                    Dim T As Integer
                    T = UGridIO1.Row
                    UGridIO1.MoveRowUp(9)
                End If
            End If
        End If

        If modStores.SecurityLevel <> ComputerSecurityLevels.seclevNoPasswords Then
            If CheckAccess("Prevent Price Adjust", , True) Then
                If IsPayment(Style) Or IsNote(Style) Or IsDLS(Style) Then
                    PriceEnabled = True
                Else
                    PriceEnabled = False
                End If
            Else
                PriceEnabled = True
            End If
        End If

        'BFH20170718 - Per-line control of fields (Order="A" only)...
        If IsItem(Style) Or IsNote(Style) Or IsDLS(Style) Or IsDiscount(Style) Then
            QuanEnabled = True
            DescEnabled = True
            MfgEnabled = True
            LocEnabled = True
        Else
            QuanEnabled = False
            DescEnabled = False
            MfgEnabled = False
            LocEnabled = False
        End If
    End Sub

    Public Function Price_LostFocus(ByVal Index As Integer) As Boolean
        Dim Col As Integer, Row As Integer

        Col = UGridIO1.Col
        Row = UGridIO1.Row

        Recalculate()  'checks for wrong price
        UGridIO1.Row = Row
        UGridIO1.Col = Col

        DeleteLine = ""
    End Function

    Public Function Quant_LostFocus(ByVal Index As Integer) As Boolean
        Dim Col As Integer, Row As Integer
        Dim OldValue As String, NewVal As String

        'With UGridIO1
        '    Col = .Col
        '    Row = .Row

        '    OldValue = LastQuant
        '    NewVal = LastGridText(Index, 4) '.GetValue(Index, 4)
        '    If NewVal = "" Then Exit Function
        '    If Trim(.GetValue(Index, 3)) = "SS" And Val(OldValue) <> 0 Then
        '        'Issue# 126@Mantis. Below two lines are commented, to rectify the wrong calculation of qty * price of SS items.
        '        'NP = GetPrice(.GetValue(Index, 6)) / Val(OldValue) * Val(NewVal)
        '        '.SetValue Index, 6, Format(NP, "###,###.00")
        '        .Refresh(True)
        '        .Row = Row
        '    End If

        '    .Row = Row
        '    .Col = Col
        'End With

        DeleteLine = ""
    End Function

    Private Sub vAdjustSalesSplits()
        'Dim A As Object, B As Object, C As Object
        Dim A As String, B As String, C As String

        A = SalesSplit1.Text
        B = SalesSplit2.Text
        C = SalesSplit3.Text

        AdjustSalesSplits(A, B, C, SplitCount(Sales1, Sales2, Sales3))

        SalesSplit1.Text = A
        SalesSplit2.Text = B
        SalesSplit3.Text = C
    End Sub

    Public Function vGetSalesCode() As String
        vGetSalesCode = Trim(getSalesNumber(Sales1.Text, "99") & " " & getSalesNumber(Sales2.Text) & " " & getSalesNumber(Sales3.Text))
    End Function

    Public Function vGetSalesSplit() As String
        vAdjustSalesSplits()
        vGetSalesSplit = GetSalesSplit(SalesSplit1.Text, SalesSplit2.Text, SalesSplit3.Text, SplitCount(Sales1, Sales2, Sales3))
    End Function

    Public Function LoadSplitsToBoxes(ByVal F As String, ByVal Count As Integer)
        Dim A As Double, B As Double, C As Double

        ParseSalesSplit(F, A, B, C, Count)
        SalesSplit1.Text = "" & A & "%"
        SalesSplit2.Text = "" & B & "%"
        SalesSplit3.Text = "" & C & "%"
        vAdjustSalesSplits()
    End Function

    Private Sub HoverPic(Optional ByVal Show As Boolean = False, Optional ByVal Style As String = "", Optional ByVal Col As Integer = 0, Optional ByVal Row As Integer = 0)
        Dim RN As Integer, F As String, X As Integer, Y As Integer

        'Debug.Print "HoverPic(Show=" & Show & ", Style=" & Style & ", ....)  "
        If Not Show Then fraHover.Visible = False : Exit Sub
        RN = GetRNByStyle(Style)
        If RN = 0 Then HoverPic() : Exit Sub
        F = ItemPXByRN(RN)
        If F = "" Or Dir(F) = "" Then
            'Debug.Print "HoverPic -- File not found: " & F
            HoverPic()
            Exit Sub
        End If

        'picHover.Picture = LoadPictureStd(F)
        picHover.Image = Image.FromFile(F)
        MaintainPictureRatio(picHover, 3000, 3000, True)
        fraHover.Text = Style
        fraHover.Width = picHover.Width + 2 * picHover.Left
        fraHover.Height = picHover.Height + 1.5 * picHover.Top
        'X = UGridIO1.Left + UGridIO1.ColLeft(Col) + 1500 ' + UGridIO1.GetDBGrid.Columns(0).Width * 1.25
        'Y = UGridIO1.Top + UGridIO1.RowTop(Row) + 250 '+ UGridIO1.GetDBGrid.RowHeight * 1.5
        'Debug.Print "X=" & X & ", y=" & Y
        'fraHover.Move(X, Y)
        fraHover.Location = New Point(X, Y)
        fraHover.Visible = True
    End Sub

    Private Sub HoverTimer(Optional ByVal Start As Boolean = False)
        'Debug.Print "HoverTimer(" & Start & ")"
        HoverPic()
        tmrHover.Enabled = False
        If Start Then
            'tmrHover.Interval = 1500
            'tmrHover.Enabled = True
        Else
            tmrHover.Tag = ""
        End If
    End Sub

    Private Sub FormatHelper(ByVal Text As String, Optional ByVal Row As Integer = -1)
        'Debug.Print "." & GetTickCount
        'If Row < 0 Then Row = UGridIO1.Row
        If Len(Text) <= 46 Then
            txtFormatHelper.Visible = False
            picFormatHelper.Visible = False
            txtFormatHelper.Text = ""
            txtFormatHelper.Left = -10000
            '    UGridIO1.Refresh
        Else
            If txtFormatHelper.Text = Text And txtFormatHelper.Visible Then Exit Sub
            txtFormatHelper.BackColor = Color.White
            txtFormatHelper.Text = WrapLongText(Text, 46, , False)
            'txtFormatHelper.Top = fraBOS2.Top + UGridIO1.RowTop(Row) - 120 - txtFormatHelper.Height
            'txtFormatHelper.Left = fraBOS2.Left + UGridIO1.ColLeft(BillColumns.eDescription) + 360
            txtFormatHelper.Visible = True
            'picFormatHelper.Move txtFormatHelper.Left + 40, txtFormatHelper.Top + 40, txtFormatHelper.Width, txtFormatHelper.Height
            picFormatHelper.Location = New Point(txtFormatHelper.Left + 40, txtFormatHelper.Top + 40)
            picFormatHelper.Size = New Size(txtFormatHelper.Width, txtFormatHelper.Height)
            picFormatHelper.Visible = True

            'picFormatHelper.ZOrder(0)
            'txtFormatHelper.ZOrder(0)
            picFormatHelper.BringToFront()
            txtFormatHelper.BringToFront()

            picFormatHelper.Parent.Controls.SetChildIndex(picFormatHelper, 0)
            txtFormatHelper.Parent.Controls.SetChildIndex(txtFormatHelper, 0)
            'UGridIO1.Refresh
            txtFormatHelper.Refresh()
            picFormatHelper.Refresh()
        End If
    End Sub

    Private Sub Desc_Change(ByVal Index As Integer)
        PriceEnabled = True
        Desc = UCase(Desc)
    End Sub

    Public Sub LoadStyle(ByVal iStyle As String)
        ' Search for Style=iStyle, Set Rn, Call GetRec.
        ' Called by OrdSelect.minvCkStyle_OKClicked
        Dim SearchObj As New CSearchNew

        With SearchObj
            If .Load(Trim(iStyle)) Then
                RN = .RN
                LoadInvDataByRN(.RN)
            End If
        End With
        SearchObj.Dispose()
        SearchObj = Nothing
    End Sub

    Private Sub LoadInvDataByRN(ByVal RN As Integer)
        Dim InvData As New CInvRec, Str As Integer

        On Error GoTo HandleErr

        If Not InvData.Load(CStr(RN), "#Rn") Then
            MsgBox("Could not locate item # " & RN & ".", vbExclamation, "Error!")
        Else

            'Mfg = InvData.Vendor
            SetMfg(X, InvData.Vendor)
            SetMfgNo(X, InvData.VendorNo)
            'If (Len(Desc) = 0) Then Desc = Trim(InvData.Desc)
            If Len(Trim(QueryDesc(X))) = 0 Then SetDesc(X, Trim(InvData.Desc))

            If True Then '  Not a kit!
                If cboCustType.Text = "Customer" Or cboCustType.Text = "" Then
                    Price = Format(InvData.OnSale, "###,###.00")
                ElseIf cboCustType.Text = "Resale 1" Then
                    Price = Format((InvData.Landed * 1.1), "###,###.00")
                ElseIf cboCustType.Text = "Resale 2" Then
                    Price = Format((InvData.Landed * 1.15), "###,###.00")
                ElseIf cboCustType.Text = "Resale 3" Then
                    Price = Format((InvData.Landed * 1.2), "###,###.00")
                ElseIf cboCustType.Text = "Resale 4" Then
                    Price = Format((InvData.Landed * 1.25), "###,###.00")
                ElseIf cboCustType.Text = "Resale 5" Then
                    Price = Format((InvData.Landed * 1.3), "###,###.00")
                ElseIf cboCustType.Text = "Resale 6" Then
                    Price = Format((InvData.Landed * 1.35), "###,###.00")

                ElseIf cboCustType.Text = "Bear Club 10" Then
                    Price = Format(InvData.List * 0.9, "###,##0.00")
                ElseIf cboCustType.Text = "Bear Club 20" Then
                    Price = Format(InvData.List * 0.8, "###,##0.00")
                End If
            End If

            Rb = InvData.Available
            For Str = 1 To Setup_MaxStores_DB
                SetBalance(Str, InvData.QueryStock(Str))
                SetOnOrder(Str, InvData.QueryOnOrder(Str))
            Next
            PoSold = InvData.PoSold
        End If
        DisposeDA(InvData)
        Exit Sub

HandleErr:
        Resume Next
    End Sub

    Public Property PriceEnabled() As Boolean
        Get
            'PriceEnabled = UGridIO1.GetColumn(BillColumns.ePrice).Locked
        End Get
        Set(value As Boolean)
            'UGridIO1.GetColumn(BillColumns.ePrice).Locked = Not value
        End Set
    End Property

    Private Property Desc() As String
        Get
            Desc = QueryDesc(CurrentLines)
        End Get
        Set(value As String)
            SetDesc(CurrentLines, value, False)
        End Set
    End Property

    Public Property CurrentLines() As Integer
        Get
            CurrentLines = mCurrentLine
        End Get
        Set(value As Integer)
            mCurrentLine = value
        End Set
    End Property

    Private Sub FormatTimer(Optional ByVal Start As Boolean = False, Optional ByVal Row As Integer = 0)
        'Debug.Print "HoverTimer(" & Start & ")"
        If Not Start Then
            FormatHelper("")
            tmrFormat.Enabled = False
            tmrFormat.Tag = ""
            Exit Sub
        End If

        If tmrFormat.Tag = "" & Row Then Exit Sub
        tmrFormat.Enabled = False

        'tmrFormat.Interval = 100
        'tmrFormat.Enabled = True
        tmrFormat.Tag = Row
        '  Debug.Print "Row=" & Row
    End Sub

    Private Function IsInventoryItemComplete(ByVal Index As Integer) As Boolean
        IsInventoryItemComplete = False
        If QueryStyle(Index) <> "" Then IsInventoryItemComplete = True
    End Function

    Private Sub mDBNotes_Init()
        mDBNotes = New CDbAccessGeneral
        mDBNotes.dbOpen(GetDatabaseAtLocation())
    End Sub

    Private Sub mDBNotes_GetRecordNotFound() Handles mDBNotes.GetRecordNotFound
        NotesInfo = ""
    End Sub

    Private Sub mDBNotes_SetRecordEvent(RS As ADODB.Recordset) Handles mDBNotes.SetRecordEvent
        On Error Resume Next
        RS("BillOSale").Value = Trim(BillOfSale.Text)
        RS("Notes").Value = NotesInfo
        RS("NoteDate").Value = Now
    End Sub

    Public Sub mDBNotes_SqlSet(ByVal T)
        mDBNotes.SQL =
        "SELECT SaleNotes.*" _
        & " From SaleNotes" _
        & " WHERE SaleNotes.BillOSale = """ & ProtectSQL(T) & """ ORDER BY" _
        & " NoteDate DESC"
    End Sub

    Public Sub HandleRecentNotes()
        Dim Ns As Boolean, NA As Boolean

        If modProgramState.Order = "" And modProgramState.Inven = "" Then Exit Sub
        Ns = AccountHasRecentSaleNotes(LeaseNo)
        NA = AccountHasRecentARNotes(Val(Index))
        If Ns And NA Then
            MsgBox("This account has recent AR and Sales Notes.", vbInformation)
        ElseIf Ns Then
            MsgBox("This account has recent Sales Notes.", vbInformation)
        ElseIf NA Then
            MsgBox("This account has recent AR Notes.", vbInformation)
        Else
            ' No recent notes.
        End If
    End Sub

    Public Function SetLeaseNo(ByVal NewLease As String) As Boolean
        SetLeaseNo = False

        If Trim(NewLease) = "" Then Exit Function
        If Len(NewLease) > 16 Then Exit Function
        If Asc(NewLease) = 0 Then Exit Function
        LeaseNo = NewLease
        txtSaleNo.Text = NewLease
        BillOfSale.Text = NewLease
        SetLeaseNo = True
    End Function

    ' BFH20050121 MODIFIED, added NoPO for OrdStatus checks to selectively have POs included (tag incoming from stock option)
    ' NoPO values:  0 = show all, 1 = show w/o POs, -1 = show only POs
    Public Function ItemsSoldOnSale(ByVal Style As String, Optional ByVal Store As Integer = 0, Optional ByVal NoPO As Integer = 0) As Double
        Dim I As Integer, IsPO As Boolean, IsLAW As Boolean, IsSO As Boolean

        For I = 0 To NewStyleLine
            If Trim(QueryStyle(I)) = Style And (QueryLoc(I) = Store Or Store = 0) Then
                IsPO = (QueryStatus(I) = "PO")
                IsLAW = (QueryStatus(I) = "LAW")
                IsSO = (QueryStatus(I) = "SO")
                If NoPO = 0 Or (NoPO > 0 And Not IsPO) Or (NoPO < 0 And IsPO) Then
                    If Not IsLAW And Not IsSO Then
                        ItemsSoldOnSale = ItemsSoldOnSale + Val(QueryQuan(I))
                    End If
                End If
            End If
        Next
    End Function

    Public Function QueryLoc(ByVal RowNum As Integer) As Integer
        QueryLoc = Val(QueryGridField(RowNum, BillColumns.eLoc))
    End Function

    Private Function AllItemsAreDelivered() As Boolean
        Dim I As Integer

        'For I = 0 To (UGridIO1.MaxRows - 1)
        '    If IsItem(QueryStyle(I)) And QueryStatus(I) <> "DELTW" Then
        '        AllItemsAreDelivered = False
        '        Exit Function
        '    End If
        'Next
        AllItemsAreDelivered = True
    End Function

    Private Function NoItemsOnSale() As Boolean
        Dim I As Integer

        NoItemsOnSale = False
        'For I = 0 To (UGridIO1.MaxRows - 1)
        '    If IsItem(QueryStyle(I)) Then Exit Function
        'Next
        NoItemsOnSale = True
    End Function

    Private Function HasNonItemsOnSale(Optional ByVal STAIN As Boolean = True, Optional ByVal Delivery As Boolean = True, Optional ByVal Labor As Boolean = True) As Boolean
        Dim I As Integer, T As String

        HasNonItemsOnSale = True
        'For I = 0 To (UGridIO1.MaxRows - 1)
        '    T = Trim(QueryStyle(I))
        '    If STAIN And T = "STAIN" Then Exit Function
        '    If Delivery And T = "DEL" Then Exit Function
        '    If Labor And T = "LAB" Then Exit Function
        'Next

        HasNonItemsOnSale = False
    End Function

    Public Sub HiLiteRow(Optional ByVal N As Integer = -1)
        On Error Resume Next
        'UGridIO1.GetDBGrid.ClearSelCols()

        ' Note: SelBookmarks property is not available for dbgrid control. So commented below lines.
        'Do While UGridIO1.GetDBGrid.SelBookmarks.Count >= 1
        '    UGridIO1.GetDBGrid.SelBookmarks.Remove(0)
        'Loop
        'If N < 0 Then Exit Sub
        'UGridIO1.GetDBGrid.SelBookmarks.Add(" " & N)


    End Sub

    Public Sub AddMarginRow(ByRef Margin As CGrossMargin)
        ' Assumes: X is the last row.  If I remember how to find the last row, this can be made much safer.

        X = X + 1
        SetStyle(X, Margin.Style)
        SetMfg(X, Margin.Vendor)
        SetLoc(X, Margin.Location)
        SetStatus(X, Margin.Status)
        SetQuan(X, Margin.Quantity)
        SetDesc(X, Margin.Desc)
        SetPrice(X, Margin.SellPrice)
    End Sub

    Public Sub SetMfg(ByVal RowNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        SetGridField(RowNum, BillColumns.eManufacturer, CellVal, NoDisplay)
    End Sub

    Public Sub SetLoc(ByVal RowNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        If CellVal = "0" Then CellVal = ""
        SetGridField(RowNum, BillColumns.eLoc, CellVal, NoDisplay)
    End Sub

    Public Sub SetStatus(ByVal RowNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        SetGridField(RowNum, BillColumns.eStatus, CellVal, NoDisplay)
    End Sub

    Private Sub dteSaleDate_Enter(sender As Object, e As EventArgs) Handles dteSaleDate.Enter
        On Error Resume Next

        If Not PollingSaleDate Then
            PollingSaleDate = True
            If Not RequestManagerApproval("Change Sale Date", True) Then
                MsgBox("You do not have access to change the delivery date.", vbExclamation, "Permission Denied")
                CustomerFirst.Select()
                PollingSaleDate = False
                Exit Sub
            End If
        Else
            PollingSaleDate = False
        End If
    End Sub

    Private Sub Sales1_Enter(sender As Object, e As EventArgs) Handles Sales1.Enter, Sales2.Enter, Sales3.Enter
        If cmdApplyBillOSale.Enabled <> False Then
            Sales1.TabStop = False
            frmSalesList.Show()
            Exit Sub
        End If
    End Sub

    Private Sub SalesSplit1_TextChanged(sender As Object, e As EventArgs) Handles SalesSplit1.TextChanged
        vAdjustSalesSplits()
    End Sub

    Private Sub SalesSplit2_TextChanged(sender As Object, e As EventArgs) Handles SalesSplit2.TextChanged
        vAdjustSalesSplits()
    End Sub

    Private Sub SalesSplit3_TextChanged(sender As Object, e As EventArgs) Handles SalesSplit3.TextChanged
        vAdjustSalesSplits()
    End Sub

    'Note: These three events are not required. Above three textchanged events are enough.
    'Private Sub SalesSplit1_Click(sender As Object, e As EventArgs) Handles SalesSplit1.Click
    '    vAdjustSalesSplits()
    'End Sub

    'Private Sub SalesSplit2_Click(sender As Object, e As EventArgs) Handles SalesSplit2.Click
    '    vAdjustSalesSplits()
    'End Sub

    'Private Sub SalesSplit3_Click(sender As Object, e As EventArgs) Handles SalesSplit3.Click
    '    vAdjustSalesSplits()
    'End Sub

    Private Sub StoreCity_Click(sender As Object, e As EventArgs) Handles StoreCity.Click
        BillOSale2_Hide()
    End Sub

    Private Sub StoreName_Click(sender As Object, e As EventArgs) Handles StoreName.Click
        BillOSale2_Hide()
    End Sub

    Private Sub StoreAddress_Click(sender As Object, e As EventArgs) Handles StoreAddress.Click
        BillOSale2_Hide()
    End Sub

    Private Sub StorePhone_Click(sender As Object, e As EventArgs) Handles StorePhone.Click
        BillOSale2_Hide()
    End Sub

    Private Sub txtSpecInst_Enter(sender As Object, e As EventArgs) Handles txtSpecInst.Enter
        SelectContents(txtSpecInst)
    End Sub

    Private Sub txtSaleNo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSaleNo.KeyPress
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'i = Asc(UCase(e.KeyChar))
        e.KeyChar = UCase(e.KeyChar)
        'If Len(txtSaleNo.Text) > 7 Then i = 32
        If Len(txtSaleNo.Text) > 7 Then e.KeyChar = Convert.ToChar(32)
    End Sub

    Private Sub dteDelivery_ValueChanged(sender As Object, e As EventArgs) Handles dteDelivery.ValueChanged
        lblDelDate.Text = DateFormat(dteDelivery.Value)
        DelDate = dteDelivery.Value
        SetDelWeekday(dteDelivery.Value)
    End Sub

    Private Sub dteSaleDate_ValueChanged(sender As Object, e As EventArgs) Handles dteSaleDate.ValueChanged
        'NOTE: THIS EVENT IS REPLACEMENT FOR Private Sub dteSaleDate_Change() EVENT OF VB6.0

        ' "Date" Datepicker.
        TransDate = DateFormat(dteSaleDate.Value)
        If CDate(TransDate) < Today Then
            If MsgBox("Are you sure you want to change the date of the sale?" & vbCrLf &
              "This is not the delivery date.", vbYesNo + vbQuestion) = vbNo Then
                dteSaleDate.Value = Today
                TransDate = Today
            End If
        ElseIf IsWilkenfeld() Then
            If CDate(TransDate) > DateAdd("d", 1, Date.Today) Then
                MsgBox("You can't set the sale date later than tomorrow.", vbCritical)
                TransDate = Today
                dteSaleDate.Value = Today
            End If
        Else
            If CDate(TransDate) > Today Then
                MsgBox("You can't set the sale date later than today.", vbCritical)
                TransDate = Today
                dteSaleDate.Value = Today
            End If
        End If
    End Sub

    Public Sub SetQuan(ByVal RowNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        If CellVal = "0" Then CellVal = ""
        SetGridField(RowNum, BillColumns.eQuant, CellVal, NoDisplay)
    End Sub

    Public Function TotalPaymentsByType(ByVal pType As String, ByRef Count As Integer, Optional ByRef WithTransID As Boolean = True) As Decimal
        Dim I As Integer, T As String

        For I = 0 To LastLineUsed
            T = IfNullThenNilString(QueryQuan(I))

            Count = 0
            If IfNullThenNilString(QueryStyle(I)) = "PAYMENT" And (pType = "" Or pType = T) Then
                Count = Count + 1
                TotalPaymentsByType = TotalPaymentsByType + IfNullThenZeroCurrency(QueryPrice(I))
            End If
        Next

    End Function

    Public ReadOnly Property LastLineUsed() As Integer
        Get
            'For LastLineUsed = 0 To UGridIO1.MaxRows - 1  ' find last item used
            '    If Trim(QueryStyle(LastLineUsed)) = "" Then Exit For
            'Next
            LastLineUsed = LastLineUsed - 1 ' went 1 too far, go back..
        End Get
    End Property

    Public ReadOnly Property LastLineWithItem() As Integer
        Get
            Dim I As Integer
            LastLineWithItem = -1
            'For I = 0 To UGridIO1.MaxRows - 1
            '    If IsItem(QueryStyle(I)) Then LastLineWithItem = I
            'Next
        End Get
    End Property

    Public Sub ReversePayment()
        Exit Sub  ' disabled for now...

        If StoreSettings.CCProcessor <> CCPROC_TC Then Exit Sub
        If Not IsIn(Quan, "3", "4", "5", "6") Then Exit Sub
        If Price = 0 Then MsgBox("This payment has already been reversed.", vbExclamation)
        If QueryTransID(X) = "" Then
            MsgBox("You cannot reverse this payment.  It has no Transaction ID.", vbInformation, ProgramMessageTitle)
            Exit Sub
        End If

        If MsgBox("Do you really want to reverse this payment?", vbQuestion + vbOKCancel, "Cancel Payment") = vbCancel Then
            RefundPayment(X)
        End If
    End Sub

    Private Property Quan() As String
        Get
            Quan = QueryQuan(CurrentLine)
        End Get
        Set(value As String)
            If value = "0" Then value = ""
            SetQuan(CurrentLine, value, False)
        End Set
    End Property

    Private Property QuanSet() As String
        Get
            Return Nothing
        End Get
        Set(value As String)
            If value = "0" Then value = ""
            SetQuan(CurrentLine, value, True)
        End Set
    End Property

    Public Sub QuanFocus(Optional ByVal nRow As Integer = -1)
        GridFocus(BillColumns.eQuant, nRow) ' 4
    End Sub

    Private Property PriceSet() As String
        Get
            Return Nothing
        End Get
        Set(value As String)
            SetPrice(CurrentLine, value, True)
        End Set
    End Property

    Private Property Price() As String
        Get
            Price = QueryPrice(CurrentLine)
        End Get
        Set(value As String)
            SetPrice(CurrentLine, value, False)
        End Set
    End Property

    Public Function QueryTransID(ByVal RowNum As Integer)
        QueryTransID = QueryGridField(RowNum, BillColumns.eTransID)
    End Function

    '  -- Disabled until we decide how we want it to go.
    '  We want to add notes/discounts for one or more items at a time, much like the cash
    '  register module.  The easiest way to do that is to charge full price for the item,
    '  then issue a NOTES line for the discount.  But then the discount isn't shown with
    '  the proper vendor/item in reports.
    '  This could potentially be addressed by adding a hidden SalePrice field to the
    '  GrossMargin table.  That field would allow us to report an item as $100 on the
    '  invoice, but $90 in reports.  Meanwhile, the accompanying NOTES would show -$10
    '  on the invoice and $0 in the reports.
    '  Side effect:  We could also calculate kit sale prices at the time of sale.
    '  What's bad about this method?

    Private Sub cmdChangePrice_Click(sender As Object, e As EventArgs) Handles cmdChangePrice.Click
        '  ChangePriceEnabled False, True  ' Toggles pricechange mode.

        If CanUseDiscountButton() Then
            'frmBOSDiscount.show   1, Me
            frmBOSDiscount.ShowDialog(Me)
            DescFocus()
        Else
            MsgBox("Invalid password or permission.", vbCritical, "Give Discounts")
        End If
    End Sub

    Private Sub cmdEmail_Click(sender As Object, e As EventArgs) Handles cmdEmail.Click
        If IsDate(cmdEmail.Tag) Then
            If Math.Abs(DateDiff("s", cmdEmail.Tag, Now)) < 10 Then
                MsgBox("Please wait a few moments for the email process to finish." & vbCrLf & " You will be notified of any success or failure.", vbExclamation, "Please Wait!")
                Exit Sub
            End If
        End If

        cmdEmail.Tag = Now

        If txtSaleNo.Text <> "" Then
            Dim SSNo As Integer, OldSSNo As Integer
            If cmdMainMenu.Text = "Back" Then SSNo = Val(cmdMainMenu.Tag)
            If SSNo = 0 Then SSNo = StoresSld
            OldSSNo = StoresSld
            If SSNo <> StoresSld Then StoresSld = SSNo

            frmEmail.EmailSale(txtSaleNo.Text)

            If OldSSNo <> 0 And StoresSld <> OldSSNo Then StoresSld = OldSSNo
        Else
            frmEmail.EmailSale(BillOfSale.Text)
        End If
    End Sub

    Private Sub cmdNoChangePrice_Click(sender As Object, e As EventArgs) Handles cmdNoChangePrice.Click
        Exit Sub ' disabled..
        '  ChangePriceEnabled True, True
    End Sub

    Private Sub ChangePriceEnabled(ByVal EditOn As Boolean, ByVal PwPrompt As Boolean)
        Exit Sub ' disabled..
        If Not EditOn Then
            If CheckAccess("Change Item Prices", PwPrompt, True, PwPrompt) Then
                cmdChangePrice.Visible = False
                cmdNoChangePrice.Visible = True
            Else
                cmdChangePrice.Visible = True
                cmdNoChangePrice.Visible = False
            End If
        Else
            cmdChangePrice.Visible = True
            cmdNoChangePrice.Visible = False
        End If
        EnableGridPrice()
    End Sub

    Private Sub EnableGridPrice()
        PriceEnabled = (cmdChangePrice.Visible)
        '  UGridIO1.GetColumn(BillColumns.ePrice).Locked =
    End Sub

    Private Function CanUseDiscountButton() As Boolean
        CanUseDiscountButton = False
        'exit sub  ' leave this line in to disable this feature for everyone

        If OrderMode("A") Then
            CanUseDiscountButton = RequestManagerApproval("Give Discounts")
        Else
            CanUseDiscountButton = False
        End If
    End Function

    Public Sub cmdProcessSale_Click(sender As Object, e As EventArgs)
        'UGridIO1.Refresh()
        cmdProcessSale.Enabled = False
        'DoEvents()
        Application.DoEvents()

        On Error GoTo NewProcessSaleError
        OrdSelect.Hide()

        'If IsDevelopment Then ControlLoading fraBOS2
        If ProcessSale2() Then
            PrintBill = True
        Else
            cmdProcessSale.Enabled = True
        End If
        'If IsDevelopment Then ControlLoadingRemove fraBOS2
        Exit Sub

NewProcessSaleError:
        MsgBox("Error #1 processing sale." & vbCrLf & "Please contact " & AdminContactName & " at " & AdminContactPhone & " with the details of this error immediately.", vbCritical, "Error")
        MsgBox("Error: " & Err.Description, vbInformation, "Error Description")
        Exit Sub
    End Sub

    Public Function ProcessSale2() As Boolean
        Dim S As New sSale

        Disable()
        S.LoadFromBillOSale()
        S.ProcessSale(txtSaleNo.Text) ' print's invoices too
        PrintBill = S.SaleNo <> ""
        DisposeDA(S)
        Disable(True, PrintBill)
        ProcessSale2 = PrintBill
    End Function

    Public Function Disable(Optional ByVal TurnBackOn As Boolean = False, Optional ByVal PostSale As Boolean = False)
        'MousePointer = IIf(TurnBackOn, vbDefault, vbHourglass)
        Me.Cursor = IIf(TurnBackOn, Cursors.Default, Cursors.WaitCursor)
        cmdProcessSale.Enabled = TurnBackOn And Not PostSale
        cmdNextSale.Enabled = TurnBackOn
        cmdMainMenu.Enabled = TurnBackOn
        cmdClear.Enabled = TurnBackOn And Not PostSale
        cmdApplyBillOSale.Enabled = TurnBackOn  ' Disable BillOSale's OK button after saving sale.
        cmdCancel.Enabled = TurnBackOn
        Notes_Open.Enabled = TurnBackOn
        cmdPrint.Enabled = TurnBackOn
        cmdEmail.Enabled = TurnBackOn
        cmdSoldTags.Enabled = TurnBackOn
    End Function

    Private Sub cmdNextSale_Click(sender As Object, e As EventArgs)
        If SaleHasCCTransactions And cmdProcessSale.Enabled = True Then
            If MsgBox("This sale has an already processed Credit Card Transaction." & vbCrLf & "If you leave this sale without processing it, you will have to manually remove the charges." & vbCrLf & "Click Cancel to Process this sale first.", vbExclamation + vbOKCancel + vbDefaultButton2, "Credit Card Transaction Already Processed") = vbCancel Then
                Exit Sub
            End If
        End If

        'Next Sale
        'Unload OrdPay
        OrdPay.Close()

        'Margin.DDelDat = ""
        OrdSelect.ArStatus = ""
        OrdSelect.TaxApplied = ""
        g_Holding.Status = ""
        'PrintBill = False

        If OrderMode("A") Then
            If Not PrintBill Then
                If MsgBox("Bill Of Sale Not Printed Or Posted!", vbExclamation + vbOKOnly, "Sale Not Completed") = vbOK Then
                    Exit Sub
                Else
                End If
            End If

            PrintBill = False
            'Unload MailCheck
            MailCheck.Close()

            'Unload BillOSale
            'BillOSale.Close()
            BalDue.Text = 0
            Sale = 0
            'Unload Me
            Me.Close()
            'Unload OrdSelect
            OrdSelect.Close()

            'Unload OrdStatus
            OrdStatus.Close()


            'Unload ArApp
            ArApp.Close()

            StyleAddEnd()
            NewStyleLine = 0
            SalesTax1 = 0
            SalesTax2 = 0
            OrdSelect.SalesTax1 = 0
            OrdSelect.SalesTax2 = 0
            Written = 0
            Mail.Index = ""
            X = 0
            Show()
            MailCheck.optTelephone.Checked = True
            MailCheck.HidePriorSales = True
        Else
            cmdNextSale.Enabled = True
            'Unload ArApp
            ArApp.Close()

            'Unload BillOSale

            'Unload Me
            Me.Close()
            Show()
            Me.Show()
            MailCheck.optSaleNo.Checked = True
        End If
        frmSalesList.SalesCode = ""
        'MailCheck.Show vbModal
        MailCheck.ShowDialog()
        MailCheck.HidePriorSales = False

        ProcessSalePOs = Nothing
    End Sub

    Private Sub cmdMainMenu_Click(sender As Object, e As EventArgs)
        If SaleHasCCTransactions And cmdProcessSale.Enabled = True Then
            If MsgBox("This sale has an already processed Credit Card Transaction." & vbCrLf & "If you leave this sale without processing it, you will have to manually remove the charges." & vbCrLf & "Click Cancel to Process this sale first.", vbExclamation + vbOKCancel + vbDefaultButton2, "Credit Card Transaction Already Processed") = vbCancel Then
                Exit Sub
            End If
        End If

        If SelectPrinter.SmallTags Then ' small tag was printed
            printer.EndDoc()
            SelectPrinter.SmallTags = False
        End If

        If OrderMode("A") And Not PrintBill Then
            If MsgBox("Bill Of Sale Not Printed Or Posted!  Press Cancel to abort sale.", vbExclamation + vbOKCancel, "Sale Not Completed") = vbOK Then
                cmdProcessSale.Enabled = True
                cmdProcessSale.Select()
                Exit Sub
            Else
            End If
        End If

        ClearBillOfSale()
        'Unload MailCheck
        MailCheck.Close()
        If cmdMainMenu.Text = "Back" Then
            ' Go back to delivery calendar (or whatever else)
        Else
            MainMenu.Show()
        End If
        'Unload Me
        Me.Close()

        If OrderMode("A", "B") Then
            'Unload InvDel
            InvDel.Close()

            'Unload OrdPay
            OrdPay.Close()

            'Unload AddOnAcc
            AddOnAcc.Close()

            'Unload ARPaySetUp
            ARPaySetUp.Close()
        End If
        modProgramState.Order = ""
    End Sub

    Public Sub ClearBillOfSale()
        Dim I As Integer

        If AddingItem Then StyleAddEnd()

        CurrentLine = 0
        NewStyleLine = 0
        DeleteLine = ""
        '  VerifyMailRecUnique MailRec, vbTrue  ' clear it if it was saved..
        MailRec = 0
        Sale = 0
        RN = 0
        Detail = 0
        NsRec1 = 0 : NsRec2 = 0 : NsRec3 = 0 : NsRec4 = 0 : NsRec5 = 0 : NsRec6 = 0

        For I = 1 To Setup_MaxStores_DB
            SetBalance(I, 0)
            SetOnOrder(I, 0)
        Next

        PoSold = 0
        Rb = 0
        NonTaxable = 0
        PrintBill = False
        MarginNo = 0
        Name1 = ""
        LeaseNo = ""
        TransDate = ""
        Written = 0
        TaxCharged1 = 0
        ArCashSls = 0
        Controll = 0
        UndSls = 0
        DelSls = 0
        TaxRec1 = 0
        TaxRec2 = 0
        Deposit = 0
        SalesTax1 = 0
        SalesTax2 = 0

        TaxCode = 0

        TotSale = 0
        Index = 0
        Typpe = ""
        Copies = 0
        Notes = ""

        'cboAdvertisingType.ListIndex = -1
        cboAdvertisingType.SelectedIndex = -1
        'cboCustType.ListIndex = -1
        cboCustType.SelectedIndex = -1
        'cboTaxZone.ListIndex = -1
        cboTaxZone.SelectedIndex = -1

        '  PoNo = 0
        LastSale = ""
        LastMfg = ""
        ProcessSalePOs = Nothing

        NotesInfo = ""
        ' End of declared variables...

        ClearGrid()

        'Unload ArApp
        ArApp.Close()

        IsInternetSale = False

        SaleHasCCTransactions = False
    End Sub

    Public Sub ClearGrid()
        'UGridIO1.Clear()
        'frmSalesList.SafeSalesClear = True
        frmSalesList.SalesCode = ""
        X = 0
        BalDue.Text = 0
        SalesTax2 = 0
        SalesTax1 = 0
        Sale = 0
        Deposit = 0
        Written = 0
        NonTaxable = 0
        SplitKits = vbFalse
    End Sub

    Private Function LinesUsed() As Integer
        With Me.UGridIO1
            For LinesUsed = 0 To .MaxRows - 1
                If .GetValue(LinesUsed, 0) = "" Then Exit Function
            Next
        End With
    End Function

    Private Function LinesFreeRemaining() As Integer
        LinesFreeRemaining = MaxLines - LinesUsed()
    End Function

    Public Function HasTax1() As Boolean
        Dim I As Integer

        HasTax1 = False
        '---> NOTE: COMMENTED THE BELOW LINE. AFTER COMPLETION OF THE CODE, REMOVE THE COMMENT.
        If StoreSettings.SalesTax = 0# Then HasTax1 = True : Exit Function
        For I = 0 To UGridIO1.MaxRows - 1
            If Trim(QueryStyle(I)) = "TAX1" Then HasTax1 = True : Exit Function
        Next
    End Function

    Public Function HasTax2() As Boolean
        Dim I As Integer

        HasTax2 = False
        For I = 0 To UGridIO1.MaxRows - 1
            If Trim(QueryStyle(I)) = "TAX2" Then HasTax2 = True : Exit Function
        Next
    End Function

    Public Function HasTaxableItems() As Boolean
        Dim I As Integer

        HasTaxableItems = False
        For I = 0 To UGridIO1.MaxRows - 1
            If IsItem(QueryStyle(I)) Then HasTaxableItems = True : Exit Function
        Next
    End Function

    Private Function LinePadRequired() As Integer
        If Not HasTax1() Then
            LinePadRequired = 4
        Else
            LinePadRequired = 2
        End If
    End Function

    Public Function IsGridFull(Optional ByVal IgnorePad As Boolean = False) As Boolean
        Dim Re As Integer

        Re = IIf(IgnorePad, 0, LinePadRequired)
        IsGridFull = LinesFreeRemaining() <= Re
    End Function

    Public Function FirstEmptyRow() As Integer
        'FirstEmptyRow = UGridIO1.FirstEmptyRow
    End Function

    Public Function RefundPayment(ByVal I As Integer) As Boolean
        Dim TC As clsTransactionCentral

        If IfNullThenNilString(QueryStyle(I)) <> "PAYMENT" Then
            Err.Raise(-1, , "BilloSale::RefundPayment: Not a Payment...")
        End If
        If IfNullThenNilString(QueryTransID(I)) = "" Then
            Err.Raise(-1, , "BilloSale::RefundPayment: No TransID")
        End If

        TC = New clsTransactionCentral
        TC.TransID = QueryTransID(I)

        If TC.ExecVoid() Then
            DisposeDA(TC)
            SetTransID(I, "")
            ' adjust onscreen
            SetPrice(I, "")
            ' adjust gm
            ExecuteRecordsetBySQL("UPDATE [GrossMargin] SET [TransID]='', [SellPrice]=0 WHERE [TransID]='" & QueryTransID(I) & "'")
            ' adjust holding
            ExecuteRecordsetBySQL("UDPATE [Holding] SET ... WHERE [LeaseNo]='" & "'")
            ' add cash line
            ' add audit line
            RefundPayment = True
        End If
    End Function

    Private Sub cmdClear_Click(sender As Object, e As EventArgs)
        If SaleHasCCTransactions And cmdProcessSale.Enabled = True Then
            If MsgBox("This sale has an already processed Credit Card Transaction." & vbCrLf & "If you leave this sale without processing it, you will have to manually remove the charges." & vbCrLf & "Click Cancel to Process this sale first.", vbExclamation + vbOKCancel + vbDefaultButton2, "Credit Card Transaction Already Processed") = vbCancel Then
                Exit Sub
            End If
        End If

        '  ClearBillOfSale
        ClearGrid()
        CurrentLine = 0
        NewStyleLine = 0

        '  Unload Me
        '  Me.Show
        StyleAddBegin(0)
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Dim X As String, Xs()

        Copies = 1

        ReDim Xs(2)
        Xs(0) = "Current Copy"   ' "File Copy"
        Xs(1) = "Customer Copy"
        Xs(2) = "Delivery Ticket"
        '  Xs(4) = "Salesman Copy"

        X = SelectOptionArray("Choose Copy", frmSelectOption.ESelOpts.SelOpt_List, Xs)
        Select Case Val(X)
            Case 0 : Exit Sub
            Case 2 : Notes = "Customer Copy"
            Case 3
                ''  delivery ticket
                'Load InvPull
                InvPull.Show()

                With InvPull
                    .Pull = 2
                    '.optPrintAll(2) = True
                    .optPrintAll3.Checked = True
                    .txtSaleNo.Text = BillOfSale.Text
                    On Error Resume Next
                    .dteFrom.Value = Today
                    .dteFrom = dteDelivery
                    .dteFrom.Value = DelDate
                    '.cmdPrint(0).Value = True
                    .cmdPrint.PerformClick()

                End With
                'Unload InvPull
                InvPull.Close()
                Exit Sub

            Case Else : Notes = "Current Copy"
        End Select


        Dim C As sSale
        Dim SSNo As Integer
        Dim OldSSNo As Integer

        OldSSNo = 0
        If cmdMainMenu.Text = "Back" Then
            SSNo = Val(cmdMainMenu.Tag)
            If SSNo = 0 Then SSNo = StoresSld
            OldSSNo = StoresSld
            If SSNo <> StoresSld Then StoresSld = SSNo
        End If

        If BillOfSale.Text <> "" Then
            PrintSale(BillOfSale.Text, , Notes)
        Else
            C = New sSale
            C.LoadFromBillOSale()
            C.PrintInvoice(Notes)
            DisposeDA(C)
        End If

        If OldSSNo <> 0 And OldSSNo <> StoresSld Then StoresSld = OldSSNo
    End Sub

    Private Sub ScanUp_Click(sender As Object, e As EventArgs) Handles ScanUp123.Click, ScanUp.Click
        Scan_Common(1)
    End Sub

    Private Sub ScanDn_Click(sender As Object, e As EventArgs) Handles ScanDn.Click
        Scan_Common(-1)
    End Sub

    Private Sub Scan_Common(Optional Offset As Integer = 0)
        Dim SQL As String
        '    Paging = "Y"

        On Error GoTo HandleErr

        SQL = GetSQLByTableLabelIndexNextPreviousCommon(
          Table:=HoldNew_TABLE _
        , Field:="LeaseNo" _
        , Value:=g_Holding.LeaseNo _
        , Direction:=Offset
        )
        g_Holding.DataAccess.Records_OpenSQL(SQL)
        If g_Holding.DataAccess.Records_Available Then
            ClearBillOfSale()  ' This should be all we need, don't have to unload and reload.  Safety first, though.

            ' BFH20111119 - removed this unload/reload... wasn't needed with a good clear function... performance is better
            '      Unload BillOSale
            '      Show
            '      BillOSale2_Show

            MailCheck.optSaleNo.Checked = True
            MailCheck.LookUpCustomer(g_Holding.LeaseNo, False)
        End If
        Exit Sub

HandleErr:
    End Sub

    Public Sub SetTransID(ByVal RowNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        SetGridField(RowNum, BillColumns.eTransID, CellVal, NoDisplay)
    End Sub

    Private Function CheckSplitKits() As Boolean
        Dim I As Integer, IsKit As Boolean, FirstLine As Integer
        Dim J As Integer, TotLanded As Decimal, TotKitPrice As Decimal
        Dim C As CInvRec, X As Decimal, Y As Decimal

        If Not StoreSettings.bShowPackageItemPrices Then Exit Function
        If SplitKits = vbUseDefault Then Exit Function

        For I = 0 To UGridIO1.LastRowUsed
            If IsItem(QueryStyle(I)) And QueryPrice(I) = 0 Then
                If Not IsKit Then FirstLine = I
                IsKit = True
            End If
            If Not IsItem(QueryStyle(I)) Then IsKit = False
            If IsKit And IsItem(QueryStyle(I)) And QueryPrice(I) <> 0 Then
                If SplitKits <> vbTrue Then
                    Dim R As MsgBoxResult
                    R = MsgBox("Set Individual Kit Item Prices?", vbQuestion + vbYesNoCancel, "On-The-Fly Kit")
                    Select Case R
                        Case vbNo : Exit Function
                        Case vbYes : SplitKits = vbTrue
                        Case vbCancel : SplitKits = vbUseDefault : Exit Function
                    End Select
                End If

                TotKitPrice = QueryPrice(I)
                TotLanded = 0

                For J = FirstLine To I
                    C = New CInvRec
                    C.Load(QueryStyle(J), "Style")
                    TotLanded = TotLanded + C.Landed * QueryQuan(J)
                    DisposeDA(C)
                Next

                Y = 0
                For J = FirstLine To I
                    C = New CInvRec
                    C.Load(QueryStyle(J), "Style")
                    X = Math.Round(TotKitPrice * (C.Landed * QueryQuan(J) / TotLanded), 2)
                    If J <> I Then
                        Y = Y + X
                        SetPrice(J, X)
                    Else
                        SetPrice(J, TotKitPrice - Y)   '' penny watch
                    End If
                    DisposeDA(C)
                Next



            End If
        Next
        CheckSplitKits = True
    End Function

    Public Property LastRecord() As Integer
        Get
            LastRecord = mLastRecord
        End Get
        Set(value As Integer)
            mLastRecord = value
        End Set
    End Property

    Private ReadOnly Property LastGridText(Optional ByVal Row As Integer = -1, Optional ByVal Col As Integer = -1) As String
        Get
            'If Row = -1 Then Row = UGridIO1.Row
            'If Col = -1 Then Col = UGridIO1.Col
            If Row <> LGTRow Or Col <> LGTCol Then
                Exit Property
            End If
            LastGridText = mLastGridText
        End Get
        'Set(value As String)
        'LGTRow = UGridIO1.Row
        'LGTCol = UGridIO1.Col
        'mLastGridText = value
        'End Set
    End Property

    Public Function GetBalance(ByVal StoreNum As Integer) As Double
        If StoreNum <= 0 Or StoreNum > Setup_MaxStores_DB Then Exit Function
        GetBalance = Ba(StoreNum - 1)
    End Function

    Public Function SetBalance(ByVal StoreNum As Integer, ByVal nBalance As Double)
        If StoreNum <= 0 Or StoreNum > Setup_MaxStores_DB Then Exit Function
        Ba(StoreNum - 1) = nBalance
    End Function

    Public Function GetTotalBalance() As Double
        Dim I As Integer

        For I = 1 To Setup_MaxStores
            GetTotalBalance = GetTotalBalance + GetBalance(I)
        Next
    End Function

    Public Function GetOnOrder(ByVal StoreNum As Integer) As Double
        If StoreNum <= 0 Or StoreNum > Setup_MaxStores_DB Then Exit Function
        GetOnOrder = OO(StoreNum - 1)
    End Function

    Public Function SetOnOrder(ByVal StoreNum As Integer, ByVal nOnOrd As Double)
        If StoreNum <= 0 Or StoreNum > Setup_MaxStores_DB Then Exit Function
        OO(StoreNum - 1) = nOnOrd
    End Function

    Public Function GetTotalOnOrder() As Double
        Dim I As Integer

        For I = 1 To Setup_MaxStores
            GetTotalOnOrder = GetTotalOnOrder + GetOnOrder(I)
        Next
    End Function

    Public Sub StyleAddEnd(Optional ByVal Restart As Boolean = False, Optional ByVal NumLines As Integer = 1)
        'Unload OrdSelect
        OrdSelect.Close()
        If Restart Then
            StyleAddBegin(0)
        Else
            NewStyleLine = NewStyleLine + NumLines
            AddingItem = False
        End If
        Recalculate()
        '  Debug.Print "Ending style addition on " & NewStyleLine
    End Sub

    Private Property Loc() As String
        Get
            Loc = QueryLoc(CurrentLine)
            If Loc = "" Then Loc = "0"
        End Get
        Set(value As String)
            If value = "0" Then value = ""
            SetLoc(CurrentLine, value, False)
        End Set
    End Property

    Private Property LocSet() As String
        Get
            Return Nothing
        End Get
        Set(value As String)
            If value = "0" Then value = ""
            SetLoc(CurrentLine, value, True)
        End Set
    End Property

    Private Property Status() As String
        Get
            Status = QueryStatus(CurrentLine)
        End Get
        Set(value As String)
            SetStatus(CurrentLine, value, False)
        End Set
    End Property

    Private Sub UGridIO1_AfterDelete()
        Recalculate()
        'If UGridIO1.GetDBGrid.FirstRow = 1 Then UGridIO1.GetDBGrid.FirstRow = 0
        'UGridIO1.Refresh(True)
    End Sub

    Private Sub UGridIO1_BeforeColUpdate(ColIndex As Integer, OldValue As Object, Cancel As Integer)
        Dim NewVal As String

        'With UGridIO1
        '    NewVal = .Text
        '    Select Case ColIndex
        '        Case BillColumns.eStyle  ' Style
        '            If Len(NewVal) > Setup_2Data_StyleMaxLen Then .Text = Microsoft.VisualBasic.Left(NewVal, Setup_2Data_StyleMaxLen)
        '        Case BillColumns.eManufacturer  ' Mfg
        '            If Len(NewVal) > Setup_2Data_ManufMaxLen Then .Text = Microsoft.VisualBasic.Left(NewVal, Setup_2Data_ManufMaxLen)
        '        Case BillColumns.eLoc  ' Loc
        '            If Not IsNumeric(NewVal) Then
        '                NewVal = ""
        '                .Text = NewVal
        '            ElseIf Not InRange(1, Val(NewVal), Setup_MaxStores) Then
        '                NewVal = FitRange(1, Val(NewVal), Setup_MaxStores)
        '                .Text = NewVal
        '            End If
        '        Case BillColumns.eStatus  ' Status
        '            If Len(NewVal) > 5 Then .Text = Microsoft.VisualBasic.Left(NewVal, 5)
        '        Case BillColumns.eQuant  ' Quantity
        '            If Not IsNumeric(NewVal) Then NewVal = "" : .Text = NewVal
        '            If QueryPrice(X) <> 0 Then
        '                .SetValueDisplay(X, BillColumns.ePrice, CurrencyFormat(Val(.Text) * QueryPrice(X) / Val(OldValue)))
        '                Recalculate()
        '                .Text = NewVal
        '            End If
        '        Case BillColumns.eDescription  ' Description
        '            If Len(NewVal) > Setup_2Data_DescMaxLen Then .Text = Microsoft.VisualBasic.Left(NewVal, Setup_2Data_DescMaxLen)
        '        Case BillColumns.ePrice  ' Price
        '            '.Text = Format(GetPrice(.Text), "###,###.00")
        '    End Select
        'End With
    End Sub

    Private WriteOnly Property StatusSet() As String
        'Get
        '    Return Nothing
        'End Get
        Set(value As String)
            SetStatus(CurrentLine, value, True)
        End Set
    End Property

    Private Property DescSet() As String
        Get
            Return Nothing
        End Get
        Set(value As String)
            SetDesc(CurrentLine, value, True)
        End Set
    End Property

    Public Sub DescFocus(Optional ByVal nRow As Integer = -1)
        GridFocus(BillColumns.eDescription, nRow) ' 5
    End Sub

    Private Property MfgNo() As String
        Get
            MfgNo = QueryMfgNo(CurrentLine)
        End Get
        Set(value As String)
            SetMfgNo(CurrentLine, value, False)
        End Set
    End Property

    Private Property TransID() As String
        Get
            TransID = QueryTransID(CurrentLine)
        End Get
        Set(value As String)
            SetTransID(CurrentLine, value, False)
        End Set
    End Property

    Private Property TransIDSet() As String
        Get
            Return Nothing
        End Get
        Set(value As String)
            SetTransID(CurrentLine, value, True)
        End Set
    End Property

    Public Sub SetStyle(ByVal RowNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        SetGridField(RowNum, BillColumns.eStyle, Microsoft.VisualBasic.Left(CellVal, Setup_2Data_StyleMaxLen), NoDisplay)
    End Sub

    Public Function QueryMfg(ByVal RowNum As Integer) As String
        QueryMfg = QueryGridField(RowNum, BillColumns.eManufacturer)
    End Function

    Public Function QueryPrice(ByVal RowNum As Integer) As Decimal
        QueryPrice = GetPrice(QueryGridField(RowNum, BillColumns.ePrice))
    End Function

    Private Sub Email_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Email.Validating
        Email.Text = Trim(Email.Text)
        If Email.Text = "" Then Exit Sub
        If InStr(Email.Text, "@") = 0 Or InStr(Email.Text, ".") = 0 Or Len(Email.Text) < 5 Then
            MsgBox("Invalid email address.  Email address must be in 'user@company.com' format.")
            'Cancel = True
            e.Cancel = True
            Exit Sub
        End If
        If InStr(Email.Text, " ") <> 0 Then
            MsgBox("Invalid email address.  Email address must not contain spaces.")
            'Cancel = True
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub chkDelivery_Click(sender As Object, e As EventArgs) Handles chkDelivery.Click
        If chkDelivery.CheckState = 1 Then
            chkPickup.CheckState = 0
            If OrderMode("A") Then
                dteDelivery.Value = DateFormat(Now)
                dteDelivery.Visible = True
                dteDelivery.Enabled = True
                lblDelDate.Visible = False
            End If

            ShowTimeWindowBox(True)

            lblDelDate.Text = DateFormat(dteDelivery.Value)
            DelDate = DateFormat(dteDelivery.Value)
        End If

        If OrderMode("A") Then SetDelWeekday(Today)

        If chkDelivery.CheckState = 0 Then
            lblDelWeekday.Text = "None"
            dteDelivery.Visible = False
            lblDelDate.Text = ""
            lblDelDate.Visible = True
            ShowTimeWindowBox(False)
        End If
    End Sub

    Private Sub chkPickup_Click(sender As Object, e As EventArgs) Handles chkPickup.Click
        If chkPickup.CheckState = 1 Then
            chkDelivery.CheckState = 0
            If OrderMode("A") Then
                dteDelivery.Visible = True
                dteDelivery.Enabled = True
                dteDelivery.Value = DateFormat(Now)
                lblDelDate.Visible = False
                ShowTimeWindowBox(True)
            End If
            lblDelDate.Text = DateFormat(dteDelivery.Value)
            DelDate = DateFormat(dteDelivery.Value)
        End If

        If OrderMode("A") Then SetDelWeekday(Today)

        If chkPickup.CheckState = 0 Then
            lblDelWeekday.Text = "None"
            dteDelivery.Visible = False
            lblDelDate.Text = ""
            lblDelDate.Visible = True
            ShowTimeWindowBox(False)
        End If
    End Sub

    Private Sub cmdApplyBillOSale_Click(sender As Object, e As EventArgs) Handles cmdApplyBillOSale.Click
        If CustomerLast.Text <> "" And CustomerLast.Text <> "CASH & CARRY" And OrderMode("A") And (IsUFO() Or StoreSettings.bRequireAdvertising) And Not IsInternetSale Then
            If cboAdvertisingType.SelectedIndex < 2 Then
                MsgBox("You must select advertising type!", vbCritical)
                Exit Sub
            End If
        End If

        If Not MailMode("ADD/Edit", "Book") And Not OrderMode("A") And Not ArMode("S", "A") Then
            If StoreSettings.bManualBillofSaleNo And Not IsInternetSale Then
                If Trim(txtSaleNo.Text) = "" Then  'manual bill of sale
                    MsgBox("Please enter a Bill of Sale Number!", vbExclamation, "No BoS Number")
                    txtSaleNo.Select()
                    Exit Sub
                End If
                'check for duplicate sale no
                'CheckSaleNo
                'If Trim(Holding.LeaseNo) = Trim(txtSaleNo) = True Then
                If Not Visible Or cmdProcessSale.Enabled Then
                    If g_Holding.Load(Trim(txtSaleNo.Text)) Then
                        MsgBox("This Sale Number has been used!", vbExclamation)
                        txtSaleNo.Text = ""
                        txtSaleNo.Select()
                        Exit Sub
                    End If
                End If
            End If
        End If

        If Trim(CustomerLast.Text) = "" And ArMode("S", "A") Then
            MsgBox("Last Name field must be filled in to set up an Installment Contract!", vbExclamation)
            Exit Sub
        End If

        'NOTE: REMOVE THIS COMMENT LATER.
        'If Trim(Email.Text) = "" And OrderMode("A") And StoreSettings.bRequestEmail And Email.Tag = "" Then
        '    If CustomerLast.Text <> "" Then
        '        If MsgBox("Add Customer Email?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
        '            On Error Resume Next
        '            Email.Select()
        '            Exit Sub
        '        Else
        '            Email.Tag = "x"
        '        End If
        '    Else
        '        Email.Tag = "x"
        '    End If
        'End If

        If Trim(CustomerLast.Text) = "" Or Trim(CustomerLast.Text) = "CASH & CARRY" Then
            If Trim(CustomerFirst.Text) <> "" Or Trim(CustomerAddress.Text) <> "" Or Trim(AddAddress.Text) <> "" Or Trim(CustomerCity.Text) <> "" Or Trim(CustomerZip.Text) <> "" Or Trim(CustomerPhone1.Text) <> "" Then
                If MsgBox("Software requires a LAST NAME to save the customer." & vbCrLf & "If you leave the last name blank, sale will be treated as CASH & CARRY." & vbCrLf2 & "Would you like to add a last name?", vbYesNo + vbQuestion, "Last Name") = vbYes Then
                    On Error Resume Next
                    CustomerLast.Select()
                    Exit Sub
                End If
            End If
        End If

        If Sales1.Text = "" Then
            If CustomerLast.Text = "" Then CustomerLast.Text = "CASH & CARRY"
            'Default
            Sales1.Text = "HOUSE" : SalesSplit1.Text = "100%"
            Sales2.Text = ""   ' Added MJK20031230
            Sales3.Text = ""
            frmSalesList.SalesCode = "99"
        End If
        ugrFake.Visible = False

        'Update mailing list
        'Updated 01/03/2002 to prevent BillOSale from loading when using Add/Edit or Old Account Setup
        If CustomerLast.Text <> "CASH & CARRY" Then 'Taken out temporarily to fix And not mailmode("ADD/Edit") And not armode("S") Then
            CorrectMailDatabase()
        End If

        If MailMode("ADD/Edit", "Book") Or ArMode("S") Then
            'if telephone number is changed
            'mailing list & a/r old customers
            If Typpe = "0" And Index <> 0 Then
                If CleanAni(MailCheck.CustomerTele) <> CleanAni(CustomerPhone1.Text) Then
                    ExecuteRecordsetBySQL("UPDATE GrossMargin SET Tele='" & CleanAni(CustomerPhone1.Text) & "' WHERE MailIndex=" & Index)
                    ExecuteRecordsetBySQL("UPDATE InstallmentInfo SET Telephone='" & CleanAni(CustomerPhone1.Text) & "' WHERE MailIndex=" & Index)
                    ExecuteRecordsetBySQL("UPDATE Service SET Telephone='" & CleanAni(CustomerPhone1.Text) & "' WHERE MailIndex=" & Index)
                    'FixPhone
                End If
            End If

            If MailMode("ADD/Edit") Then
                cmdApplyBillOSale.Enabled = True
                'Unload BillOSale
                Me.Close()
                Show()
                MailCheck.optTelephone.Checked = True
                'MailCheck.Show(vbModal, BillOSale)
                MailCheck.ShowDialog(Me)
                Exit Sub
            ElseIf MailMode("Book") Then
                cmdApplyBillOSale.Enabled = True
                'Unload BillOSale
                Me.Close()
                MailBook.Show()
                Exit Sub
            End If
        End If

        If OrderMode("A") Then '<> "ADD/Edit" And not armode("S") Then
            'this is where BOS2 gets opened
            On Error Resume Next
            Arrange(True, False)

            If Not BOS2IsHidden Then
                Dim PrePrintBill As Boolean
                PrePrintBill = PrintBill      '' dont clear it here
                BillOSale2_Show()
                PrintBill = PrePrintBill
                '        OrdSelect.Show ' bfh20060113 - when the grid selects its row (first col), it will open this automatically for us!!
            End If
            If Not OrderMode("A") Then cmdApplyBillOSale.Enabled = False
        End If

        If ArMode("S") Then
            ARPaySetUp.Show()
        End If

        If ArMode("A") Then
            ArApp.Show()
            ArApp.txtFirstName = CustomerFirst
            ArApp.txtLastName = CustomerLast
            ArApp.txtAddress = CustomerAddress
            ArApp.txtCity = CustomerCity
            ArApp.txtZip = CustomerZip
            ArApp.txtTele1 = DressAni(CleanAni(CustomerPhone1.Text))
            ArApp.txtTele2 = DressAni(CleanAni(CustomerPhone2.Text))
            ArApp.lblTelephone.Text = cboPhone1.Text
            'Unload Me
            Me.Close()
        End If

        'Added this line to connect and execute Axdatagrid1RowColumnChange sub. Because AxDataGrid1_RowColChange event in vb 6 is auto executing
        'but Is not auto executing in vb.net.
        UGridIO1.Axdatagrid1RowColumnChange(Nothing, -1)

    End Sub

    Private Sub UGridIO1_KeyPress(sender As Object, e As KeyPressEventArgs)
        ResetLastLoginExpiry()

        'Select Case UGridIO1.Col
        '    Case BillColumns.eStyle
        '    Case BillColumns.eManufacturer
        '        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        '        e.KeyChar = UCase(e.KeyChar)
        '        'If KeyAscii = Asc(",") Then KeyAscii = Asc(";") :                 '   change , to ;
        '        If e.KeyChar = "," Then e.KeyChar = ";"

        '    Case BillColumns.eStatus
        '        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        '        e.KeyChar = UCase(e.KeyChar)
        '    Case BillColumns.eDescription
        '        If Len(UGridIO1.GetDBGrid.Text) >= Setup_2Data_DescMaxLen Then
        '            'If KeyAscii > 26 Then ' ignore control codes
        '            If e.KeyChar > Convert.ToChar(26) Then
        '                UGridIO1.GetDBGrid.Text = Microsoft.VisualBasic.Left(UGridIO1.GetDBGrid.Text, Setup_2Data_DescMaxLen - 1) ' + the new character will be MAX
        '            End If
        '        End If

        '        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        '        e.KeyChar = UCase(e.KeyChar)
        '        'If KeyAscii = Asc(",") Then KeyAscii = Asc(";") :                 '   change , to ;
        '        If e.KeyChar = "," Then e.KeyChar = ";"

        '    Case BillColumns.eLoc
        '    Case BillColumns.ePrice
        'End Select
    End Sub

    Private Sub UGridIO1_Leave(sender As Object, e As EventArgs)
        FormatHelper("")
    End Sub

    Public Function QueryMfgNo(ByVal RowNum As Integer) As String
        QueryMfgNo = QueryGridField(RowNum, BillColumns.eManufacturerNo)
    End Function

    Private Sub cmdPrintLabel_Click(sender As Object, e As EventArgs) Handles cmdPrintLabel.Click
        PrintDYMOMailingLabel(CustomerLast.Text, CustomerFirst.Text, CustomerAddress.Text, AddAddress.Text, CustomerCity.Text, CustomerZip.Text, IIf(opt30252.Checked, 30252, 30323))
    End Sub

    Private Sub tmrFormat_Tick(sender As Object, e As EventArgs) Handles tmrFormat.Tick
        'Debug.Print "tmrHover_Timer()  "
        tmrFormat.Enabled = False
        'FormatHelper(UGridIO1.GetValue(tmrFormat.Tag, 5), tmrFormat.Tag)
        '  Debug.Print "Display=" & tmrFormat.Tag
    End Sub

    Private Sub Notes_Open_Click(sender As Object, e As EventArgs)
        frmNotes.DoNotes(0, BillOfSale.Text)
        Exit Sub
    End Sub

    Private Sub tmrHover_Tick(sender As Object, e As EventArgs) Handles tmrHover.Tick
        'Debug.Print "tmrHover_Timer()  "
        tmrHover.Enabled = False
        'HoverPic(True, UGridIO1.GetValue(Val(tmrHover.Tag), 0), 0, Val(tmrHover.Tag))
    End Sub

    Private Sub UGridIO1_DoubleClick(sender As Object, e As EventArgs)
        'X = UGridIO1.Row
        'Select Case UGridIO1.Col
        '    Case BillColumns.eStyle
        '        Style_DblClick()
        '    Case BillColumns.eManufacturer
        '    Case BillColumns.eDescription
        '    Case BillColumns.eStatus
        '        Status_DblClick()
        '    Case BillColumns.ePrice
        '        'Pop up a discount window?  -- Disabled until we decide how we want it to go.
        '        'Dim Discount As String
        '        'If Not CheckAccess("Give Discounts", True, False, True) Then Exit Sub
        '        'Discount = InputBox("Enter discount percentage:", , "0")
        'End Select
    End Sub

    Private Sub UGridIO1_MouseMoveOverCell(ByVal Col As Integer, ByVal Row As Integer)
        On Error Resume Next

        ResetLastLoginExpiry()
        '  FormatHelper ""

        If OrderMode("A", "E") Then    ' show hover pic in sale creation and viewing
            'Debug.Print "Col=" & Col, "row=" & Row & ", x=[" & tmrHover.Tag & "]"
            If Col = 0 Then
                If tmrHover.Tag <> "" And CInt(tmrHover.Tag) = Row Then Exit Sub
                tmrHover.Tag = Row
                HoverTimer(True)
            Else
                HoverTimer()
            End If
            If OrderMode("E") Then
                If Col = 5 Then
                    FormatTimer(True, Row)
                Else
                    FormatTimer()
                End If
            End If
        End If
    End Sub

    Public Sub SetMfgNo(ByVal RowNum As Integer, ByVal CellVal As String, Optional ByVal NoDisplay As Boolean = False)
        SetGridField(RowNum, BillColumns.eManufacturerNo, CellVal, NoDisplay)
    End Sub

    Public Property StyleEnabled() As Boolean
        Get
            StyleEnabled = UGridIO1.GetColumn(BillColumns.eStyle).Locked
        End Get
        Set(value As Boolean)
            UGridIO1.GetColumn(BillColumns.eStyle).Locked = Not value
        End Set
    End Property

    Public Property MfgEnabled() As Boolean
        Get
            MfgEnabled = UGridIO1.GetColumn(BillColumns.eManufacturer).Locked
        End Get
        Set(value As Boolean)
            UGridIO1.GetColumn(BillColumns.eManufacturer).Locked = Not value
        End Set
    End Property

    Public Property LocEnabled() As Boolean
        Get
            LocEnabled = UGridIO1.GetColumn(BillColumns.eLoc).Locked
        End Get
        Set(value As Boolean)
            UGridIO1.GetColumn(BillColumns.eLoc).Locked = Not value
        End Set
    End Property

    Public Property StatusEnabled() As Boolean
        Get
            StatusEnabled = UGridIO1.GetColumn(BillColumns.eStatus).Locked
        End Get
        Set(value As Boolean)
            UGridIO1.GetColumn(BillColumns.eStatus).Locked = Not value
        End Set
    End Property

    Public Property QuanEnabled() As Boolean
        Get
            QuanEnabled = UGridIO1.GetColumn(BillColumns.eQuant).Locked
        End Get
        Set(value As Boolean)
            UGridIO1.GetColumn(BillColumns.eQuant).Locked = Not value
        End Set
    End Property

    Public Property DescEnabled() As Boolean
        Get
            DescEnabled = UGridIO1.GetColumn(BillColumns.eDescription).Locked
        End Get
        Set(value As Boolean)
            On Error Resume Next
            UGridIO1.GetColumn(BillColumns.eDescription).Locked = Not value
        End Set
    End Property

    Private Function IsRowEmpty(ByVal RowNum As Integer) As Boolean
        IsRowEmpty = False
        If Trim(QueryStyle(RowNum)) = "" Then IsRowEmpty = True
        If Trim(QueryPrice(RowNum)) = "" Then IsRowEmpty = True
    End Function

    'Public Function MaxRow() As Integer
    '    For MaxRow = 0 To UGridIO1.MaxRows
    '        If IsRowEmpty(MaxRow) Then Exit Function
    '    Next
    'End Function

    'Private Function IsRowComplete() As Boolean
    '    IsRowComplete = False
    '    If (UGridIO1.MaxRows = 1) Then IsRowComplete = True : Exit Function
    '    If Style = "" Then Exit Function
    '    If Price = "" Then Exit Function
    '    IsRowComplete = True
    'End Function

    Private WriteOnly Property MfgNoSet As String
        'Get

        'End Get
        Set(value As String)
            SetMfgNo(CurrentLine, value, True)
        End Set
    End Property


End Class