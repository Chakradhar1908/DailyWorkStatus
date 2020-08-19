Imports WinCDS
Public Class OnScreenReport
    Dim PoNo As Integer  ' Saved between calls to MakePO.
    Dim Margin As New CGrossMargin
    Dim MarginNo As Integer
    Dim Row As Integer
    Private balRow As Integer
    Dim Mail As MailNew
    Dim LastName As String
    Dim Tele As String
    Public Index As String
    Dim OrdTotal As Decimal
    Dim TotDue As Decimal
    Dim TaxBackedOut As Boolean
    Dim KitStart As Integer, IsKit As Boolean, KitTotalCost As Decimal

    Private mCurrentLine As Integer 'Current line selected
    Dim Counter As Integer
    Dim mLoading As Boolean
    Dim Lines As Integer

    ' These need to be replaced!  We can do the same thing better with hidden grid columns.
    Dim Quantity(500) As Object
    Dim InvRn(500) As Object
    Dim Cost(500) As Object
    Dim Freight(500) As Object
    Dim Depts(500) As Object
    Dim Vends(500) As Object
    Dim DetailRec(500) As Object

    Public Balance As Decimal, TotTax As Decimal
    Dim SaleNo As String
    Dim Detail As Integer

    'Dim NoOnHand As String             ' Was never used..
    Dim FirstTime As Boolean
    'Private AddedInventory As Boolean  ' Was set but never used..
    Dim LastMfg As String

    Dim LastSale As String                         ' For determining which PO items go on.
    Dim Sales As String
    Dim TaxLoc As Integer
    Dim TaxRate As Integer
    Dim Rate As Object
    Dim SalesTax As Boolean
    Dim PriceChg As String
    Dim SubBalance As Decimal
    Dim PriorBal As Decimal
    Dim NonTaxable As Decimal
    Public LeaveCreditBalance As Boolean

    Private WithEvents MailCheckRef As MailCheck
    Private SaleFound As Boolean

    Dim WasDelSale As Boolean, WriteOutAddedItems As Boolean, WriteOutRemovedAllUndelivered As Boolean
    Dim AskedForTaxRate As Boolean

    Const AllowAdjustDel As Boolean = True
    Const MaxAdjustments As Integer = 30
    Public MailCheckSaleNoChecked As Boolean
    Dim FromCmdMenu As Boolean

    Private Sub OnScreenReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim P As Object
        SetButtonImage(cmdReturn, 11)

        'SetButtonImage(cmdAdd, 0). Below line is replacement for it. Because, it is checkbox not button, so loaded image directly here instead caling setbuttonimage funtion.
        cmdAdd.Image = MainMenu.imlStandardButtons.Images(0)

        SetButtonImage(cmdApply, 2)
        SetButtonImage(cmdNext2, 6)
        SetButtonImage(cmdMenu2, 9)
        SetButtonImage(cmdPrint2, 19)

        SetButtonImage(cmdNext, 6)
        SetButtonImage(cmdPrint, 19)
        SetButtonImage(cmdMenu, 9)
        SetButtonImage(cmdAllStores, 24)

        UGridIO2.Visible = False
        Text = StoreSettings.Name
        Margin.Name = ""
        Margin.Phone = ""
        Margin.Salesman = ""
        Counter = 0

        Lines = 0
        Row = 0
        mCurrentLine = 0
        balRow = 0

        '    Text3.Visible = False
        TotDue = 0
        FirstTime = True
        LeaveCreditBalance = False

        If ReportsMode("I") Then
            fraControls2.Visible = False
            fraControls1.Visible = True
            fraControls1.Top = UGridIO2.Top + UGridIO2.Height + 360
            Height = fraControls1.Top + fraControls1.Height + 120 + (Height - Me.ClientSize.Height)

            txtDiffTax0.Visible = False
            txtBalDue.Visible = False
            lblDiffTax.Visible = False
            lblBalDue.Visible = False
            cmdNext.Visible = False
            lblPrevBal2.Visible = False
            lblPrevBal.Visible = False

            UGridIO1.Height = 5415

            Text = "Items Sold Since:  " & InvReports.dteStartDate.Value
            'Left = (Screen.Width - Width) / 2
            Left = (Screen.PrimaryScreen.Bounds.Width - Width) / 2

            UGridIO1.AddColumn(0, "Style", 2000, True, False)
            UGridIO1.AddColumn(1, "Manufacturer", 2000, True, False)
            UGridIO1.AddColumn(2, "Status", 700, True, False)
            UGridIO1.AddColumn(3, "Quan", 550, True, False)
            UGridIO1.AddColumn(4, "Description", 4450, True, False)
            'UGridIO1.AddColumn(5, "Price", 1050, True, False, MSDBGrid.AlignmentConstants.dbgRight)
            UGridIO1.AddColumn(5, "Price", 1050, True, False, MSDataGridLib.AlignmentConstants.dbgRight)

            UGridIO1.MaxCols = 6
            UGridIO1.MaxRows = 1000
            UGridIO1.Initialize()
            '.RowHeight = UGridIO1.Height / (items_per_page + 2) ' This handles the height of the individual rows. Which will indirectly effect the number of rows displayed.
            UGridIO1.Activated = True
            UGridIO1.Refresh()
            UGridIO1.Col = 0
            UGridIO1.Row = 0

            UGridIO2.MaxRows = 10
            UGridIO2.Refresh()
            OnScreenReport
            Exit Sub
        ElseIf ReportsMode("H") Or OrderMode("Credit") Then
            fraControls1.Visible = True
            fraControls2.Visible = False
            'fraControls1.Top = UGridIO2.Top + UGridIO2.Height + 120
            fraControls1.Top = UGridIO2.Top + UGridIO2.Height + 12
            'Height = fraControls2.Top + fraControls1.Height + 120
            Height = fraControls2.Top + fraControls1.Height + 120

            If ReportsMode("H") Then  'customer history
                fraControls1.Visible = True
                fraControls2.Visible = False
                fraControls1.Top = UGridIO2.Top + UGridIO2.Height + 360
                Height = fraControls2.Top + fraControls1.Height + 120 + (Height - Me.ClientSize.Height)

                txtDiffTax0.Visible = False
                txtBalDue.Visible = False
                lblDiffTax.Visible = False
                lblBalDue.Visible = False
                lblPrevBal2.Visible = False
                lblPrevBal.Visible = False

                UGridIO1.Height = 5415
            End If

            If OrderMode("Credit") Then
                fraControls1.Visible = False
                fraControls2.Visible = True
                'fraControls2.Top = UGridIO2.Top + UGridIO2.Height + 360
                fraControls2.Top = UGridIO2.Top + UGridIO2.Height + 36
                'Height = fraControls2.Top + fraControls2.Height + 120 + (Height - Me.ClientSize.Height)
                Height = fraControls2.Top + fraControls2.Height + 40 + (Height - Me.ClientSize.Height)

                lblCaption.Text = "Customer Adjustment"
                lblCaption.Left = Width / 3
                'cmdMenu2.Cancel = True
                Me.CancelButton = cmdMenu2

                txtDiffTax0.Visible = True
                txtBalDue.Visible = True
                lblDiffTax.Visible = True
                lblBalDue.Visible = True

                'use for customer returns
                'UGridIO1.Height = 3000
                UGridIO1.Height = 200
                UGridIO2.Visible = True
                'UGridIO2.Top = 4000
                UGridIO2.Top = 265
                fraControls2.Top = UGridIO2.Top + UGridIO2.Height + 46
                fraControls2.Left = Width / 3

                '    With UGridIO2
                UGridIO2.AddColumn(0, "Sale No", 80, True, False)
                UGridIO2.AddColumn(1, "Style", 90, False, False)
                UGridIO2.AddColumn(2, "Manufacturer", 150, False, False)
                UGridIO2.AddColumn(3, "Loc", 30, False, False)
                UGridIO2.AddColumn(4, "Status", 40, False, False)
                UGridIO2.AddColumn(5, "Quan", 30, False, False)
                UGridIO2.AddColumn(6, "Description", 180, False, False)
                'UGridIO2.AddColumn(7, "Price", 70, False, False, MSDBGrid.AlignmentConstants.dbgRight)
                UGridIO2.AddColumn(7, "Price", 70, False, False, MSDataGridLib.AlignmentConstants.dbgRight)
                'UGridIO2.AddColumn(8, "Difference", 80, True, False, MSDBGrid.AlignmentConstants.dbgRight)
                UGridIO2.AddColumn(8, "Difference", 80, True, False, MSDataGridLib.AlignmentConstants.dbgRight)
                UGridIO2.AddColumn(9, "Unit Price", 80, True, False, , False)
                UGridIO2.AddColumn(10, "MarginLine", 50, True, False, , False)
                UGridIO2.AddColumn(11, "Commission", 50, True, False, , False) ' MJK 20131026

                UGridIO2.GetColumn(7).NumberFormat = "###,##0.00"
                UGridIO2.GetColumn(8).NumberFormat = "###,##0.00"

                UGridIO2.MaxCols = 12
                UGridIO2.MaxRows = MaxAdjustments
                UGridIO2.Initialize()
                '.RowHeight = UGridIO2.Height / (items_per_page + 2) ' This handles the height of the individual rows. Which will indirectly effect the number of rows displayed.
                UGridIO2.Refresh()
                UGridIO2.Col = 0
                UGridIO2.Row = 0
            End If
            OrdTotal = 0
            TotDue = 0
            SubBalance = 0
            '  With UGridIO1
            UGridIO1.AddColumn(0, "Sale No", 80, True, False)
            UGridIO1.AddColumn(1, "Style", 90, True, False)
            UGridIO1.AddColumn(2, "Manufacturer", 150, True, False)
            UGridIO1.AddColumn(3, "Loc", 30, True, False)
            UGridIO1.AddColumn(4, "Status", 40, True, False)
            'UGridIO1.AddColumn(5, "Quan", 30, True, False, MSDBGrid.AlignmentConstants.dbgRight)
            UGridIO1.AddColumn(5, "Quan", 30, True, False, MSDataGridLib.AlignmentConstants.dbgRight)
            UGridIO1.AddColumn(6, "Description", 180, True, False)
            UGridIO1.AddColumn(7, "Price", 70, True, False, MSDataGridLib.AlignmentConstants.dbgRight)
            'UGridIO1.AddColumn(8, "Total Due", 70, True, False, MSDBGrid.AlignmentConstants.dbgRight)
            UGridIO1.AddColumn(8, "Total Due", 70, True, False, MSDataGridLib.AlignmentConstants.dbgRight)
            UGridIO1.AddColumn(9, "MarginLine", 70, True, False, , False)
            UGridIO1.AddColumn(10, "KitPrice", 70, True, False, , False)
            UGridIO1.AddColumn(11, "Commission", 50, True, False, , False) ' MJK 20131026

            UGridIO1.MaxCols = 12
            UGridIO1.MaxRows = 800
            UGridIO1.Initialize()
            UGridIO1.Refresh()
            UGridIO1.Col = 0
            UGridIO1.Row = 0
            'If Order <> "Credit" Then OnScreenReport
            Row = 0
            mLoading = False
            Exit Sub
        End If
        mLoading = False
    End Sub

    Private Sub OnScreenReport()
        'MousePointer = 11
        Me.Cursor = Cursors.WaitCursor

        If ReportsMode("I") Then
            Printer.FontName = "Arial"
            Printer.FontSize = 18
            Printer.CurrentY = 100
            Printer.CurrentX = 0
            Row = 0
            TotDue = 0

            Printer.CurrentY = 800
            Printer.FontSize = 12

            If InvReports.Item = 4 Then
                'What's sold
                Printer.FontSize = 10
                Printer.FontBold = True
                Print(" Style Number      Manufacturer                  Status Quan        Description", TAB(109), "Sell Price")
                Printer.FontBold = False

                Dim cTa As CDataAccess
                cTa = Margin.DataAccess()
                cTa.Records_OpenSQL(SQL:=cTa.getFieldIndexSQL("SellDate", Val(Margin.SellDte)))
                Do While cTa.Records_Available()
                    If Val(Margin.SellDte) <> 0 Then
                        On Error Resume Next
                        If DateDiff("d", InvReports.dteStartDate.Value, Margin.SellDte) >= 0 Then
                            If Trim(Margin.Style) <> "STAIN" And Trim(Margin.Style) <> "DEL" And Trim(Margin.Style) <> "LAB" And Trim(Margin.Style) <> "TAX1" And Trim(Margin.Style) <> "TAX2" And Trim(Margin.Style) <> "NOTES" And Trim(Margin.Style) <> "SUB" And Trim(Margin.Style) <> "PAYMENT" And Trim(Margin.Status) <> "VOID" Then
                                ReadOut()
                                TotDue = TotDue + Margin.SellPrice
                            End If
                        End If
                    End If
                Loop
                cTa.Records_Close()
            End If
            Margin.Style = ""
            Margin.Vendor = ""
            Margin.Status = ""
            Margin.Quantity = 0
            Margin.Desc = "   Total --->"
            Margin.SellPrice = TotDue
            ReadOut()
        End If
CustNotFound:
        Lines = 0
        'MousePointer = 0
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cmdAllStores_Click(sender As Object, e As EventArgs) Handles cmdAllStores.Click
        Dim X As Integer
        Dim I As Integer, pHolding As cHolding
        Dim J As Integer
        Dim R As ADODB.Recordset

        X = StoresSld

        UGridIO1.Clear()
        UGridIO1.Refresh()
        'UGridIO1.MaxRows = 2
        Row = 0
        TotDue = 0

        For I = 1 To LicensedNoOfStores()
            StoresSld = I
            R = GetRecordsetBySQL("SELECT * FROM [Mail] WHERE [Tele]='" & CleanAni(Tele) & "'", , GetDatabaseAtLocation(I))

            Do While Not R.EOF
                pHolding = New cHolding

                pHolding.DataAccess.Records_OpenFieldIndexAtNumber("Index", R("index").Value, "LeaseNo")
                'If Holding.Load(Trim(Index), "#Index") Then
                If Not pHolding.DataAccess.Record_EOF Then
                    Do While pHolding.DataAccess.Records_Available
                        PoNo = 0
                        EnableControls(True)
                        If Trim(pHolding.Status) <> "V" Then
                            OrdTotal = 0
                            OrdTotal = Format(pHolding.Sale - pHolding.Deposit, "###,###.00")
                            TotDue = TotDue + OrdTotal
                            GetMarginRecords(pHolding.LeaseNo)
                            UGridIO1.Refresh()
                            If Not pHolding.DataAccess.Record_EOF Then Row = Row + 1
                        End If
                        'Holding.DataAccess.Records_MoveNext
                    Loop
                    txtBalDue.Text = CurrencyFormat(TotDue)
                End If

                DisposeDA(pHolding)
                R.MoveNext()
            Loop
        Next

        StoresSld = X
    End Sub

    Private Sub GetMarginRecords(ByVal LeaseNo As String)
        Dim cTa As CDataAccess, dT As Date
        Dim I As Integer

        TaxLoc = 0 ' Default to No Tax Applied.
        'lblRate(0).Tag = ""
        lblRate0.Tag = ""

        'Do While lblRate.UBound >= 1
        'Unload lblRate(lblRate.UBound)
        'Unload txtDiffTax(txtDiffTax.UBound)
        'Loop
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 7) = "lblRate" Then
                I = I + 1
            End If
        Next
        If I >= 1 Then
            For Each C As Control In Me.Controls
                If C.Name = "lblRate" & I Then
                    C.Hide()
                End If
                If C.Name = "txtDiffTax" & I Then
                    C.Hide()
                End If
            Next
        End If

        cTa = Margin.DataAccess()
        AskedForTaxRate = False
        TaxBackedOut = False

        cTa.DataBase = GetDatabaseAtLocation()
        cTa.Records_OpenSQL(SQL:=cTa.getFieldIndexSQL("SaleNo", Trim(LeaseNo), "MarginLine"))
        If cTa.Record_Count > MaxLines - 20 Then
            'MsgBox "This sale already has " & cTa.Record_Count & " lines." & vbCrLf & "This is approaching the maximum number of sale lines of " & MaxLines & "." & vbCrLf & "Please close this sale.", vbInformation, "Cannot adjust sale"
            MessageBox.Show("This sale already has " & cTa.Record_Count & " lines." & vbCrLf & "This is approaching the maximum number of sale lines of " & MaxLines & "." & vbCrLf & "Please close this sale.", "Cannot adjust sale", MessageBoxButtons.OK, MessageBoxIcon.Information)
            EnableControls(True, True)
        End If

        IsKit = False
        Do While cTa.Records_Available()
            Margin.cDataAccess_GetRecordSet(cTa.RS)
            SaleNo = Margin.SaleNo
            LastName = Trim(Margin.Name)
            If Margin.Index <> 0 Then Index = Margin.Index
            txtLocation.Text = Margin.Store
            Sales = Margin.Salesman

            If dT <> DateValue(Margin.SellDte) Then
                UGridIO1.SetValueDisplay(Row, 1, "SALE DATE:")
                UGridIO1.SetValueDisplay(Row, 2, Margin.SellDte)
                UGridIO1.SetValueDisplay(Row, 6, "Store #" & StoresSld)
                UGridIO1.Refresh()
                Row = Row + 1
            End If
            dT = Margin.SellDte

            If Trim(Margin.Style) = "PAYMENT" Then
                Margin.SellPrice = -Margin.SellPrice
            End If
            If Trim(Margin.Style) = "TAX1" Then
                TaxLoc = Margin.Quantity
            ElseIf Trim(Margin.Style) = "TAX2" Then
                'If lblRate(0).Tag = "" Then
                If lblRate0.Tag = "" Then
                    lblRate0.Tag = Margin.Quantity
                    lblRate0.Text = GetTax2Rate(Margin.Quantity)
                    'lblRate0.ToolTipText = GetTax2String(Margin.Quantity)
                    ToolTip1.SetToolTip(lblRate0, GetTax2String(Margin.Quantity))
                Else
                    CheckTaxLoc(Margin.Quantity)
                End If
                TaxLoc = -1
                '      If TaxLoc = 0 Then
                '        TaxLoc = Margin.Quantity + 1  ' bfh20090422 - not sure why this was out...  added if blocks
                '      End If
            End If
            ReadOut()
        Loop

        cTa.Records_Close()
        'Put in totals
        RowCheck()
        UGridIO1.SetValueDisplay(Row, 6, "         TOTAL DUE -->")
        UGridIO1.SetValueDisplay(Row, 8, Format(OrdTotal, "###,###.00"))
        UGridIO1.Refresh()
        lblPrevBal.Text = Format(OrdTotal, "Currency")
        Row = Row + 1
    End Sub

    Public Sub ReadOut()
        'grid gets loaded
        If ReportsMode("I") Then
            RowCheck()
            UGridIO1.SetValueDisplay(Row, 0, Margin.Style)
            UGridIO1.SetValueDisplay(Row, 1, Margin.Vendor)
            UGridIO1.SetValueDisplay(Row, 2, Margin.Status)
            UGridIO1.SetValueDisplay(Row, 3, Margin.Quantity)
            UGridIO1.SetValueDisplay(Row, 4, Margin.Desc)
            UGridIO1.SetValueDisplay(Row, 5, CurrencyFormat(Margin.SellPrice))
            UGridIO1.Refresh()
            Row = Row + 1
        ElseIf ReportsMode("H") Or OrderMode("Credit") Then
            If Microsoft.VisualBasic.Left(Margin.Desc, 25) = "PRICE WITH TAX BACKED OUT" Then
                TaxBackedOut = True
            End If
            Dim KI As Integer, KR As Double
            RowCheck()
            UGridIO1.SetValueDisplay(Row, 0, Margin.SaleNo)
            UGridIO1.SetValueDisplay(Row, 1, Margin.Style)
            UGridIO1.SetValueDisplay(Row, 2, Margin.Vendor)
            UGridIO1.SetValueDisplay(Row, 3, Margin.Location)
            UGridIO1.SetValueDisplay(Row, 4, Margin.Status)
            UGridIO1.SetValueDisplay(Row, 5, Margin.Quantity)
            UGridIO1.SetValueDisplay(Row, 6, Margin.Desc)
            UGridIO1.SetValueDisplay(Row, 7, CurrencyFormat(Margin.SellPrice))
            UGridIO1.SetValue(Row, 9, Margin.MarginLine)

            If Not IsItem(Margin.Style) Then
                IsKit = False
            Else
                If Not IsKit And Margin.SellPrice <> 0 Then
                    UGridIO1.SetValue(Row, 10, CurrencyFormat(Margin.SellPrice))
                ElseIf Not IsKit And Margin.SellPrice = 0 Then
                    IsKit = True
                    KitStart = Row
                    KitTotalCost = Margin.Cost
                    '        UGridIO1.SetValue Row, 10, CurrencyFormat(Margin.Cost)
                ElseIf IsKit And Margin.SellPrice = 0 Then
                    UGridIO1.SetValue(Row, 10, CurrencyFormat(Margin.Cost))
                    KitTotalCost = KitTotalCost + Margin.Cost
                ElseIf IsKit And Margin.SellPrice <> 0 Then
                    UGridIO1.SetValue(Row, 10, CurrencyFormat(Margin.Cost))
                    KitTotalCost = KitTotalCost + Margin.Cost
                    IsKit = False
                    If KitTotalCost <> 0 Then KR = Margin.SellPrice / KitTotalCost Else KR = 0
                    For KI = KitStart To Row
                        UGridIO1.SetValue(KI, 10, CurrencyFormat(GetPrice(UGridIO1.GetValue(KI, 10)) * KR))
                    Next
                End If
            End If
            UGridIO1.SetValue(Row, 11, Margin.Commission)  ' MJK 20131026
            UGridIO1.Refresh()

            Lines = Lines + 1
            Row = Row + 1
        End If
    End Sub

    Private Sub ChangeGrid2(ByVal Row As Integer, ByVal Col As Integer, ByVal NewVal As String)
        If NewVal <> UGridIO2.GetValue(Row, Col) Then
            UGridIO2.SetValueDisplay(Row, Col, NewVal)
        End If
    End Sub

    Private Sub WriteOut()
        WriteOutAddedItems = False
        For Row = 0 To Counter - 1
            If UGridIO2.GetValue(Row, 1) = "TAX1" Or UGridIO2.GetValue(Row, 1) = "TAX2" Then GoTo SkipRow
            Margin.SaleNo = UGridIO2.GetValue(Row, 0)           'add stuff
            Margin.Style = UGridIO2.GetValue(Row, 1)
            Margin.Vendor = UGridIO2.GetValue(Row, 2)
            Margin.Location = GetPrice(UGridIO2.GetValue(Row, 3))
            Margin.Status = UGridIO2.GetValue(Row, 4)
            Margin.Quantity = GetPrice(UGridIO2.GetValue(Row, 5))
            Margin.Index = GetPrice(Index)

            'BFH20081209 - Copied down DelDate for returns showing up correctly on Margin Report, Delivered by Salesperson
            '      If Margin.Status = "DELTW" Then
            If IsDelivered(Margin.Status) Then
                Margin.DDelDat = Today
            Else
                Margin.DDelDat = ""
                If Margin.Quantity < 0 Then
                    ' bfh20080517 - no longer set deldate for cancelled items.
                    '          Margin.DDelDat = Date
                ElseIf Margin.Status = "ST" Then
                    MessageBox.Show("To set up delivery dates for new items, go to the Check Order Status panel.", "Delivery Dates", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

            Margin.Desc = UGridIO2.GetValue(Row, 6)
            Margin.SellPrice = GetPrice(UGridIO2.GetValue(Row, 7))
            Margin.Salesman = Sales
            Margin.GM = 0

            If Margin.Quantity < 0 Then                 ' returned items
                ChgMarginStatus()
                ' These changes get saved in AddMarginLine.
                Margin.Detail = DetailRec(Row)

                Select Case Margin.Status
                    Case "SS", "SSREC", "FND"
                        Margin.RN = 0
                    Case Else
                        Margin.RN = InvRn(Row)
                End Select
                Margin.Cost = Cost(Row)
                Margin.ItemFreight = Freight(Row)
                '        Margin.DeptNo = GetDeptNoFromRn(Margin.Rn)
                'BFH20140227 - this line removed...  If you set "C" on return item line, the commissions report prints out two lines.  The new xDEL line must not be marked as commissioned, even if the original was.
                'Margin.Commission = UGridIO2.GetValue(Row, 11)  ' MJK 20131026
            ElseIf Margin.SellPrice = 0 Then
                ' didn't cause an addeditem in effect
            Else                                        ' added items
                Margin.Cost = Cost(Row)   ' BFH20140223 - wasn't resetting cost from first item.
                If IsItem(Margin.Style) Then Margin.GM = CalculateGM(Margin.SellPrice, Margin.Cost + Margin.ItemFreight, , 0)
                WriteOutAddedItems = True
            End If

            AddMarginLine()                          ' Save margin record, create detail, etc.
SkipRow:
        Next
    End Sub

    Private Sub AddMarginLine()
        Margin.MarginLine = 0         'add change
        Margin.DeptNo = Depts(Row)    'Change this for multi-digit departments
        If Margin.Vendor = "" Then
            Margin.VendorNo = "000"
        Else
            Margin.VendorNo = Format(Trim(Vends(Row)), "000")
        End If
        'Margin.CommPd = 0             ' Commission has not been paid on the adjustment.
        Margin.CommPd = #1/1/0001#
        'Margin.Commission = ""        ' No. Commission is carried through from parent row for returns.
        If Margin.Commission <> "" And Margin.Quantity >= 0 Then Margin.Commission = "" ' Only returns could possibly have "C" here.  Others are unfortunate bleed-through cases.

        If Trim(Margin.Style) <> "TAX1" And Trim(Margin.Style) <> "SUB" And Trim(Margin.Style) <> "--- Adj ---" Then

            If Margin.Quantity < 0 Then   'returns
                Margin.Status = "x" + Margin.Status

                '04-14-2003 took out quantity added - sign
                Margin.Cost = -Cost(Row) '* Val(Quantity(row)) 'changed - on 10-08-01
                Margin.ItemFreight = -Freight(Row) ' * Val(Quantity(row)) 'changed - on 10-08-01
            Else
                If Val(Cost(Row)) = 0 Then Cost(Row) = 0
                ' bfh20060123 - getitemcost
                Margin.Cost = GetItemCost(Margin.Style, Margin.Location, , Margin.Quantity) '  Cost(Row) * Val(Quantity(Row))
                Margin.ItemFreight = Freight(Row) * Margin.Quantity 'Val(Quantity(Row))
            End If
            Margin.SellDte = DateFormat(Now)
            Margin.RN = InvRn(Row)    ' This is wrong if they've changed style!
            If Margin.Location = 0 Then Margin.Location = GetPrice(txtLocation.Text)
        End If
        If IsIn(Trim(Margin.Style), "TAX1", "SUB", "--- Adj ---") Then
            Margin.Location = 0
        End If

        Margin.DataAccess().Records_Add()
        Margin.cDataAccess_SetRecordSet(Margin.DataAccess.RS)
        Margin.Save()
        MarginNo = Margin.MarginLine

        If Not IsIn(Trim(Margin.Style), "TAX1", "SUB", "--- Adj ---") Then
            Margin.Load(CStr(MarginNo), "#MarginLine")
            WhatToDo()  ' Updates inventory count, creates detail, etc.
            Margin.Save()    ' Save the new detail number.
        End If
        Margin.DataAccess.Records_Close()
    End Sub

    Private Sub WhatToDo()
        'process added items, Reduce Inventory & S/O
        Select Case Trim(Margin.Status)
            Case "ST", "SO", "SOREC", "LAW", "PO", "DELTW"
                ' creates a PO via MakePO inside of here for SO
                AddToStock()
    '***** if its a S/S it doesn't go into inventory or detail ***
            Case "SS" : MakePo()
        End Select
    End Sub

    Private Sub AddToStock()
        On Error GoTo HandleErr
        Dim InvData As CInvRec, InvDetail As CInventoryDetail

        InvData = New CInvRec
        InvDetail = New CInventoryDetail

        If Trim(Margin.RN) = "" Then GoTo ExitSub
        If Not InvData.Load(CStr(InvRn(Row)), "#Rn") Then Exit Sub    ' Couldn't load Style

        If Trim(Margin.Status) = "SO" Or Trim(Margin.Status) = "PO" Then
            InvData.ItemsSold(Margin.Quantity, Today)

            If Trim(Margin.Status) = "PO" Then InvData.PoSold = InvData.PoSold + Margin.Quantity

            InvData.Save()

            'fix detail in location
            InvDetail.SetLocationQuantity(Margin.Location, Margin.Quantity)
            AddDetail(Margin, InvData, InvDetail)
            If Trim(Margin.Status) = "SO" Then
                MakePo()
                ' made detail show po number in name like normal po entries in detail do...  bfh20050621
                InvDetail.Name = Trim(Trim(InvDetail.Name) & " " & PoNo)
                InvDetail.Save()
            End If
            GoTo ExitSub
        ElseIf Trim(Margin.Status) <> "LAW" Then
            InvData.Available = InvData.Available - Margin.Quantity
            If IsDelivered(Margin.Status) Then
                InvData.OnHand = InvData.OnHand - Margin.Quantity
            End If
        End If

        'Reduced written void sales
        InvData.ItemsSold(Margin.Quantity, Today)
        '  If NoOnHand = "N" Then   ' This is never N, so why do it?
        '    InvData.OnHand = Val(InvData.OnHand) - Val(Margin.Quantity)
        '  End If

        If Trim(Margin.Status) <> "LAW" Then
            InvData.AddLocationQuantity(Margin.Location, -Margin.Quantity)
        End If

        'Lay-a-Way need this
        InvDetail.SetLocationQuantity(Margin.Location, Margin.Quantity)

        InvData.Save()

        AddDetail(Margin, InvData, InvDetail)
ExitSub:

        DisposeDA(InvData, InvDetail)
        Exit Sub

HandleErr:
        MessageBox.Show(Err.Description & ", " & Err.Source, Err.ToString, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        If Err.Number = 13 Then Resume Next
    End Sub

    Private Sub AddDetail(ByRef Margin As CGrossMargin, ByRef InvData As CInvRec, ByRef InvDetail As CInventoryDetail)
        On Error GoTo HandleErr
        InvDetail.Name = Margin.Name  ' Transplanted from GetMarginRecords MJK 20030428
        InvDetail.Style = Margin.Style
        InvDetail.Lease1 = Trim(Margin.SaleNo)
        InvDetail.Trans = "NS"
        InvDetail.Store = StoresSld  ' margin.location?

        InvDetail.InvRn = InvData.RN
        InvDetail.MarginRn = Margin.MarginLine

        If IsDelivered(Margin.Status) Then
            InvDetail.DDate1 = DateFormat(Now)
            InvDetail.SO1 = 0
            InvDetail.Trans = "DS"
            InvData.OnHand = InvData.OnHand - Val(BillOSale.QueryQuan(BillOSale.X))
        ElseIf Trim(Margin.Status) = "SO" Or Trim(Margin.Status) = "NOTES" Then
            '        InvDetail.Name = Trim(Margin.Name + " " & Trim(PoNo))
            InvDetail.AmtS1 = 0
            InvDetail.Ns1 = 0
            InvDetail.SO1 = Margin.Quantity
            InvDetail.LAW = 0
        ElseIf Trim(Margin.Status) = "LAW" Then
            InvDetail.AmtS1 = 0
            InvDetail.Ns1 = 0
            InvDetail.SO1 = 0
            InvDetail.LAW = Margin.Quantity
        ElseIf Trim(Margin.Status) = "ST" Or IsDelivered(Margin.Status) Then
            InvDetail.AmtS1 = Margin.Quantity
            InvDetail.Ns1 = 0
            InvDetail.SO1 = 0
            InvDetail.LAW = 0
        ElseIf Trim(Margin.Status) = "SOREC" Or Trim(Margin.Status) = "POREC" Then
            InvDetail.Name = Margin.Name + " " + Trim(Margin.SaleNo)
            InvDetail.DDate1 = DateFormat(Now)
            InvDetail.Misc = 0
            InvDetail.SO1 = 0
            InvDetail.Trans = "IN"
            InvDetail.Ns1 = Margin.Quantity
            InvDetail.ItemCost = Margin.Cost    ' bfh20060124
        ElseIf Trim(Margin.Status) = "PO" Then
            InvDetail.Trans = "PO"
            InvDetail.AmtS1 = 0
            InvDetail.Ns1 = 0
            InvDetail.SO1 = Margin.Quantity
            InvDetail.LAW = 0
        End If

        ' This is always a new record.
        InvDetail.DataAccess.Records_Add()
        InvDetail.Save()
        Detail = InvDetail.DetailID
        Margin.Detail = Detail
        '    NoOnHand = ""  ' It's always "", don't bother clearing it.

        Exit Sub
HandleErr:
        'MsgBox "ERROR in Detail ONScreenReport.AddDetail: " & Err.Description & ", " & Err.Source & ", " & Err.Number
        MessageBox.Show("ERROR in Detail ONScreenReport.AddDetail: " & Err.Description & ", " & Err.Source & ", " & Err.Number)
        Resume Next
    End Sub

    Private Sub MakePo()
        Dim PO As cPODetail

        If Microsoft.VisualBasic.Left(Margin.Style, 4) = KIT_PFX Then Exit Sub ' BFH20100826

        If PoNo = 0 Or Trim(LastSale) <> Trim(Margin.SaleNo) Or Trim(LastMfg) <> Trim(Margin.Vendor) Then 'first item & sale
            PoNo = GetPoNo()
        End If
        On Error GoTo HandleErr
        PO = New cPODetail
        PO.PoNo = PoNo
        PO.SaleNo = Trim(Margin.SaleNo)
        LastSale = Trim(Margin.SaleNo)
        PO.PoDate = DateFormat(Now)
        PO.Name = Margin.Name
        PO.Vendor = Margin.Vendor
        LastMfg = Trim(Margin.Vendor)

        PO.InitialQuantity = Margin.Quantity
        PO.Quantity = Margin.Quantity
        PO.Style = Margin.Style
        PO.Desc = Margin.Desc
        PO.Cost = CurrencyFormat(Margin.Cost)
        PO.Location = IIf(Margin.Store = 0, StoresSld, Margin.Store) 'changed 03-02-2002, again 20131115

        If IsDoddsLtd And PO.Location = 0 Then MessageBox.Show("bad po, loc=0")

        PO.SoldTo = "1"
        PO.ShipTo = "2"
        If StoreSettings.bPOSpecialInstr Then
            PO.Note1 = "1"
            PO.Note2 = "1"
        Else
            PO.Note1 = "0"
            PO.Note2 = "0"
        End If
        PO.Note3 = "0"
        PO.Note4 = "0"
        PO.PoNotes = ""
        PO.AckInv = ""
        PO.Posted = ""
        PO.PrintPo = ""
        PO.wCost = "1" 'Print w/Cost
        If StoreSettings.bPrintPoNoCost Then PO.wCost = "0"
        PO.RN = Val(InvRn(Row))
        PO.Detail = Detail
        PO.MarginLine = MarginNo
        PO.PoDate = DateFormat(Now)

        PO.Save()
        DisposeDA(PO)
        Exit Sub
HandleErr:
        If Err.Number = 13 Then Resume Next ' type mismatch
    End Sub

    Private Sub ChgMarginStatus()
        'change Original record to void
        Dim Margin As New CGrossMargin, S As String
        If Margin.Load(UGridIO2.GetValue(Row, 10), "#MarginLine") Then  '  CStr(MarginnoRec2(Row))
            Select Case Trim(Margin.Status)
                Case "SS"
                    VoidPO(Margin)
                    Margin.Status = "xSS"
                Case "SO"
                    VoidPO(Margin)
                    ReturnToStock(Margin)
                    Margin.Status = "xSO"
                Case "ST", "PO", "POREC", "SOREC", "DELTW", "LAW", "DELST", "DELPO", "DELPOR", "DELPOREC", "DEL"
                    ReturnToStock(Margin)
                    Margin.Status = "x" & Margin.Status
                Case "SSLAW"
                    Margin.Status = "xSLAW"
                Case "FND"
                    Margin.Status = "xFND"
                Case "SSREC"
                    MessageBox.Show("You MUST manually add this returned item into the inventory!")
                    Margin.Status = "xSSRC"
                Case "DELSO", "DELSOR", "DELSOREC", "DELSS", "DELSSR", "DELSSREC", "DELFND", "DELFN"
                    ReturnToStock(Margin)
                    Margin.Status = "x" & Margin.Status
            End Select

            Margin.Save()
        End If
        DetailRec(Row) = Margin.Detail
        InvRn(Row) = Margin.RN
        Cost(Row) = Margin.Cost
        Freight(Row) = Margin.ItemFreight
        Depts(Row) = Margin.DeptNo
        Vends(Row) = Margin.VendorNo

        DisposeDA(Margin)
    End Sub

    Private Sub ReturnToStock(ByRef Margin As CGrossMargin)
        Margin.ReturnToStock(Today, True)
    End Sub

    Private Sub VoidPO(ByVal vMargin As CGrossMargin)
        Dim X As Integer, F As String, S As String, R As ADODB.Recordset
        '  F = ""
        '  F = F & "WHERE LeaseNo='" & Trim(Margin.SaleNo) & "' "
        '  F = F & "AND Style='" & Trim(Margin.Style) & "' "
        '  F = F & "AND InitialQuantity=" & -(Margin.Quantity) & " "
        '  F = F & "AND PrintPO<>'v'"

        F = ""
        F = F & "WHERE MarginLine=" & vMargin.MarginLine

        S = "SELECT POID FROM PO " & F
        R = GetRecordsetBySQL(S, , GetDatabaseInventory)

        S = "UPDATE PO SET PrintPO='v' " & F
        If Not R.EOF Then S = S & " AND POID=" & R("POID").Value     ' only do the first

        DisposeDA(R)
        ExecuteRecordsetBySQL(S, , GetDatabaseInventory)
    End Sub

    Private Sub CheckTaxLoc(ByVal tL As Integer)
        Dim X As Integer, A As Integer, B As Integer
        Dim CountlblRate As Integer

        If TaxLocHandled(tL) Then Exit Sub

        'X = lblRate.UBound + 1
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 7) = "lblRate" Then
                CountlblRate = CountlblRate + 1
            End If
        Next
        X = CountlblRate + 1
        'Load txtDiffTax(X)
        'txtDiffTax(X).Visible = True
        For Each C As Control In Me.Controls
            If C.Name = "txtDiffTax" & X Then
                C.Visible = True
                C.Text = ""
                'txtDiffTax(X).Top = txtDiffTax(X - 1).Top + txtDiffTax(X).Height + 60
                'txtBalDue.Top = txtDiffTax(X).Top + txtDiffTax(X).Height + 60
                txtBalDue.Top = C.Top + C.Height + 6
                Exit For
            End If
        Next
        'Load lblRate(X)
        'lblRate(X).Visible = True
        For Each C As Control In Me.Controls
            If C.Name = "lblRate" & X Then
                C.Visible = True
                C.Tag = tL
                If tL = 0 Then
                    C.Text = StoreSettings.SalesTax
                    'lblRate(X).ToolTipText = "(default)"
                    ToolTip1.SetToolTip(C, "(default)")
                Else
                    C.Text = GetTax2Rate(tL)
                    'lblRate(X).ToolTipText = GetTax2String(tL)
                    ToolTip1.SetToolTip(C, GetTax2String(tL))
                End If
                'lblRate(X).Top = lblRate(X - 1).Top + lblRate(X).Height + 60
                'lblBalDue.Top = lblRate(X).Top + lblRate(X).Height + 120
                lblBalDue.Top = C.Top + C.Height + 12
                Exit For
            End If
        Next

        'lblRate(X).Tag = tL
        'If tL = 0 Then
        '    lblRate(X) = StoreSettings.SalesTax
        '    lblRate(X).ToolTipText = "(default)"
        'Else
        '    lblRate(X) = GetTax2Rate(tL)
        '    lblRate(X).ToolTipText = GetTax2String(tL)
        'End If
        'lblRate(X).Top = lblRate(X - 1).Top + lblRate(X).Height + 60

        'txtDiffTax(X) = ""
        'txtDiffTax(X).Top = txtDiffTax(X - 1).Top + txtDiffTax(X).Height + 60

        'lblBalDue.Top = lblRate(X).Top + lblRate(X).Height + 120
        'txtBalDue.Top = txtDiffTax(X).Top + txtDiffTax(X).Height + 60

        A = fraControls2.Top + fraControls2.Height + 12
        B = txtBalDue.Top + txtBalDue.Height + 12
        Height = IIf(A > B, A, B) + (Height - Me.ClientSize.Height)
    End Sub

    Private Function TaxLocHandled(ByVal tL As Integer) As Boolean
        Dim K As Integer
        Dim I As Integer

        If lblRate0.Tag = "" Then Exit Function
        'For K = txtDiffTax.LBound To txtDiffTax.UBound
        '    If Val(lblRate(K)) = GetTax2Rate(tL) Then TaxLocHandled = True : Exit Function
        'Next
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "txtDiffTax" Then
                I = I + 1
            End If
        Next
        For K = 0 To I
            For Each C As Control In Me.Controls
                If Mid(C.Name, 1, 7) = "lblRate" Then
                    If Val(C.Text) = GetTax2Rate(tL) Then TaxLocHandled = True : Exit Function
                End If
            Next
        Next
    End Function

    Public Sub CustomerAdjustment()
        ' Load and show this form..
        ' Customer and Sale information is loaded by MailCheck's Sale Found events.
        Show()
        'cmdNext.Value = True
        'cmdNext.PerformClick()
        cmdNext_Click(cmdNext, New EventArgs)
    End Sub

    Public Sub CustomerHistory()
        ' Load and show this form..
        ' Customer and sale information is loaded by MailCheck's Customer Found events.
        'Form_Load
        OnScreenReport_Load(Me, New EventArgs)
        Show()
        'cmdNext.Value = True
        cmdNext.PerformClick()
    End Sub

    Private Sub MailCheckRef_CustomerFound(MailIndex As Integer, ByRef Cancel As Boolean) Handles MailCheckRef.CustomerFound
        ' For reportsmode("H"), load customer info and continue report.
        If ReportsMode("H") Then
            Row = 0
            Balance = 0
            TotDue = 0
            txtBalDue.Visible = True
            lblBalDue.Visible = True
            LoadCustomerInfo(MailIndex)

            g_Holding.DataAccess.Records_OpenFieldIndexAtNumber("Index", Trim(Index), "LeaseNo")
            'If Holding.Load(Trim(Index), "#Index") Then
            If Not g_Holding.DataAccess.Record_EOF Then
                Do While g_Holding.DataAccess.Records_Available
                    PoNo = 0
                    EnableControls(True)
                    If Trim(g_Holding.Status) <> "V" Then
                        OrdTotal = 0
                        OrdTotal = Format(g_Holding.Sale - g_Holding.Deposit, "###,###.00")
                        TotDue = TotDue + OrdTotal
                        GetMarginRecords(g_Holding.LeaseNo)
                        UGridIO1.Refresh()
                        If Not g_Holding.DataAccess.Record_EOF Then Row = Row + 1
                    End If
                    'Holding.DataAccess.Records_MoveNext
                Loop
                txtBalDue.Text = CurrencyFormat(TotDue)
            End If

            SaleFound = True
            Cancel = True
            cmdApply.Enabled = True
        End If
    End Sub

    Private Sub UGridIO1_DoubleClick(sender As Object, e As EventArgs) Handles UGridIO1.DoubleClick
        ' This is a good place to stop.
        If ReportsMode("H") And Trim(UGridIO1.GetValue(UGridIO1.Row, 0)) <> "" Then
            ' Load up billosale
            Order = "E"
            MailCheck.optSaleNo.Checked = True
            MailCheck.InputBox.Text = Trim(UGridIO1.GetValue(UGridIO1.Row, 0))
            'MailCheck.cmdOK.Value = True
            MailCheck.cmdOK.PerformClick()
            BillOSale.Show()
            BillOSale.cmdApplyBillOSale.Enabled = False
            BillOSale.cmdCancel.Enabled = False

            BillOSale.BillOSale2_Show()
            BillOSale.cmdClear.Enabled = False
            BillOSale.cmdNextSale.Enabled = False
            BillOSale.cmdProcessSale.Enabled = False
            BillOSale.ScanDn.Enabled = False
            BillOSale.ScanUp.Enabled = False
            BillOSale.UGridIO1.GetDBGrid.AllowUpdate = False
            BillOSale.cmdMainMenu.Text = "Back"

            modProgramState.Order = ""
        End If
    End Sub

    Private Sub UGridIO2_AfterColUpdate(ColIndex As Integer) Handles UGridIO2.AfterColUpdate
        'Debug.Print ColIndex, UGridIO2.GetValue(UGridIO2.row, ColIndex), UGridIO2.GetDBGrid.Text
        '  Dim Row As Integer
        If ColIndex = BillColumns.eQuant + 1 Then
            UGridIO2.SetValue(UGridIO2.Row, 7, Format(GetPrice(UGridIO2.GetValue(UGridIO2.Row, 9)) * GetPrice(UGridIO2.Text), "###,##0.00"))
            ' Don't change the per-item price.
            UGridIO2.Refresh(True)
            Recalculate()
        ElseIf ColIndex = BillColumns.ePrice + 1 Then
            If Val(UGridIO2.GetValue(UGridIO2.Row, 5)) = 0 Then
                UGridIO2.SetValue(UGridIO2.Row, 9, UGridIO2.Text)   ' Per-unit price
            Else
                UGridIO2.SetValue(UGridIO2.Row, 9, GetPrice(UGridIO2.Text) / UGridIO2.GetValue(UGridIO2.Row, 5)) ' Per-unit price
            End If
            '.SetValueDisplay Row, 7, Format(UGridIO2.GetDBGrid.Text, "###,##0.00")
            UGridIO2.Refresh(True)
            Recalculate()
        Else
            UGridIO2.SetValue(UGridIO2.Row, ColIndex, UGridIO2.Text)
        End If
    End Sub

    Public Sub Recalculate()
        Dim Style As String, Desc As String, I As Integer, ML As Integer
        Dim Cur As Decimal, Tax As Decimal, P As Decimal
        Dim NewDiff As Decimal
        Dim K As Integer, KR As Double
        Dim DiffTaxCount As Integer
        '  Dim RemTax As Currency, AddTax As Currency

        'For K = txtDiffTax.LBound To txtDiffTax.UBound
        '    txtDiffTax(K) = CurrencyFormat(0)
        'Next
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "txtDiffTax" Then
                C.Text = CurrencyFormat(0)
                DiffTaxCount = DiffTaxCount + 1
                'Exit For
            End If
        Next

        Balance = 0
        NonTaxable = 0
        For I = 0 To Counter - 1
            Cur = GetPrice(UGridIO2.GetValue(I, 7))
            Balance = Balance + Cur
            Style = UGridIO2.GetValue(I, 1)
            Desc = UGridIO2.GetValue(I, 6)
            ML = Val(UGridIO2.GetValue(I, 10))
            If (Style = "DEL" And Not StoreSettings.bDeliveryTaxable) Or (Style = "LAB" And Not StoreSettings.bLaborTaxable) Then
                NonTaxable = NonTaxable + Cur     ' Nontaxable
            ElseIf Style = "TAX1" Or Style = "TAX2" Then
                Tax = Tax + Cur
                Balance = Balance - Cur
            ElseIf Style = "NOTES" And Microsoft.VisualBasic.Left(Desc, 25) = "PRICE WITH TAX BACKED OUT" Then
                Balance = Balance - Cur
            Else
                ' Taxable
                NewDiff = Cur
                If Cur < 0 And TaxBackedOut Then
                    '        NewDiff = NewDiff / (1 + Rate)
                    Balance = GetPrice(Balance - Cur + NewDiff)
                End If

                If Not GetSalesTax() Then  ' Get sales tax rate
                    P = 0
                    'For K = txtDiffTax.LBound To txtDiffTax.UBound
                    '    KR = Val(lblRate(K))
                    '    txtDiffTax(K) = CurrencyFormat(GetPrice(txtDiffTax(K)) + (NewDiff * KR))
                    '    P = Round(P + GetPrice(NewDiff * KR), 2)
                    '    If KR = 0 Then NonTaxable = NonTaxable + Cur
                    'Next
                    For K = 0 To DiffTaxCount
                        For Each C As Control In Me.Controls
                            If C.Name = "lblRate" & K Then
                                KR = Val(C.Text)
                            ElseIf C.Name = "txtDiffTax" & K Then
                                C.Text = CurrencyFormat(GetPrice(C.Text) + (NewDiff * KR))
                                P = Math.Round(P + GetPrice(NewDiff * KR), 2)
                            End If
                            If KR = 0 Then NonTaxable = NonTaxable + Cur
                        Next

                    Next
                Else
                    P = GetPrice(NewDiff * Rate)
                    'txtDiffTax(0) = CurrencyFormat(GetPrice(txtDiffTax(0)) + P)
                    txtDiffTax0.Text = CurrencyFormat(GetPrice(txtDiffTax0.Text) + P)
                    If Val(Rate) = 0 Then NonTaxable = NonTaxable + Cur
                End If
                ' hard to differentiate..  But, notes can be negative and we must allow to remove them and their tax
                '      If P < 0 Or Style = "NOTES" Then RemTax = RemTax + P Else AddTax = AddTax + P
            End If
            If I <> Counter - 1 Then UGridIO2.SetValueDisplay(I, 8, "")
        Next

        ' can't return more than we've charged!
        If TaxLoc = 1 Then
            If GetPrice(txtDiffTax0.Text) < 0 Then
                '        RemTax = FitRange(-SaleTax1Amount, RemTax, 0)
                '      txtDiffTax(0) = CurrencyFormat(FitRange(-SaleTax1Amount, RemTax, 0))
                If Not InRange(-SaleTax1Amount(), GetPrice(txtDiffTax0.Text), 0) Then
                    Dim MsgTxt As String
                    MsgTxt = MsgTxt & "Can't refund more tax than charged." & vbCrLf
                    MsgTxt = MsgTxt & "Maximum tax refund is " & CurrencyFormat(SaleTax1Amount) & "." & vbCrLf
                    '        MsgTxt = MsgTxt & vbCrLf & "Adjustment Cancelled."
                    MessageBox.Show(MsgTxt, "Adjustment Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation) ' "Adjustment Cancelled"
                    '        cmdApply.Enabled = False
                End If
                txtDiffTax0.Text = CurrencyFormat(FitRange(-SaleTax1Amount(), GetPrice(txtDiffTax0.Text), 0))
            End If
        Else
            '            For K = txtDiffTax.LBound To txtDiffTax.UBound
            '                If GetPrice(txtDiffTax(K)) < 0 Then
            '                    If Not InRange(-SaleTax2Amount(lblRate(K).Tag), GetPrice(txtDiffTax(K)), GetPrice(txtDiffTax(K))) Then
            '                        MsgBox "Can't refund more tax than charged." & vbCrLf & "Maximum tax refund is " & CurrencyFormat(SaleTax2Amount(lblRate(K).Tag)) & "." & vbCrLf2 & "Adjustment Cancelled.", vbExclamation, "Adjustment Cancelled"
            ''        cmdApply.Enabled = False
            '                    End If
            '                    txtDiffTax(K) = CurrencyFormat(FitRange(-SaleTax2Amount(lblRate(K).Tag), GetPrice(txtDiffTax(K)), GetPrice(txtDiffTax(K))))
            '                    '          RemTax = RemTax + GetPrice(txtDiffTax(K))
            '                End If
            '            Next
            For K = 0 To DiffTaxCount
                For Each C As Control In Me.Controls
                    If C.Name = "txtDiffTax" & K And GetPrice(C.Text) < 0 Then
                        For Each Cc As Control In Me.Controls
                            If Cc.Name = "lblRate" & K Then
                                If Not InRange(-SaleTax2Amount(Cc.Tag), GetPrice(C.Text), GetPrice(C.Text)) Then
                                    MessageBox.Show("Can't refund more tax than charged." & vbCrLf & "Maximum tax refund is " & CurrencyFormat(SaleTax2Amount(Cc.Tag)) & "." & vbCrLf2 & "Adjustment Cancelled.", "Adjustment Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                End If
                            End If
                        Next
                        For Each Ccc As Control In Me.Controls
                            If C.Name = "lblRate" & K Then
                                C.Text = CurrencyFormat(FitRange(-SaleTax2Amount(Ccc.Tag), GetPrice(C.Text), GetPrice(C.Text)))
                            End If
                        Next

                    End If
                Next
            Next

        End If

        Tax = 0
        'For K = txtDiffTax.LBound To txtDiffTax.UBound
        '    Tax = Tax + Round(GetPrice(txtDiffTax(K)), 2)
        'Next
        For K = 0 To DiffTaxCount - 1
            For Each C As Control In Me.Controls
                If C.Name = "txtDiffTax" & K Then
                    Tax = Tax + Math.Round(GetPrice(C.Text), 2)
                End If
            Next
        Next

        UGridIO2.SetValueDisplay(Counter - 1, 8, CurrencyFormat(Balance))
        If TaxLoc = 1 Then txtDiffTax0.Text = CurrencyFormat(Tax)
        txtBalDue.Text = CurrencyFormat(GetPrice(lblPrevBal.Text) + Balance + IIf(TaxBackedOut, 0, Tax))
        TotTax = Tax
    End Sub

    Public Function SaleTax2Amount(Optional ByVal Zone As Integer = 0) As Decimal
        Dim I As Integer
        For I = 1 To UGridIO1.LastRowUsed
            If Trim(UGridIO1.GetValue(I, 1)) = "TAX2" Then
                If Zone = 0 Or (Zone = UGridIO1.GetValue(I, 5)) Then
                    SaleTax2Amount = SaleTax2Amount + GetPrice(UGridIO1.GetValue(I, 7))
                End If
            End If
        Next
    End Function

    Public Function SaleTax1Amount() As Decimal
        Dim I As Integer
        For I = 1 To UGridIO1.LastRowUsed
            If Trim(UGridIO1.GetValue(I, 1)) = "TAX1" Then SaleTax1Amount = SaleTax1Amount + GetPrice(UGridIO1.GetValue(I, 7))
        Next
    End Function

    Private Function GetSalesTax() As Boolean
        If TaxLoc < 0 Then Exit Function
        GetSalesTax = True
        If TaxLoc = 0 Then
            Rate = 0
            'If cmdAdd.Value Then AskForTaxRate : GetSalesTax = False  --> cmdAdd is a button in vb6. But in vb.net to get the same result, use checkbox. Read note in SelectEntry sub of AddOnAcc.vb form for clarity.
            If cmdAdd.Checked = True Then AskForTaxRate() : GetSalesTax = False
        ElseIf TaxLoc = 1 Then
            Rate = GetStoreTax1()
        Else
            Rate = 0
            GetSalesTax = False
        End If
        Rate = Val(Rate)
    End Function

    Private Sub AskForTaxRate()
        Dim ST2L As Object, S As Object, N As Integer
        Dim Taxes() As Object, Tax As Object

        If AskedForTaxRate Then Exit Sub
        AskedForTaxRate = True
        ST2L = QuerySalesTax2List()
        ReDim Taxes(0 To (1 + SalesTax2Count()))
        Taxes(0) = StoreSettings.SalesTax & " (Default)"
        N = 1
        If SalesTax2Count() > 0 Then
            For Each S In ST2L
                Taxes(N) = ST2L(N - 1)
                N = N + 1
            Next
        End If
        Taxes(N) = "Non-Taxable"

        Tax = SelectOptionArray("Tax Rate", frmSelectOption.ESelOpts.SelOpt_List, Taxes, "Se&lect")

        If Tax <= N Then
            TaxLoc = Tax
            If Tax = 1 Then
                Rate = StoreSettings.SalesTax
                lblRate0.Text = Rate
                lblRate0.Tag = "" & (Tax - 1)
            Else
                Rate = QuerySalesTax2(TaxLoc - 2)
                lblRate0.Text = Rate
                lblRate0.Tag = "" & (Tax - 1)
            End If
        End If
    End Sub

    Private Sub OnScreenReport_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'QueryUnload event of vb6
        'If UnloadMode = vbFormControlMenu Then
        If FromCmdMenu = False Then
            If e.CloseReason = CloseReason.UserClosing Then

                If cmdMenu2.Visible Then
                    'cmdMenu2.Value = True  ' This unloads us and does other cleanup..
                    'cmdMenu2.PerformClick()
                    cmdMenu2_Click(cmdMenu2, New EventArgs)
                ElseIf cmdMenu.Visible Then
                    'cmdMenu.Value = True  ' This unloads us and does other cleanup..
                    'cmdMenu.PerformClick()
                    cmdMenu_Click(cmdMenu, New EventArgs)
                Else
                    'Cancel = True
                    e.Cancel = True
                End If
            End If
        End If

        'Unload event of vb6
        DisposeDA(MailCheckRef, Margin)
        'Unload BillOSale
        BillOSale.Close()
        'Unload Me
        If FromCmdMenu = False Then
            Me.Close()
        End If
    End Sub

    Private Function LoadCustomerInfo(ByVal MailIndex As Integer) As Boolean
        'getmail name and index
        Dim RS As ADODB.Recordset

        RS = getRecordsetByTableLabelIndexNumber("Mail", "Index", CStr(MailIndex))
        If Not RS.EOF Then
            CopyMailRecordsetToMailNew(RS, Mail)
            Index = Trim(Mail.Index)
            LastName = Mail.Last
            Tele = Mail.Tele
            'Caption = "Customer Consolidated History:  " & Mail.Last & "  " & Mail.Tele
            Text = "Customer Consolidated History:  " & Mail.Last & "  " & Mail.Tele
            LoadCustomerInfo = True
        Else
            'MsgBox "This customer could not be found in the database.", vbExclamation
            MessageBox.Show("This customer could not be found in the database.")
            Index = ""
            LastName = ""
            Tele = ""
            '    Telephone = False
            LoadCustomerInfo = False
        End If
        RS.Close()
        RS = Nothing
    End Function

    Private Property CurrentLine() As Integer
        Get
            CurrentLine = mCurrentLine
        End Get
        Set(value As Integer)
            mCurrentLine = value
        End Set
    End Property

    Private Sub EnableControls(ByVal OnOff As Boolean, Optional ByVal Processed As Boolean = False)
        'MousePointer = IIf(OnOff, vbDefault, vbHourglass)
        Cursor = IIf(OnOff, Cursors.Default, Cursors.WaitCursor)

        UGridIO1.GetDBGrid.Enabled = OnOff
        UGridIO2.GetDBGrid.Enabled = OnOff
        cmdNext.Enabled = OnOff
        cmdPrint.Enabled = OnOff
        cmdMenu.Enabled = OnOff
        cmdMenu2.Enabled = OnOff
        cmdNext2.Enabled = OnOff

        cmdAdd.Enabled = OnOff And Not Processed
        cmdReturn.Enabled = OnOff And Not Processed
        cmdApply.Enabled = OnOff And Not Processed
    End Sub

    Private Sub MailCheckRef_SaleFound(Holding As cHolding, ByRef Cancel As Boolean) Handles MailCheckRef.SaleFound
        Dim S As String

        S = Holding.Status
        If S = "V" Then
            MessageBox.Show("This sale is void and can not be changed.")
            Cancel = True
            Exit Sub
        End If

        SalePackageUpdate(Holding.LeaseNo) ' refresh it first... why not..


        If IsIn(S, "D", "C", "B") Then
            If Not AllowAdjustDel Then
                MessageBox.Show("This sale is delivered and must be voided to change!")
                Cancel = True
                Exit Sub
            End If
            WasDelSale = True
        Else
            WasDelSale = False
        End If

        ' Load sale data into the form...
        Row = 0
        Balance = 0
        TotDue = 0
        txtBalDue.Visible = True
        lblBalDue.Visible = True
        OrdTotal = Format(Holding.Sale - Holding.Deposit, "###,###.00")
        TotDue = TotDue + OrdTotal
        GetMarginRecords(Holding.LeaseNo)
        UGridIO1.Refresh()
        txtBalDue.Text = CurrencyFormat(TotDue)
        SaleFound = True
        EnableControls(True)
        Cancel = True
    End Sub

    Private Sub MailCheckRef_SaleNotFound() Handles MailCheckRef.SaleNotFound
        ' Do nothing.. MailCheck will continue trying.
        SaleFound = False
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        'Next
        ' Clean up subforms
        'Unload BillOSale
        BillOSale.Close()

        EnableControls(True)
        OrdTotal = 0
        TotDue = 0
        UGridIO1.Clear()
        UGridIO1.Refresh()

        ' Find the customer/sale with MailCheck instead of EntryForm.
        '  If reportsmode("H") Or ordermode("Credit") Then
        MailCheckRef = MailCheck
        SaleFound = False
        If ReportsMode("H") Then
            MailCheckRef.optTelephone.Checked = True
        Else
            MailCheckRef.optSaleNo.Checked = True
            MailCheckSaleNoChecked = True
        End If
        'MailCheckRef.Show vbModal, Me
        MailCheckRef.ShowDialog(Me)
        'Unload MailCheckRef
        MailCheckRef.Close()
        MailCheckRef = Nothing
        MailCheckSaleNoChecked = False
        'If Not SaleFound Then cmdMenu.Value = True
        If Not SaleFound Then cmdMenu.PerformClick()
    End Sub

    Private Sub cmdNext2_Click(sender As Object, e As EventArgs) Handles cmdNext2.Click
        'Unload Me
        Me.Close()
        'Load Me
        CustomerAdjustment()
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        PrintReport()
        Printer.EndDoc()
    End Sub

    Private Sub PrintReport()
        OrdTotal = 0
        TotDue = 0

        Headings()

        '********************
        ' read out grid and print
        '********************
        Dim X As Integer

        If ReportsMode("H") Then
            For X = CurrentLine To UGridIO1.MaxRows - 1
                Dim C As Integer
                For C = 0 To UGridIO1.MaxCols - 1
                    If Counter = 60 Then
                        Counter = 0
                        Printer.NewPage()
                        Headings()
                    End If

                    Margin.SaleNo = UGridIO1.GetValue(X, 0)
                    Margin.Style = UGridIO1.GetValue(X, 1)
                    Margin.Vendor = UGridIO1.GetValue(X, 2)
                    Margin.Location = GetPrice(UGridIO1.GetValue(X, 3))
                    Margin.Status = UGridIO1.GetValue(X, 4)
                    Margin.Quantity = GetPrice(UGridIO1.GetValue(X, 5))
                    Margin.Desc = UGridIO1.GetValue(X, 6)
                    Margin.SellPrice = GetPrice(UGridIO1.GetValue(X, 7))
                Next

                ' BFH20090811 - the alignment is wrong... this would fix it if it was finished
                '            PrintToTab Printer, Margin.SaleNo, 0
                '            PrintToTab Printer, Margin.Style, 12
                '            PrintToTab Printer, Margin.Vendor, 32
                '            PrintToTab Printer, Margin.Status, 50
                '            PrintToTab Printer, Margin.Quantity, 58
                '            Printer.FontSize = 8
                '            PrintToTab Printer, Margin.Desc, 65
                '            Printer.FontSize = 10
                '            PrintToTab Printer, AlignString(CurrencyFormat(Margin.SellPrice), 13, vbAlignRight, False), 104


                Printer.Print(Margin.SaleNo, TAB(12), Margin.Style, TAB(32), Microsoft.VisualBasic.Left(Margin.Vendor, 12), TAB(50), Margin.Status, TAB(58), Margin.Quantity, TAB(65))
                Printer.FontSize = 8
                Printer.Print(Microsoft.VisualBasic.Left(Margin.Desc, 34))
                Printer.FontSize = 10
                Printer.Print(TAB(104), AlignString(CurrencyFormat(Margin.SellPrice), 13, VBRUN.AlignConstants.vbAlignRight, False))
                Counter = Counter + 1

                If Trim(Margin.SaleNo) = "" And Microsoft.VisualBasic.Left(UCase(Margin.Style), 10) <> "SALE DATE:" Then
                    OrdTotal = GetPrice(UGridIO1.GetValue(X, 8))
                    Printer.Print(TAB(95), "Order Total:         ", AlignString(CurrencyFormat(OrdTotal), 13, VBRUN.AlignConstants.vbAlignRight, False)) ' Spc(NoOfSpaces2); PstrFieldText
                    TotDue = TotDue + Format(OrdTotal, "###,###.00")

                    Counter = Counter + 1
                    OrdTotal = 0
                End If
                If Trim(UGridIO1.GetValue(X, 0)) = "" And Trim(UGridIO1.GetValue(X + 1, 0)) = "" And Trim(UGridIO1.GetValue(X + 2, 0)) = "" And Trim(UGridIO1.GetValue(X + 3, 0)) = "" Then Exit For
            Next
            Printer.Print()
            Printer.Print(TAB(96), "Total Due:          ", AlignString(CurrencyFormat(TotDue), 13, VBRUN.AlignConstants.vbAlignRight, False)) 'Spc(NoOfSpaces2); PstrFieldText
            Counter = Counter + 2
        ElseIf ReportsMode("I") Then
            ItemHistoryHeading()
            Counter = 0
            '********************
            ' read out grid and print
            '********************

            For X = CurrentLine To UGridIO1.MaxRows - 1
                Margin.Style = UGridIO1.GetValue(X, 0)
                Margin.Vendor = UGridIO1.GetValue(X, 1)
                Margin.Status = UGridIO1.GetValue(X, 2)
                Margin.Quantity = Val(UGridIO1.GetValue(X, 3))
                Margin.Desc = UGridIO1.GetValue(X, 4)
                Margin.SellPrice = GetPrice(UGridIO1.GetValue(X, 5))
                Printer.Print(Margin.Style, TAB(24), Margin.Vendor, TAB(48), Margin.Status, TAB(57), Margin.Quantity, TAB(64), Microsoft.VisualBasic.Left(Margin.Desc, 35), TAB(113), AlignString(CurrencyFormat(Margin.SellPrice), 13, VBRUN.AlignConstants.vbAlignRight, False)) 'Spc(NoOfSpaces); strFieldText
                Counter = Counter + 1

                If Counter = 60 Then
                    Printer.NewPage()
                    Counter = 0
                    ItemHistoryHeading()
                End If
                If Trim(UGridIO1.GetValue(X, 0)) = "" And Trim(UGridIO1.GetValue(X + 1, 0)) = "" Then Exit For
            Next
        End If
    End Sub

    Private Sub ItemHistoryHeading()
        Printer.FontName = "Arial"
        Printer.FontSize = 18
        Printer.CurrentY = 200
        Printer.CurrentX = 0

        Printer.Print(TAB(25), "Items Sold Since: ", Today) ' InvReports.ReportDate

        Printer.FontSize = 8
        Printer.CurrentY = 300
        Printer.CurrentX = 0
        Printer.Print("Date: ", DateFormat(Now))
        Printer.Print("Time: ", Format(Now, "h:mm:ss am/pm"))

        Printer.CurrentX = 10500
        Printer.CurrentY = 300
        Printer.Print("  Page:", Printer.Page)

        Printer.CurrentY = 600
        Printer.CurrentX = 0
        PrintCentered(StoreSettings.Name & "  " & "  " & StoreSettings.Address & "  " & "  " & StoreSettings.City)

        Printer.CurrentY = 800
        Printer.FontSize = 10

        Printer.FontBold = True
        Printer.Print(" Style Number            Manufacturer              Status Quan        Description", TAB(110), "Sell Price")
        Printer.FontBold = False
    End Sub

    Private Sub Headings()
        If ReportsMode("H") Then
            'Customer history
            Printer.FontName = "Arial"
            Printer.FontSize = 18
            Printer.CurrentY = 200
            Printer.CurrentX = 0

            Printer.Print(TAB(15), "Customer Consolidated History:  ")

            Printer.FontSize = 10
            Printer.CurrentY = 300

            Printer.Print(Mail.Last & "  " & Mail.Tele)
            Printer.FontSize = 8

            Printer.CurrentY = 300
            Printer.CurrentX = 0
            Printer.Print("Date: ", DateFormat(Now))
            Printer.Print("Time: ", Format(Now, "h:mm:ss am/pm"))

            Printer.CurrentX = 10500
            Printer.CurrentY = 300
            Printer.Print("Page:", Printer.Page)

            Printer.CurrentY = 600
            Printer.CurrentX = 0
            PrintCentered(StoreSettings.Name & "    " & StoreSettings.Address & "    " & StoreSettings.City)
            Printer.CurrentY = 800
            Printer.FontSize = 10

            Printer.FontSize = 12
            Printer.CurrentY = 800
            Printer.CurrentX = 0
            Printer.FontBold = True
            Printer.Print("Sale No   Style No       Manufacturer        Stat  Quan  Description                                       Price  Bal Due")
            Printer.FontBold = False

            Printer.CurrentX = 0
            Printer.FontSize = 10
        End If
    End Sub

    Private Sub cmdMenu_Click(sender As Object, e As EventArgs) Handles cmdMenu.Click
        ' quit
        'Unload BillOSale
        BillOSale.Close()
        'Unload Me
        FromCmdMenu = True
        Me.Close()

        If ReportsMode("I") Then
            InvReports.Show()
        ElseIf ReportsMode("H") Or OrderMode("Credit") Then
            modProgramState.Order = ""
            modProgramState.Reports = ""
            MainMenu.Show()
        End If
    End Sub

    Private Sub cmdMenu2_Click(sender As Object, e As EventArgs) Handles cmdMenu2.Click
        On Error Resume Next
        If SelectPrinter.SmallTags Then ' small tag was printed
            Printer.EndDoc()
            SelectPrinter.SmallTags = False
        End If
        'cmdMenu.Value = True
        'cmdMenu.PerformClick()
        cmdMenu_Click(cmdMenu, New EventArgs)
    End Sub

    Private Sub cmdReturn_Click(sender As Object, e As EventArgs) Handles cmdReturn.Click
        Dim I As Integer
        On Error GoTo ErrHand
        ' Copies from Grid 1 to Grid 2

        If FirstTime = True Then
            Row = 0                      '*** should be 0 but doesnt work
            FirstTime = False
        End If

        If Row > 9 Then
            MessageBox.Show("You're adjusting too many items.", "Too many adjustments", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub  ' Max number of rows
        End If

        Lines = UGridIO1.Row
        If Microsoft.VisualBasic.Left(UGridIO1.GetValue(Lines, 4), 1) = "x" Then
            MessageBox.Show("This item has already been returned.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If Microsoft.VisualBasic.Left(UGridIO1.GetValue(Lines, 4), 2) = "VD" Then
            MessageBox.Show("This item has already been voided.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If Val(UGridIO1.GetValue(Lines, 5)) < 0 Then
            MessageBox.Show("You cannot return items with a negative quantity already.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        For I = 0 To UGridIO2.MaxRows - 1
            If UGridIO2.GetValue(I, 10) = UGridIO1.GetValue(Lines, 9) Then
                If UGridIO2.GetValue(I, 10) = "" Then Exit Sub
                MessageBox.Show("This item is already being returned.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Next

        Select Case UCase(Trim(UGridIO1.GetValue(Lines, 1)))
            Case "SUB", "--- ADJ ---" ', "TAX1", "TAX2"
                MessageBox.Show("You can't return this item.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            Case "PAYMENT"
                MessageBox.Show("You can't return payments this way.  Add a negative amount on the Payment On Account screen instead.")
                Exit Sub
            Case "TAX1", "TAX2"
                If Row <> 0 Then
                    MessageBox.Show("If you remove tax from a sale, it must be the only item adjusted.", "Cannot Remove Tax With Other Items", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
        End Select

        'With UGridIO2
        UGridIO2.SetValue(Row, 0, UGridIO1.GetValue(Lines, 0))       ' Sale
        UGridIO2.SetValue(Row, 1, UGridIO1.GetValue(Lines, 1))       ' Style
        UGridIO2.SetValue(Row, 2, UGridIO1.GetValue(Lines, 2))       ' Vendor
        UGridIO2.SetValue(Row, 3, UGridIO1.GetValue(Lines, 3))       ' Loc
        UGridIO2.SetValue(Row, 4, UGridIO1.GetValue(Lines, 4))       ' Status
        UGridIO2.SetValue(Row, 5, -Val(UGridIO1.GetValue(Lines, 5))) ' Quantity
        UGridIO2.SetValue(Row, 6, UGridIO1.GetValue(Lines, 6))       ' Description
        UGridIO2.SetValue(Row, 7, -GetPrice(UGridIO1.GetValue(Lines, 10))) ' Price
        UGridIO2.SetValue(Row, 8, "")  ' Difference, subject to Recalculate.
        If Val(UGridIO1.GetValue(Lines, 5)) = 0 Then
            UGridIO2.SetValue(Row, 9, UGridIO1.GetValue(Lines, 7))     ' Per-unit price
        Else
            UGridIO2.SetValue(Row, 9, UGridIO1.GetValue(Lines, 7) / UGridIO1.GetValue(Lines, 5))  ' Per-unit price
        End If
        UGridIO2.SetValue(Row, 10, UGridIO1.GetValue(Lines, 9))      ' Marginline
        UGridIO2.SetValue(Row, 11, UGridIO1.GetValue(Lines, 11))     ' Commission

        balRow = Row

        Counter = Counter + 1
        Row = Row + 1

        If (UGridIO2.Row > 4) Then UGridIO2.MoveRowDown(Val(UGridIO2.Row - 4))
        UGridIO2.Row = UGridIO2.Row + 1
        UGridIO2.Refresh(True)

        Recalculate()
        Exit Sub
ErrHand:
        MessageBox.Show("An error occurred.  If this persists, please contact " & AdminContactString(Format:=1, Phone:=False) & "." & vbCrLf & "Err: " & Err.Number & " - " & Err.Description & vbCrLf & "Ref: OnScreenReport::cmdReturn_Click() - 1", "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        'NOTE: cmdAdd is not button. It is checkbox with appearance set as button.
        On Error GoTo ErrHand
        'add items
        If FirstTime Then
            Row = 0                      '*** should be 0 but doesnt work
            FirstTime = False
        End If

        If Row >= MaxAdjustments Then
            MessageBox.Show("You're adding too many adjustments!", "Too many adjustments", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Margin.Quantity = 0

        '  txtLocation.Visible = True
        '  lblLocation.Visible = True
        '  InvCkStyle.Show vbModal
        '  If Not InvCkStyle.Canceled Then AddInventory InvCkStyle.Rn, InvCkStyle.StyleCkIt
        '  Unload InvCkStyle
        'frmAdjustAdd.Show 1
        frmAdjustAdd.ShowDialog(Me)
        Exit Sub
ErrHand:
        MessageBox.Show("An error occurred.  If this persists, please contact " & AdminContactCompany & "." & vbCrLf & "Err: " & Err.Number & " - " & Err.Description & vbCrLf & "Ref: OnScreenReport::cmdAdd_Click() - 1", "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub cmdAdjustTax_Click(sender As Object, e As EventArgs) Handles cmdAdjustTax.Click
        'Load frmAdjustTax
        frmAdjustTax.LoadSale(SaleNo)
        'frmAdjustTax.Show vbModal
        frmAdjustTax.ShowDialog()
        'Unload frmAdjustTax
        frmAdjustTax.Close()
        Exit Sub

        Dim T1 As Boolean, T2 As Boolean, TAmt As Decimal
        Dim R As VBA.VbMsgBoxResult, Z As Integer

        T1 = SaleHasTax1()
        T2 = SaleHasTax2()

        If T1 Then
            TAmt = SaleTax1Amount()
            R = MessageBox.Show("Remove default tax (TAX1 = " & StoreSettings.SalesTax & ")?", "Adjust Tax", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If R = vbCancel Then Exit Sub
            If R = vbYes Then
                If TAmt = 0 Then
                    MessageBox.Show("The total for this tax type is $0.00!", "No TAX1", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Else
                    AddInventory(0, "TAX1", "", 0, "", "", -TAmt)
                End If
                Exit Sub
            End If
        End If

        If T2 Then
            Z = SaleTax2Zone()
            TAmt = SaleTax2Amount()
            R = MessageBox.Show("Remove TAX2 (" & StoreSettings.SalesTax & ")?", "Adjust Tax", MessageBoxButtons.YesNoCancel)
            If R = vbCancel Then Exit Sub
            If R = vbYes Then
                If TAmt = 0 Then
                    MessageBox.Show("The total for this tax type is $0.00!", "No TAX1", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Else
                    AddInventory(0, "TAX1", "", 0, "", "", -TAmt)
                End If
                Exit Sub
            End If
        End If

        If SaleTax1Amount() = 0 Then
            R = MsgBox("Add default tax (Tax = " & StoreSettings.SalesTax & ")?", vbQuestion + vbYesNoCancel)
            If R = vbCancel Then Exit Sub
            If R = vbYes Then
                TAmt = 55 '###!!!
                AddInventory(0, "TAX1", "", 0, "", "", TAmt)
                Exit Sub
            End If
        End If

        R = MessageBox.Show("Add variable tax (Tax = " & StoreSettings.SalesTax & ")?", "", MessageBoxButtons.YesNoCancel)
    End Sub

    Public Sub AddInventory(ByVal RN As Integer, ByVal StyleCkIt As String, Optional ByVal Status As String = "#", Optional ByVal Quan As Double = 1, Optional ByVal Desc As String = "#", Optional ByVal Vend As String = "#", Optional ByVal Price As Decimal = -1, Optional ByVal Location As Integer = 0)
        Dim InvData As CInvRec
        Dim Lt As Decimal

        Counter = Counter + 1
        InvData = New CInvRec
        StyleCkIt = Trim(StyleCkIt)
        If Location <= 0 Then Location = StoresSld ' BFH20060815

        If Microsoft.VisualBasic.Left(StyleCkIt, 4) = KIT_PFX Then
            Dim TRec As CInvRec
            Dim Tot As Decimal, PACK As Decimal, Factor As Double, Sum As Object
            Dim RS As ADODB.Recordset, I As Integer, Kst As String, Krn As Integer, Kqu As Double, Kpr As Decimal

            'Debug.Print "x": DoEvents
            '    Status = SelectStatus()
            'Debug.Print "y": DoEvents
            '    If Status = "" Then Exit Sub

            RS = GetRecordsetBySQL("SELECT * FROM [InvKit] WHERE [KitStyleNo]='" & StyleCkIt & "'", , GetDatabaseAtLocation(1))
            If Not RS.EOF Then
                Tot = IfNullThenZeroCurrency(RS("OnSale").Value)
                PACK = IfNullThenZeroCurrency(RS("PackPrice").Value)
                Price = PACK

                Lt = IfNullThenZeroCurrency(RS("Landed").Value)

                Factor = PACK / Tot
                For I = 1 To Setup_MaxKitItems
                    Kst = IfNullThenNilString(RS("Item" & I).Value)
                    Krn = IfNullThenZero(RS("Item" & I & "Rec").Value)
                    If IsFormLoaded("frmKitLevels") Then
                        Kqu = frmKitLevels.ItemQuantityByStyle(IfNullThenNilString(RS("Item" & I).Value))
                    Else
                        Kqu = IfNullThenZeroDouble(RS("Quan" & I).Value)
                    End If

                    If Kst <> "" Then
                        TRec = New CInvRec
                        TRec.Load(Kst, "Style")
                        Kpr = Math.Round(TRec.OnSale * Factor, 2)
                        DisposeDA(TRec)

                        If Sum + Kpr > PACK Then Kpr = PACK - Sum ' pennies, rounding, and perfection

                        Sum = Sum + Kpr

                        If Not StoreSettings.bShowPackageItemPrices Then
                            AddInventory(Krn, Kst, Status, Kqu, , , 0, Location)
                        Else
                            Dim C As CInvRec, KP As Decimal
                            C = New CInvRec
                            C.Load(Kst, "Style")
                            If Lt <> 0 Then
                                KP = IfNullThenZeroCurrency(RS("PackPrice").Value) * C.Landed / Lt
                            Else
                                KP = 0
                            End If
                            AddInventory(Krn, Kst, Status, Kqu, , , KP, Location)
                            DisposeDA(C)
                        End If
                    End If
                Next
            End If
        End If

        If RN <> 0 Then   'item in inventory else S/S
            If Not InvData.Load(RN, "#Rn") Then
                ' Couldn't load Style.
            End If
        End If

        'Unload frmKitLevels
        frmKitLevels.Close()

        If Not Microsoft.VisualBasic.Left(StyleCkIt, 4) = KIT_PFX Or Not StoreSettings.bShowPackageItemPrices Then
            UGridIO2.SetValueDisplay(Row, 0, Trim(SaleNo))
            UGridIO2.SetValueDisplay(Row, 1, IIf(InvData.Style = "", StyleCkIt, InvData.Style))
            UGridIO2.SetValueDisplay(Row, 2, IIf(Vend = "#", InvData.Vendor, Vend))
            UGridIO2.SetValueDisplay(Row, 3, CStr(Location)) 'StoresSld '"1" ' Loc

            If Microsoft.VisualBasic.Left(StyleCkIt, 4) = KIT_PFX Then Status = ""

            If Status = "#" Then
                If Val(RN) <> 0 Then   'item in inventory else S/S
                    If IsUFO() Or IsAuthenTeak() Then
                        UGridIO2.SetValueDisplay(Row, 4, "LAW")
                    Else
                        UGridIO2.SetValueDisplay(Row, 4, "ST")
                    End If
                ElseIf Microsoft.VisualBasic.Left(StyleCkIt, 4) = KIT_PFX Then
                    UGridIO2.SetValueDisplay(Row, 1, StyleCkIt)
                    UGridIO2.SetValueDisplay(Row, 4, "ST")
                Else
                    UGridIO2.SetValueDisplay(Row, 1, StyleCkIt)
                    UGridIO2.SetValueDisplay(Row, 4, "SS")
                    If IsIn(StyleCkIt, "DEL", "LAB", "NOTES", "STAIN", "TAX1") Then UGridIO2.SetValueDisplay(Row, 4, " ")
                End If
            Else
                UGridIO2.SetValueDisplay(Row, 4, Status)
                If Status = "SS" Then UGridIO2.SetValueDisplay(Row, 1, StyleCkIt)
            End If

            UGridIO2.SetValueDisplay(Row, 5, Str(Quan))
            UGridIO2.SetValueDisplay(Row, 6, IIf(Desc = "#", InvData.Desc, Desc))

            Select Case StyleCkIt
                Case "DEL" : UGridIO2.SetValueDisplay(Row, 6, "DELIVERY CHARGE")
                Case "LAB" : UGridIO2.SetValueDisplay(Row, 6, "LABOR CHARGE")
                Case "STAIN" : UGridIO2.SetValueDisplay(Row, 6, IIf(IsBFMyer, "SAFEWARE PROTECTION PLAN", "STAIN PROTECTION")) : UGridIO2.SetValueDisplay(Row, 5, "")
            End Select

            If Price <> -1 Then
                UGridIO2.SetValueDisplay(Row, 7, Format(Price, "###,##0.00"))
            Else
                UGridIO2.SetValueDisplay(Row, 7, Format(InvData.OnSale * Quan, "###,##0.00"))
            End If

            balRow = Row                        ' This is handled by Recalculate.

            InvRn(Row) = RN
            Cost(Row) = GetItemCost(InvData.Style, Location, False, Quan) ' BFH20061010
            Freight(Row) = InvData.Landed - InvData.Cost
            Vends(Row) = InvData.VendorNo
            Depts(Row) = InvData.DeptNo
            DetailRec(Row) = Margin.Detail      'doesn't work
            Quantity(Row) = 1 ' shouldn't this be  Val(Quan), not 1 ???????? -- bfh20090409

            UGridIO2.SetValue(Row, 9, InvData.OnSale)    ' Single-item price

            Row = Row + 1

            If (UGridIO2.Row > 4) Then UGridIO2.MoveRowDown(Val(UGridIO2.Row - 4))
            UGridIO2.Row = UGridIO2.Row + 1

        End If
        UGridIO2.Refresh()

        DisposeDA(InvData)
        Recalculate()
    End Sub

    Private Sub UGridIO2_BeforeColUpdate(ByVal ColIndex As Integer, ByRef OldValue As Object, ByRef Cancel As Integer) Handles UGridIO2.BeforeColUpdate
        Dim newValue As String, Row As Integer

        newValue = UGridIO2.Text
        Row = UGridIO2.Row

        If UGridIO2.GetValue(UGridIO2.Row, 1) = "NOTES" Then Exit Sub
        If Val(OldValue) >= 0 And Val(newValue) < 0 Then

            ' jk 20070626 I took this out to return an item from a package w/no selling price
            ' MsgBox "Use Return to return an item.", vbExclamation, "Negative Add Prohibited"
            ' Cancel = True
        End If
        'BFH20070424
        '  If Val(OldValue) < 0 And Val(UGridIO2.Text) >= 0 Then
        '    MsgBox "Use Add to add items.", vbExclamation, "Positive Return Prohibited"
        '    Cancel = True
        '  End If
        If UGridIO2.GetValue(Row, 4) = "SO" And ColIndex = 5 Then
            MessageBox.Show("Cannot adjust SO lines.", "SO Adjustment Prohibited'", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Cancel = True
        End If

        If ColIndex = BillColumns.eQuant + 1 Then
            If Val(UGridIO2.GetValue(Row, 5)) < 0 Then
                Dim ML As Integer, I As Integer
                Dim Sty As String
                Dim C As Integer, V As String, P As Decimal
                Dim SaleAmt As Double
                Dim X2 As Integer
                ML = UGridIO2.GetValue(Row, 10)
                If ML = 0 Then Exit Sub

                For I = 0 To UGridIO1.LastRowUsed
                    If Val(UGridIO1.GetValue(I, 9)) = ML Then GoTo FoundML
                Next

                Exit Sub

FoundML:

                SaleAmt = Val(UGridIO1.GetValue(I, BillColumns.eQuant + 1))
                If -Val(OldValue) <> SaleAmt Then
                    MessageBox.Show("You have already adjusted this quatity during which an additional adjustment line for the amount not retruned was created." & vbCrLf & "To ensure the total value does not change, please restart the adjustment process for this sale if you need to change this again.")
                    Exit Sub
                End If

                If MessageBox.Show("You are attempting to do a partial return." & vbCrLf2 & "In order to properly account for inventory, WinCDS will do a full return on this line, and will place the unreturned items as an additional line item." & vbCrLf2 & "Would you like to proceed?", "Comfirm Partial Return", MessageBoxButtons.OKCancel) = DialogResult.Cancel Then
                    Cancel = True
                    Exit Sub
                End If

                Cancel = True
                X2 = UGridIO2.FirstEmptyRow
                Sty = UGridIO2.GetValue(Row, BillColumns.eStyle + 1)
                V = SaleAmt + Val(newValue)
                P = GetPrice(UGridIO2.GetValue(X2 - 1, BillColumns.ePrice + 2)) + GetPrice(UGridIO2.GetValue(X2, BillColumns.ePrice + 1))
                AddInventory(GetRNByStyle(Sty), Sty, , V, , , P)

                For C = 0 To UGridIO2.MaxCols - 1
                    If C = BillColumns.eQuant + 1 Then
                        V = SaleAmt + Val(newValue)
                    ElseIf C = BillColumns.ePrice + 1 Then
                        V = FormatCurrency(GetPrice(UGridIO1.GetValue(I, C)) / Val(UGridIO1.GetValue(I, 5)) * Val(UGridIO2.GetValue(X2, 5)))
                    ElseIf C = BillColumns.ePrice + 2 Then
                        V = CurrencyFormat(GetPrice(UGridIO2.GetValue(X2 - 1, BillColumns.ePrice + 2)) + GetPrice(UGridIO2.GetValue(X2, BillColumns.ePrice + 1)))
                    ElseIf C = 10 Then
                        V = ""
                    Else
                        V = UGridIO1.GetValue(I, C)
                    End If
                    UGridIO2.SetValue(X2, C, V)
                Next
                Counter = Counter + 1
                UGridIO2.Refresh()
                Recalculate()
            End If
        End If
    End Sub

    Private Sub UGridIO2_RowColChange(LastRow As Object, LastCol As Object, newRow As Object, newCol As Object, ByRef Cancel As Boolean) Handles UGridIO2.RowColChange
        PriorBal = 0 'used to tab off field

        If mLoading = True Then Exit Sub
        If Row = 0 Or FirstTime = True Then
            If newRow <> 0 Then UGridIO2.Row = 0
            If newCol <> 0 Then UGridIO2.Col = 0
            Exit Sub
        End If
        'If IsEmpty(LastRow) And IsEmpty(LastCol) And IsEmpty(newRow) And IsEmpty(newCol) Then Exit Sub
        If IsNothing(LastRow) And IsNothing(LastCol) And IsNothing(newRow) And IsNothing(newCol) Then Exit Sub
        If newRow = -1 Then Exit Sub
        If newRow >= Row Then Cancel = True : Exit Sub
        CurrentLine = UGridIO2.Row
        If Not IsNothing(newRow) Then
            If LastRow = Str(newRow) And LastCol = newCol Then
                Exit Sub
            Else
                On Error Resume Next
                Select Case newCol
' BFH20101110 - Commented out... Jerry's email was
' "We still have this old adjustment in Customer Adjustment.  Please comment out."
'        Case 4
'          cmbGrid2.Tag = Format(newRow, "000") & Format(newCol, "000")
'          cmbGrid2.Visible = False
'          UGridIO2.AdjustControlToCell cmbGrid2, newRow, newCol, UGridIO2.Left, UGridIO2.Top
'          cmbGrid2.Clear
'          cmbGrid2.AddItem "ST"
'          cmbGrid2.AddItem "SO"
'          cmbGrid2.AddItem "SS"
'          cmbGrid2.AddItem "FND"
'          cmbGrid2.AddItem "PO"
'          cmbGrid2.AddItem "SS"
'          cmbGrid2.AddItem "SSLAW"
'          cmbGrid2.AddItem "DELTW"
'          cmbGrid2.Text = UGridIO2.GetValue(newRow, newCol)
'          cmbGrid2.Visible = True
'          cmbGrid2.SetFocus
                    Case 2
                        cmbGrid2.Tag = Format(newRow, "000") & Format(newCol, "000")
                        cmbGrid2.Visible = False
                        UGridIO2.AdjustControlToCell(cmbGrid2, newRow, newCol, UGridIO2.Left, UGridIO2.Top)
                        LoadMfgNamesIntoComboBox(cmbGrid2, "", True, True)
                        cmbGrid2.Text = UGridIO2.GetValue(newRow, newCol)
                        cmbGrid2.Visible = True
                        cmbGrid2.Select()
                        'cmbGrid2.Location = New Point(25, 350)
                        cmbGrid2.BringToFront()
                    Case Else
                        cmbGrid2.Visible = False
                End Select
            End If
        End If

    End Sub

    Private Sub cmbGrid2_Leave(sender As Object, e As EventArgs) Handles cmbGrid2.Leave
        'Lost focus
        On Error Resume Next
        Dim R As Integer, C As Integer
        cmbGrid2.Visible = False
        UGridIO2.SetValue(Val(Microsoft.VisualBasic.Left(cmbGrid2.Tag, 3)), Val(Microsoft.VisualBasic.Right(cmbGrid2.Tag, 3)), cmbGrid2.Text)
        '  UGridIO2.Col = UGridIO2.Col + 1
        UGridIO2.Select()
        R = UGridIO2.Row
        C = UGridIO2.Col
        UGridIO2.Refresh()
        UGridIO2.Row = R
        UGridIO2.Col = C
    End Sub

    Private Sub RowCheck()
        If Row >= UGridIO1.MaxRows Then UGridIO1.MaxRows = UGridIO1.MaxRows + 100
    End Sub

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        'apply
        Dim T As Boolean, H As Boolean
        Dim OldStatus As String

        '  If IsIn(g_Holding.Status, "F", "S") Then
        '    If GetPrice(txtBalDue) > 0 Then
        '      If MsgBox("This sale has an open Store Finance account." & vbCrLf & "Adding an additional balance due of " & txtBalDue & " will create an Add-On to the original contract." & vbCrLf2 & "Proceed with altering the terms of the contract?", vbExclamation + vbOKCancel, "Confirm Change Contract Terms") = vbCancel Then Exit Sub
        '    ElseIf GetPrice(txtBalDue) < 0 Then
        '      If MsgBox("This sale has an open Store Finance account." & vbCrLf & "Removing an balance due of " & txtBalDue & " affect a payment on this account." & vbCrLf2 & "Proceed with altering the terms of the contract?", vbExclamation + vbOKCancel, "Confirm Change Contract Terms") = vbCancel Then Exit Sub
        '    End If
        '  End If
        '
        On Error GoTo ErrHand
        EnableControls(False)

        UGridIO2.Update()

        H = g_Holding.Load(Trim(SaleNo)) ' bfh20051107 - added load here b/c Audit uses Holding w/o checks..

        WriteOutRemovedAllUndelivered = False
        T = OrderHasUndeliveredItems(Margin.SaleNo)
        If Row > 0 Then AddAdjustment(g_Holding.Status = "D")    'Adj heading line
        WriteOut()       ' add lines to margin
        UGridIO2.GetDBGrid.Refresh()

        If T Then
            If OrderHasUndeliveredItems(Margin.SaleNo) Then WriteOutRemovedAllUndelivered = True
        End If

        Dim K As Integer, KB As Decimal
        'For K = txtDiffTax.LBound To txtDiffTax.UBound
        '    KB = KB + GetPrice(txtDiffTax(K))
        'Next
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "txtDiffTax" Then
                KB = KB + GetPrice(C.Text)
            End If
        Next

        'add to holding
        If H Then
            OldStatus = g_Holding.Status
            If g_Holding.Status = "B" And WriteOutRemovedAllUndelivered Then
                g_Holding.Status = "D"
                MessageBox.Show("Switching order status from 'backorder' to 'delivered'.")
            ElseIf WasDelSale And g_Holding.Status = "D" And WriteOutAddedItems Then
                g_Holding.Status = "B"
                MessageBox.Show("Switching order status from 'delivered' to 'backorder'.")
                If GetPrice(Balance) > 0 Then ' BFH20060522 - add BO cash line if they owe more!
                    AddNewCashJournalRecord(Account_Backorders, GetPrice(Balance) + GetPrice(TotalTax), g_Holding.LeaseNo, "", Today)
                End If
            End If
            g_Holding.Sale = Format(g_Holding.Sale + Balance + IIf(TaxBackedOut, 0, KB), "###,###.00")
            g_Holding.NonTaxable = GetPrice(g_Holding.NonTaxable) + NonTaxable ' bfh20060216 - added holding.nontaxable to this calculation... it only handled the new part of the info
            '    If KB = 0 Then
            '      g_Holding.NonTaxable = GetPrice(g_Holding.NonTaxable) + Balance  ' bfh20140701 - non taxable sales added to with no tax must increase the non-taxable amount
            '    End If
            g_Holding.Save()
        Else
            ' Couldn't find the LeaseNo.
        End If

        If Balance <> 0 Or KB <> 0 Then  ' bfh20060119 - moved below H block b/c it may change holding status
            AddTaxLine()
        End If

        ' BFH20060525 the refund was put in with balance <> 0
        ' if there was an overpayment on sale creation,
        ' and then you return and add the same amount,
        ' no audit line was created, but the money was refunded.  This put the
        ' DA report off..

        'Credit Balance
        If txtBalDue.Text < 0 Then
            CustAdjRefund.CreditAmt = -txtBalDue.Text
            CustAdjRefund.Margin = Margin
            CustAdjRefund.SaleNo = Margin.SaleNo
            'CustAdjRefund.Show vbModal, Me
            CustAdjRefund.ShowDialog(Me)
        End If

        Audit()

        If IsIn(OldStatus, "B", "F", "C") Then
            AddNewCashJournalRecord(IIf(OldStatus = "C", "13200", "11200"), Balance + TotalTax() - IIf(GetPrice(txtBalDue.Text) < 0, GetPrice(txtBalDue.Text), 0), SaleNo, LastName, Today)
        End If

        UGridIO2.GetDBGrid.Refresh()

        SalePackageUpdate(SaleNo:=SaleNo, AllowCache:=False)

        If IsIn(g_Holding.Status, "S", "F") Then
            If MessageBox.Show("This Sale is Store Financed, do you want to do an Add on Sale?", SALE_STATUS_OPENFINANCE, MessageBoxButtons.OKCancel) = DialogResult.OK Then
                ARPaySetUp.Show()
            Else
                MessageBox.Show("Add On was cancelled." & vbCrLf & "Adjustments were made on this Installment Sale." & vbCrLf & "To make it balance, make a payment to the sale or reverse the adjustment back to its original version.")
            End If
        End If

        EnableControls(True, True) 'Uncomment this to allow continuation after Apply for accounts with remaining credit ''Val(txtBalDue.Text) >= 0

        Exit Sub
ErrHand:
        MessageBox.Show("An error occurred.  If this persists, please contact " & AdminContactCompany & "." & vbCrLf & "Err: " & Err.Number & " - " & Err.Description & vbCrLf & "Ref: OnScreenReport::cmdApply_Click() - 1", "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '  Resume Next
    End Sub

    Private Sub Audit()
        '  AddNewAuditRecord SaleNo, "Adj.  " + Trim(LastName),date,balance,totaltax,balance+totaltax,
        Dim NewAudit As SalesJournalNew
        NewAudit.SaleNo = Trim(SaleNo)
        NewAudit.Name1 = "Adj.  " + Trim(LastName)
        NewAudit.TransDate = Trim(DateFormat(Now))
        NewAudit.Written = Trim(CurrencyFormat(Balance))
        NewAudit.TaxCharged1 = Trim(TotalTax)
        NewAudit.Cashier = GetCashierName

        ' we have to handle adjusting delivered sales differently b/c the
        ' delivery did some of the work for us, and also created some extra tasks
        If Not WasDelSale And g_Holding.Status <> "D" Then
            NewAudit.ArCashSls = Trim(CurrencyFormat(Balance + TotalTax()))
            NewAudit.Cashier = GetCashierName

            If txtBalDue.Text < 0 Then
                If Not LeaveCreditBalance Then
                    NewAudit.Control = Math.Abs(CByte(txtBalDue.Text)) ' This assumes the refund is happening immediately.  Since "Leave Credit Balance" is an option, that is not the case.
                Else
                    NewAudit.Control = 0 'GetPrice(txtBalDue)
                    '        Dim cHold As cHolding
                    '        Set cHold = New cHolding
                    '        cHold.Load Trim(SaleNo), "LeaseNo"
                    '        cHold.Status = "B"
                    '        cHold.Save
                End If
            Else
                NewAudit.Control = 0
            End If
            NewAudit.UndSls = Trim(CurrencyFormat(Balance + TotalTax()))
            NewAudit.DelSls = 0
            NewAudit.TaxRec1 = 0
        Else
            NewAudit.Written = Trim(CurrencyFormat(Balance))
            NewAudit.TaxCharged1 = TotalTax()
            ' bfh20051014
            ' for already delivered sales being adjusted, these values will already have been
            ' handled by the DS line in the audit table.  So for the Adj. line, we have
            ' to do things a little differently..
            ' These values are already cancelled out..
            NewAudit.ArCashSls = 0
            If LeaveCreditBalance Then
                NewAudit.Control = GetPrice(txtBalDue.Text)
                g_Holding.Status = "B"
                g_Holding.Save()
            Else
                NewAudit.Control = 0
            End If
            NewAudit.UndSls = 0
            ' Delivered Sales now must be decreased because the delivery recorded them
            NewAudit.DelSls = Trim(CurrencyFormat(Balance))
            NewAudit.TaxRec1 = Trim(TotalTax)
        End If

        If TaxLoc = 0 Then
            NewAudit.TaxCode = 1
        ElseIf TaxLoc = 1 Then
            NewAudit.TaxCode = 1  ' this is here b/c taxcode=1 is TAX1 and doesn't get adjusted by -1 like the TAX2 sales
        ElseIf TaxLoc = -1 Then
            NewAudit.TaxCode = lblRate0.Tag  ' BFH20060216 - Added " - 1".  Didn't take into account the internal adjustment of +1 for TAX2 codes for differentiation
        Else
            NewAudit.TaxCode = TaxLoc - 1  ' BFH20060216 - Added " - 1".  Didn't take into account the internal adjustment of +1 for TAX2 codes for differentiation
        End If
        NewAudit.Salesman = Trim(Margin.Salesman)

        NewAudit.NonTaxable = NonTaxable ' BFH20060519 no nontaxable info changed

        SalesJournal_AddRecordNew(NewAudit)

        OrdTotal = 0
    End Sub

    Private Sub AddTaxLine()
        Dim K As Integer
        Dim Tk As Decimal
        'For K = txtDiffTax.LBound To txtDiffTax.UBound
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "txtDiffTax" Then
                'If GetPrice(txtDiffTax(K).Text) <> 0 Then
                If GetPrice(C.Text) <> 0 Then
                    Margin.SaleNo = SaleNo
                    Margin.Style = "SUB"
                    Margin.Vendor = ""
                    Margin.Status = ""
                    Margin.Quantity = 0
                    Margin.Desc = "                   Sub Total ="
                    'Margin.SellPrice = CurrencyFormat(OrdTotal + Balance + IIf(GetPrice(txtDiffTax(K)) > 0, Tk, 0))
                    Margin.SellPrice = CurrencyFormat(OrdTotal + Balance + IIf(GetPrice(C.Text) > 0, Tk, 0))
                    'Tk = Tk + GetPrice(txtDiffTax(K))
                    Tk = Tk + GetPrice(C.Text)
                    Margin.GM = 0
                    AddMarginLine()

                    If TaxBackedOut Then
                        Margin.Vendor = ""
                        Margin.Style = "NOTES"
                        Margin.Quantity = 0
                        Margin.Desc = "PRICE WITH TAX BACKED OUT: " & CurrencyFormat(Balance)
                        'Margin.SellPrice = -CurrencyFormat(GetPrice(txtDiffTax(K)))
                        Margin.SellPrice = -CurrencyFormat(GetPrice(C.Text))
                        AddMarginLine()
                    End If

                    'only uses these codes.
                    Margin.SaleNo = SaleNo
                    Margin.Style = IIf(TaxLoc = 1, "TAX1", "TAX2")
                    Margin.Vendor = ""
                    Margin.Status = ""
                    If TaxLoc = 1 Then
                        Margin.Quantity = TaxLoc
                        Margin.Desc = "SALES TAX DIFF.             ="
                    Else
                        K = Mid(C.Name, 11)
                        'Margin.Quantity = Val(lblRate(K).Tag)
                        For Each L As Control In Me.Controls
                            If L.Name = "lblRate" & K Then
                                Margin.Quantity = Val(L.Tag)
                                Exit For
                            End If
                        Next
                        Margin.Desc = "SALES TAX DIFF.             " & GetTax2String(Margin.Quantity, True) & " ="
                    End If
                    'Margin.SellPrice = CurrencyFormat(GetPrice(txtDiffTax(K)))
                    Margin.SellPrice = CurrencyFormat(GetPrice(C.Text))
                    AddMarginLine()
                End If
            End If
        Next
    End Sub

    Private Function TotalTax() As Decimal
        Dim K As Integer
        'For K = txtDiffTax.LBound To txtDiffTax.UBound
        '    TotalTax = TotalTax + GetPrice(txtDiffTax(K))
        'Next
        For Each C As Control In Me.Controls
            If Mid(C.Name, 1, 10) = "txtDiffTax" Then
                TotalTax = TotalTax + GetPrice(C.Text)
            End If
        Next
    End Function

    Private Sub AddAdjustment(Optional ByVal IsDelivered As Boolean = False)
        Margin.Name = LastName
        Margin.Index = Index
        Margin.SaleNo = SaleNo
        Margin.Style = "--- Adj ---"
        Margin.Vendor = "--Adjustments--"
        Margin.Status = ""
        Margin.Quantity = 0
        Margin.Desc = "--- Adjustments --- " & DateFormat(Now) & "  Sub Total ="
        Margin.SellPrice = OrdTotal
        Margin.Detail = 0
        Margin.Commission = ""
        Margin.SellDte = DateFormat(Today)

        If IsDelivered Then
            Margin.DDelDat = DateFormat(Today)
        End If
        AddMarginLine()
    End Sub

    Public Function SaleHasTax1() As Boolean
        Dim I As Integer
        For I = 1 To UGridIO1.LastRowUsed
            If Trim(UGridIO1.GetValue(I, 1)) = "TAX1" Then SaleHasTax1 = True : Exit Function
        Next
    End Function

    Public Function SaleHasTax2(Optional ByVal Zone As Integer = 0) As Boolean
        Dim I As Integer
        For I = 1 To UGridIO1.LastRowUsed
            If Trim(UGridIO1.GetValue(I, 1)) = "TAX2" Then
                If Zone = 0 Or (Zone = UGridIO1.GetValue(I, 5)) Then
                    SaleHasTax2 = True
                    Exit Function
                End If
            End If
        Next
    End Function

    Public Function SaleTax2Zone(Optional ByVal Zone As Integer = 0) As Integer
        Dim I As Integer
        For I = 1 To UGridIO1.LastRowUsed
            If Trim(UGridIO1.GetValue(I, 1)) = "TAX2" Then
                If Zone = 0 Then
                    SaleTax2Zone = Val(UGridIO1.GetValue(I, 5))
                    Exit Function
                ElseIf Zone = UGridIO1.GetValue(I, 5) Then
                    SaleTax2Zone = Zone
                    Exit Function
                End If
            End If
        Next
    End Function

    Public Function DeveloperEx() As String
        DeveloperEx = "DCK Difference Tax" & vbCrLf & " Recalculate"
    End Function

    Private Sub UGridIO2_Change() Handles UGridIO2.Change
        Static Running As Boolean
        If Running Then Exit Sub
        Dim tRow As Integer
        tRow = UGridIO2.Row

        Select Case UGridIO2.Col
            Case BillColumns.eQuant + 1
                Quantity(tRow) = UGridIO2.GetDBGrid.Text
            Case BillColumns.eManufacturer + 1
                ChangeGrid2(tRow, 2, UCase(UGridIO2.GetDBGrid.Text))
            Case BillColumns.eStatus + 1
                ChangeGrid2(tRow, 4, UCase(UGridIO2.GetDBGrid.Text))
            Case BillColumns.eDescription + 1
                ChangeGrid2(tRow, 6, UCase(UGridIO2.GetDBGrid.Text))
            Case BillColumns.ePrice + 1
        End Select
    End Sub

    Private Sub txtDiffTax0_DoubleClick(sender As Object, e As EventArgs) Handles txtDiffTax0.DoubleClick
        Recalculate()
    End Sub

    Public Function OrderHasUndeliveredItems(ByRef SaleNo As String) As Boolean
        Dim tMargin As New CGrossMargin
        tMargin.DataAccess.DataBase = Margin.DataAccess.DataBase
        tMargin.DataAccess.Records_OpenSQL("SELECT * FROM GrossMargin WHERE SaleNo='" & SaleNo & "'")
        Do While tMargin.DataAccess.Records_Available
            tMargin.cDataAccess_GetRecordSet(tMargin.DataAccess.RS)
            If IsItem(tMargin.Style) Then
                If Not IsDelivered(tMargin.Status) Then
                    OrderHasUndeliveredItems = True
                    tMargin.DataAccess.Records_Close()
                    tMargin = Nothing
                    Exit Function
                End If
            End If
        Loop
        tMargin.DataAccess.Records_Close()
        DisposeDA(tMargin)
    End Function

    Private Sub AdjustStoreFinance(ByVal LeaseNo As String, ByVal NewBalance As Decimal)
        ARPaySetUp.LoadAdjustmentContract(LeaseNo, NewBalance)
    End Sub

    'Form intialize event replacement in vb.net is sub new constructor.
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        mLoading = True
    End Sub

    Public Sub Cash(ByVal PayMethod As String, ByVal PayMethodInd As Integer, ByVal Memo As String)
        AddNewCashJournalRecord(PayMethodInd, GetPrice(txtBalDue.Text), SaleNo, LastName & "  " & Memo, Today)
    End Sub

    Public Sub AddCashLine(ByVal PayMethod As String, ByVal PayMethodInd As Integer, Optional ByVal Appr As String = "", Optional ByVal Amt As Decimal = -1)
        On Error Resume Next
        If Amt = -1 Then Amt = GetPrice(txtBalDue.Text)

        ' Called by CustAdjRefund.cmdApply_Click
        Margin.Name = LastName
        Margin.SaleNo = SaleNo
        Margin.Style = "SUB"
        Margin.Vendor = ""
        Margin.Status = ""
        Margin.Quantity = 0
        Margin.Desc = "                   Sub Total ="
        Margin.SellPrice = GetPrice(g_Holding.Sale) - GetPrice(g_Holding.Deposit)
        Margin.Index = Index
        AddMarginLine()

        Margin.Name = LastName
        Margin.SaleNo = SaleNo
        Margin.Style = "PAYMENT"
        Margin.Vendor = ""
        Margin.Status = ""
        If PayMethodInd = 21500 Then
            Margin.Quantity = 2  ' Display Company Check refunds as normal checks on the bill of sale.
        Else
            Margin.Quantity = PayMethodInd
        End If
        Margin.Desc = Trim("Refund By: " & PayMethod & " " & DateFormat(Today) & " " & Appr)
        Margin.SellPrice = -Math.Abs(Amt)
        Margin.SellDte = DateFormat(Today)
        'Margin.DataAccess().Records_AddAndClose
        Margin.DataAccess().Records_AddAndClose1()
        Margin.cDataAccess_SetRecordSet(Margin.DataAccess.RS)
        Margin.DataAccess.Records_AddAndClose2()

        ' reverse credit balance on sale
        g_Holding.Deposit = g_Holding.Deposit + GetPrice(Amt)
        g_Holding.Save()
    End Sub

End Class