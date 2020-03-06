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

    Private Sub OnScreenReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdAllStores_Click(sender As Object, e As EventArgs) Handles cmdAllStores.Click
        Dim X As Long
        Dim I As Long, pHolding As cHolding
        Dim J As Long
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

    Private Sub EnableControls(ByVal OnOff As Boolean, Optional ByVal Processed As Boolean = False)
        'MousePointer = IIf(OnOff, vbDefault, vbHourglass)
        Me.Cursor = IIf(OnOff, Cursors.Default, Cursors.WaitCursor)

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
                    lblRate0 = GetTax2Rate(Margin.Quantity)
                    lblRate0.ToolTipText = GetTax2String(Margin.Quantity)
                Else
                    CheckTaxLoc(Margin.Quantity)
                End If
                TaxLoc = -1
                '      If TaxLoc = 0 Then
                '        TaxLoc = Margin.Quantity + 1  ' bfh20090422 - not sure why this was out...  added if blocks
                '      End If
            End If
            ReadOut
        Loop

        cTa.Records_Close()
        'Put in totals
        RowCheck
        UGridIO1.SetValueDisplay(Row, 6, "         TOTAL DUE -->")
        UGridIO1.SetValueDisplay(Row, 8, Format(OrdTotal, "###,###.00"))
        UGridIO1.Refresh()
        lblPrevBal.Text = Format(OrdTotal, "Currency")
        Row = Row + 1
    End Sub


    Public Sub CustomerAdjustment()
        ' Load and show this form..
        ' Customer and Sale information is loaded by MailCheck's Sale Found events.
        Show()
        'cmdNext.Value = True
        cmdNext.PerformClick()
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

End Class