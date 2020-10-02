Public Class InvDel
    Public X As Integer
    Dim Margin As New CGrossMargin  '+NEW 2003-01-31AA:
    Public TaxRec1 As Decimal, TaxRec2 As Decimal, MiscDisc As Decimal, Tax2Zone As Integer
    Public TransDate As String, BSRowNum As Integer
    Private ShowDept As Boolean, ShowVend As Boolean
    Private PollingSaleDate As Boolean
    Private Const FRM_W1 = 2520
    Private Const FRM_W2 = 5610
    Private DoDeliverAll As Boolean, ContinueDelivery As Boolean
    Private NoFormLoad As Boolean

    Private Sub ShowControls()
        ' This function makes the Dept and Vendor controls visible when needed.
        ' It makes up for the CRAZY form not letting us do this when we want to.
        lblDept.Visible = ShowDept
        cboDept.Visible = ShowDept
        lblVendor.Visible = ShowVend
        cboVendor.Visible = ShowVend
        CheckDeliverEnabled(True)
    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDept.SelectedIndexChanged
        CheckDeliverEnabled(False)
    End Sub

    Private Sub cboVendor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboVendor.SelectedIndexChanged
        CheckDeliverEnabled(False)
    End Sub

    Private Sub cmdDeliverAll_Click(sender As Object, e As EventArgs) Handles cmdDeliverAll.Click
        DoDeliverAll = True
        Do While Not IsFormLoaded("OrdPay")
            ContinueDelivery = False
            'cmdDeliver.Value = True
            cmdDeliver_Click(cmdDeliver, New EventArgs)
            If Not ContinueDelivery Then GoTo GetOut
        Loop
GetOut:
        DoDeliverAll = False
    End Sub

    Private Sub cmdDeliver_Click(sender As Object, e As EventArgs) Handles cmdDeliver.Click
        If Trim(Margin.Status) = "SO" Or Trim(Margin.Status) = "PO" Then
            Dim M As String, C As CInvRec   ' BFH20050912 - Changed to prompt
            M = ""
            M = M & "This item has not been received.  If you deliver this item without first receiving it,"
            M = M & vbCrLf & "the presold will be changed to reflect this value, but if there was a PO created for"
            M = M & vbCrLf & "these items, it will no longer be able to take these items into account and your"
            M = M & vbCrLf & "On Hand could could become unreliable."
            M = M & vbCrLf
            M = M & vbCrLf & "Are you sure you want to deliver this item not received?"
            If MessageBox.Show(M, Trim(Margin.Status) & " Not Received", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Exit Sub
            End If

            C = New CInvRec
            If C.Load(Margin.Style, "Style") Then
                C.PoSold = C.PoSold - Margin.Quantity
                C.Save()
            End If
            DisposeDA(C)
        End If

        ' S/S
        If Margin.Status <> "FND" Then
            If cboDept.Visible And cboDept.SelectedIndex < 1 Then
                MessageBox.Show("You must select a department!", "Delivery")
                Exit Sub
            End If
            If cboVendor.Visible And cboVendor.SelectedIndex < 1 Then
                MessageBox.Show("You must select a vendor!", "Delivery")
                Exit Sub
            End If
        End If

        Cost.Visible = False
        Freight.Visible = False
        lblCost.Visible = False
        lblFreight.Visible = False

        BillOSale.HiLiteRow()

        DeliverInventory(Margin)

        If OrderMode("B") Then
            BillOSale.X = X    'AA:  This may cause problems, At some point, Will need to find out what BillOSale does with this.

            ' moved jk 08-06-03
            If IsIn(Trim(Margin.Status), "SS", "SO", "SOREC", "SSREC", "FND") Or IsIn(Trim(Margin.Style), "DEL", "LAB", "NOTES", "STAIN") Then
                Margin.Cost = GetPrice(Cost.Text)
                Margin.ItemFreight = Freight.Text
            End If

            ' Save new department and vendor values.
            If cboDept.Visible Then
                'Margin.DeptNo = cboDept.itemData(cboDept.ListIndex)
                Margin.DeptNo = CType(cboDept.Items(cboDept.SelectedIndex), ItemDataClass).ItemData
            End If
            If cboVendor.Visible Then
                If cboVendor.Text <> "" And cboVendor.SelectedIndex = -1 Then
                    Margin.Vendor = cboVendor.Text
                    Margin.VendorNo = ""
                Else
                    'Margin.Vendor = cboVendor.List(cboVendor.ListIndex)
                    Margin.Vendor = cboVendor.Items(cboVendor.SelectedIndex).ToString
                    'Margin.VendorNo = Format(cboVendor.itemData(cboVendor.ListIndex), "000")
                    Margin.VendorNo = Format(CType(cboVendor.Items(cboVendor.SelectedIndex), ItemDataClass).ItemData, "000")
                End If
            End If

            If Microsoft.VisualBasic.Left(Margin.Status, 1) <> "x" Then     'Items only
                If Margin.Location >= 1 And Trim(Margin.Status) <> "SS" And Trim(Margin.Status) <> "SSREC" And Trim(Margin.Style) <> "NOTES" And Trim(Margin.Status) <> "FND" Then
                    DeliverDetail
                End If

                Select Case Trim(Margin.Status)                   ' change status
                    Case "", "ST", "SO", "SS", "PO", "TW", "FND"
                        Margin.Status = "DEL" & Trim(Margin.Status)
                    Case "SOREC" : Margin.Status = "DELSOR"
                    Case "SSREC" : Margin.Status = "DELSSR"
                    Case "POREC" : Margin.Status = "DELPOR"
                    Case "LAW" : Margin.Status = "DELLAW"  ' Convert to stock in ConvertLAW before delivering.  This line should be unreachable.
                End Select
            End If

            If Microsoft.VisualBasic.Left(Margin.Status, 1) <> "x" Then               ' moved jk 08-06-03
                BillOSale.SetStatus(X, Margin.Status)
            End If

            Margin.DDelDat = DDate.Value
            Margin.Save()

            ContinueDelivery = GetNextItemOrUnload()
            BillOSale.GridMove(X)
        End If
    End Sub

    Private Sub DeliverDetail()
        Dim InvDetail As CInventoryDetail, Detail As Integer
        InvDetail = New CInventoryDetail

        If Trim(Margin.Status) = "" Or Trim(Margin.Status) = "xST" Then Exit Sub ' for notes
        On Error GoTo HandleErr

        If Margin.Detail = 0 Then 'no detail rec no
            InvDetail.DataAccess.Records_OpenSQL("SELECT * FROM Detail WHERE trim(SaleNo)=""" & ProtectSQL(Trim(Margin.SaleNo)) & """ AND trim(Style)=""" & ProtectSQL(Trim(Margin.Style)) & """ ORDER BY DetailID")
            If Not InvDetail.DataAccess.Records_Available Then
                DisposeDA(InvDetail)
                Exit Sub
            End If
            Margin.Detail = InvDetail.DetailID
            If Margin.Detail = 0 Then
                DisposeDA(InvDetail)
                Exit Sub
            End If
        Else
            If Not InvDetail.Load(CStr(Margin.Detail), "#DetailID") Then
                MessageBox.Show("Error in InvDel.DeliverDetail: Can't load Detail Record #" & Margin.Detail & ".", "Error!")
                DisposeDA(InvDetail)
                Exit Sub
            End If
        End If
        InvDetail.Trans = "DS"
        InvDetail.DDate1 = DDate.Value
        gblLastDeliveryDate = DDate.Value    ' Save delivery date in global storage.


        'needs to be moved jk 01-19-04
        If InvDetail.LAW > 0 Then
            InvDetail.AmtS1 = InvDetail.AmtS1 + Margin.Quantity
            InvDetail.LAW = 0
        End If

        InvDetail.Save()
        DisposeDA(InvDetail)
        Exit Sub

HandleErr:
        MessageBox.Show("ERROR in InvDel.DeliverDetail: " & Err.Description & ", " & Err.Source)
        If Err.Number = 13 Then Close()
        Resume Next
    End Sub

    Private Sub DeliverInventory(ByRef Margin As CGrossMargin)
        Select Case Trim(Margin.Status)
            Case "SO", "PO"
      ' Don't reduce OnHand.  This item was received without going through the PO module.
            Case "ST", "SOREC", "POREC", "LAW", "TW", "TWDEL", "DELTW"
                'If Trim(Margin.Status) = "ST" Or Trim(Margin.Status) = "SO" Or Trim(Margin.Status) = "SOREC" Or Trim(Margin.Status) = "LAW" Or Trim(Margin.Status) = "PO" Or Trim(Margin.Status) = "POREC" Then
                ' Reduce On Hand
                Dim InvData As CInvRec
                InvData = New CInvRec
                If InvData.Load(Trim(Margin.Style), "Style") Then
                    If Trim(Margin.Status) = "LAW" Then ConvertLAW(Margin, InvData)
                    InvData.OnHand = InvData.OnHand - Margin.Quantity
                    InvData.Save()
                Else
                    If Trim(Margin.Status) = "LAW" Then
                        MessageBox.Show("We could not find style " & Margin.Style & " while delivering this item." & vbCrLf & "The item quantities will not be updated.", "Unspecified Error")
                    End If
                    ' Style not found, don't hold up delivery but don't try to update anything.
                End If
                DisposeDA(InvData)
                'End If
        End Select
    End Sub

    Private Sub ConvertLAW(ByRef Margin As CGrossMargin, ByRef InvData As CInvRec)
        Margin.Status = "ST"     ' Convert the sale to Stock before delivering it.
        InvData.Available = InvData.Available - Margin.Quantity
        InvData.AddLocationQuantity(Margin.Location, -Margin.Quantity)
        If InvData.QueryStock(Margin.Location) <= 0 Then MessageBox.Show(" Item is last item or oversold! ")
    End Sub

    Private Sub cmdNotes_Click(sender As Object, e As EventArgs) Handles cmdNotes.Click
        frmNotes.DoNotes(0, Margin.SaleNo)
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        'BillOSale.cmdPrint.Value = True
        BillOSale.cmdPrint.PerformClick()
    End Sub

    Private Sub cmdSkip_Click(sender As Object, e As EventArgs) Handles cmdSkip.Click
        'Skip
        If OrderMode("B") Then
            GetNextItemOrUnload()
        End If
    End Sub

    Private Sub Cost_Leave(sender As Object, e As EventArgs) Handles Cost.Leave
        Cost.Text = CurrencyFormat(Cost)
    End Sub

    Private Sub Cost_Enter(sender As Object, e As EventArgs) Handles Cost.Enter
        SelectContents(Cost)
    End Sub

    Private Sub DDate_CloseUp(sender As Object, e As EventArgs) Handles DDate.CloseUp
        TransDate = DDate.Value
        SetLastDeliveryDate(TransDate)
    End Sub

    Private Sub DDate_Enter(sender As Object, e As EventArgs) Handles DDate.Enter
        On Error Resume Next
        If Not PollingSaleDate Then
            PollingSaleDate = True
            If Not RequestManagerApproval("Change Payment Dates", True) Then
                MessageBox.Show("You do not have access to change the sale date.", "Permission Denied")
                'cmdDeliver.SetFocus
                cmdDeliver.Select()
                PollingSaleDate = False
                Exit Sub
            End If
        Else
            PollingSaleDate = False
        End If
    End Sub

    Private Sub InvDel_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        ShowControls()
        'PositionForm()   'Note: Not required. Assigned centerscreen in properties window.
    End Sub

    Private Sub PositionForm()
        CenterForm(Me)
        ' BOS form is centered horizontally, but at the to of the screen..
        ' Hence, adjustments are relative to center horizontally, but absolute wrt top
        Left = Left + 2800
        Top = 3600
    End Sub

    Private Sub Freight_Leave(sender As Object, e As EventArgs) Handles Freight.Leave
        Freight.Text = CurrencyFormat(Freight)
    End Sub

    Private Sub Freight_Enter(sender As Object, e As EventArgs) Handles Freight.Enter
        SelectContents(Freight)
    End Sub

    Public Sub InvDelFormLoad()
        InvDel_Load(Me, New EventArgs)
    End Sub

    Private Sub InvDel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If NoFormLoad = True Then Exit Sub
        ColorDatePicker(DDate)
        'SetCustomFrame Me, ncBasicTool

        LoadDeptNamesIntoComboBox(cboDept)
        'cboDept.AddItem "Select Department", 0
        cboDept.Items.Insert(0, "Select Department")
        cboDept.SelectedIndex = 0

        LoadMfgNamesIntoComboBox(cboVendor)
        'cboVendor.AddItem "Select Manufacturer", 0
        cboVendor.Items.Insert(0, "Select Manufacturer")
        cboVendor.SelectedIndex = 0

        On Error GoTo AnError
        ' S/O S/S part of form
        Cost.Visible = False
        Freight.Visible = False
        lblCost.Visible = False
        lblFreight.Visible = False
        TaxRec1 = 0
        TaxRec2 = 0
        MiscDisc = 0
        Tax2Zone = 0

        If OrderMode("B") Then Text = "Deliver Sale"

        'DDate.Value = DateFormat(GetLastDeliveryDate)   ' Use last delivery date...
        DDate.Value = Date.Parse(DateFormat(GetLastDeliveryDate), Globalization.CultureInfo.InvariantCulture)   ' Use last delivery date...
        TransDate = DDate.Value

        'Style.ForeColor = &H8000000E
        'Style.BackColor = &H8000000D
        '  Style.BackColor = &HFFFF00
        X = -1

        '  PositionForm ' It didn't work here

        'Unload OrdPay ' jic
        OrdPay.Close()

        GetMarginItems()       ' Set up Margin data access.  This needs to be speed-enhanced somehow.
        '  GetNextItemOrUnload  ' And move to the first record.
        '  CheckDeliverEnabled
        Exit Sub

AnError:
        MessageBox.Show("ERROR in DeliverItems: " & Err.Description)
        Resume Next
    End Sub

    Private Sub GetMarginItems()
        Dim SQL As String
        ' Have to include delivered items here, or it gets out of synch with BillOSale.
        SQL = "SELECT * FROM GrossMargin where SaleNo=""" & ProtectSQL(Trim(BillOSale.BillOfSale.Text)) & """ ORDER BY MarginLine"

        ' Open Margin to the list of all undelivered items in the sale.
        Margin.DataAccess.Records_OpenSQL(SQL)
        'BillOSale.UGridIO1.GetDBGrid.FirstRow = 0
        'BillOSale.UGridIO1.GetDBGrid.FirstRow = 1
        BSRowNum = 0
    End Sub

    Private Sub InvDel_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'RemoveCustomFrame Me
        'Top = 3650
        DisposeDA(Margin)    ' Clean up the object.
    End Sub

    Public Sub ShowModal(ByRef ParentForm As Form, Optional ByVal Mdl As Boolean = True)
        If GetNextItemOrUnload() Then
            If Mdl Then
                'Show vbModal, ParentForm
                NoFormLoad = True
                Me.ShowDialog(ParentForm)
            Else
                Show()
            End If
        End If
    End Sub

    Private Function GetNextItemOrUnload() As Boolean
        If Not GetNextItem() Then UnloadForm : Exit Function
        'If Not GetNextItem() Then
        '    'UnloadForm
        '    Me.Close()
        '    Exit Function
        'End If
        GetNextItemOrUnload = True
        BillOSale.HiLiteRow(X)
    End Function

    Private Sub UnloadForm()
        'Top = 3650
        'Unload Me
        Me.Close()
        '  If GetPrice(BillOSale.BalDue) > 0 Then
        'Load OrdPay
        'OrdPay.HelpContextID = 43000
        OrdPay.Show() 'vbModal, BillOSale  ' Should this only show if there's a balance due?
        '  End If
    End Sub

    Private Function GetNextItem() As Boolean ' True if found
        On Error GoTo AnError
        With Margin.DataAccess()
            ' DataAccess is already open to the Sale.
            ' All we have to do is get the next record and return true,
            ' or return false if there are no more records.
            If Not .Records_Available Then
                GetNextItem = False
            Else
                Margin.cDataAccess_GetRecordSet(Margin.DataAccess.RS)
                BSRowNum = BSRowNum + 1
                'If BSRowNum Mod 19 = 0 Then BillOSale.UGridIO1.GetDBGrid.FirstRow = BSRowNum \ 19
                MoveBox()
                GetNextItem = True
                X = X + 1
                With Margin
                    ' Sales tax needs to be recorded even if it was previously delivered.
                    If Trim(.Style) = "TAX1" Then TaxRec1 = TaxRec1 + .SellPrice
                    If Trim(.Style) = "TAX2" Then TaxRec2 = TaxRec2 + .SellPrice : Tax2Zone = Trim(.Quantity)
                    If Trim(.Style) = "PAYMENT" And Trim(.Quantity) = "10" Then MiscDisc = MiscDisc + .SellPrice

                    ' Don't give the option to deliver delivered items.
                    If IsNothing(.Status) Then GetNextItem = GetNextItem() : Exit Function
                    If IsDelivered(.Status) Then GetNextItem = GetNextItem() : Exit Function
                    If Microsoft.VisualBasic.Left(.Status, 1) = "x" Then GetNextItem = GetNextItem() : Exit Function
                    Style.Text = .Style

                    If Trim(.Status) = "SS" Or Trim(.Status) = "SSLAW" Or Trim(.Status) = "FND" Then
                        If .Vendor <> "" And .VendorNo <> "" Then
                            ' Select that vendor in the box..
                            ' Or better yet, don't show the boxes.
                            cboVendor.Visible = False
                            lblVendor.Visible = False
                            ShowVend = False
                        Else
                            If .Vendor = "" Then
                                'cboVendor.ListIndex = 0
                                cboVendor.SelectedIndex = 0
                            Else
                                cboVendor.Text = .Vendor
                            End If
                            cboVendor.Visible = True
                            lblVendor.Visible = True
                            ShowVend = True
                        End If

                        If .DeptNo <> "" Then
                            cboDept.Visible = False
                            lblDept.Visible = False
                            ShowDept = False
                        Else
                            lblDept.Visible = True
                            'cboDept.ListIndex = 0
                            cboDept.SelectedIndex = 0
                            cboDept.Visible = True  ' This doesn't work if the form isn't loaded already!
                            ShowDept = True
                        End If
                    Else
                        cboDept.Visible = False
                        cboVendor.Visible = False
                        lblDept.Visible = False
                        lblVendor.Visible = False
                        ShowDept = False
                        ShowVend = False
                    End If

                    If IsIn(Trim(.Status), "SS", "SO", "SOREC", "SSREC", "FND") Or IsIn(Trim(.Style), "DEL", "LAB", "NOTES", "STAIN") Then
                        CorrectPrice()
                        .Cost = GetPrice(Cost.Text)
                        .ItemFreight = GetPrice(Freight.Text)
                        ' .Code = Code.Text
                        ' Code = Left(cboDept.Text, 1)
                        ' Mfg = Trim(Right(Combo1, 16))
                    End If
                End With
            End If
        End With
        CheckDeliverEnabled(True)
        Exit Function

AnError:
        MessageBox.Show("ERROR in DeliverItems: " & Err.Description, "WinCDS")
        Resume Next
    End Function

    Private Sub MoveBox()
        If X > 5 Then Top = 1570
    End Sub

    Private Sub CorrectPrice()
        On Error GoTo AnError

        ' S/O S/S part of form
        Cost.TabIndex = 2
        Freight.TabIndex = 3
        cmdDeliver.TabIndex = 5

        Cost.Visible = True
        'Cost.SelStart = 0
        Cost.SelectionStart = 0
        Freight.Visible = True
        lblCost.Visible = True
        lblFreight.Visible = True
        Cost.Text = CurrencyFormat(Margin.Cost)
        Freight.Text = CurrencyFormat(Margin.ItemFreight)
        Exit Sub

AnError:
        MessageBox.Show("ERROR in DeliverItems: " & Err.Description, "WinCDS")
        Resume Next
    End Sub

    Private Sub CheckDeliverEnabled(ByVal ResizeForm As Boolean)
        cmdDeliver.Enabled = True
        If Margin.Status <> "FND" Then  ' don't require dept/vendor for FND items (bfh20050815)
            'If cboVendor.Visible And cboVendor.ListIndex < 1 Then cmdDeliver.Enabled = False
            If cboVendor.Visible And cboVendor.SelectedIndex < 1 Then cmdDeliver.Enabled = False
            'If cboDept.Visible And cboDept.ListIndex < 1 Then cmdDeliver.Enabled = False
            If cboDept.Visible And cboDept.SelectedIndex < 1 Then cmdDeliver.Enabled = False
        End If

        If ResizeForm Then
            If Not cmdDeliver.Enabled Or IsDLS(Margin.Style) Or lblVendor.Visible Then
                Width = FRM_W2
            Else
                Width = FRM_W1
            End If
        End If
    End Sub
End Class