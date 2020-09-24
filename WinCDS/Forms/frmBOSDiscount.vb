Public Class frmBOSDiscount
    Private Const C_NOTHINGSELECTED As String = "Enter Discount or Select From List"
    Private SetupIndex As Integer, SetupResult As String

    Private Sub frmBOSDiscount_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdOK, 2)
        SetButtonImage(cmdCancel, 3)
        'SetCustomFrame Me, ncBasicDialog -This line is not required. It is using to set the font and color properties using modNeoCaption module.

        SetupIndex = 0
        txtDiscountType.Visible = False
        cmbDiscountType.Visible = True
        txtDiscountAmount.Text = ""
        txtDiscountAmount.SelectionStart = 0
        txtDiscountAmount.SelectionLength = Len(txtDiscountAmount.Text)
        optAllItems.Checked = True
        UpdateControls()
        InitDiscountTypes()
        cmdCancel.Text = "&Cancel"
        'cmdCancel.Cancel = True
        Me.CancelButton = cmdCancel
        'cmbDiscountType_Click(cmbDiscountType, New EventArgs)
        cmbDiscountType_SelectedIndexChanged(cmbDiscountType, New EventArgs)
        cmbDiscountType.Select()
    End Sub

    Public Function DoDiscountTypeSetup(ByVal Index As Integer, Optional ByRef DoDelete As Boolean = False) As String
        If Index = 0 Then Err.Raise(-1, , "Invalid Index to DoDiscountTypeSetup")

        SetupIndex = Index
        SetupResult = ""

        txtDiscountType.Visible = True
        cmbDiscountType.Visible = False

        If Index <> -1 Then
            cmdCancel.Text = "&Delete"
            'cmdOK.Cancel = True
            Me.CancelButton = cmdOK
            SelectDiscountType(Index)
        End If

        'Show 1
        Me.ShowDialog()

        DoDelete = (SetupResult = "" And SetupIndex > 0)
        DoDiscountTypeSetup = SetupResult
        SetupIndex = 0

        'Unload Me
        Me.Close()
    End Function

    Private Sub UpdateControls()
        txtLastNItems.Enabled = optLastNItems.Checked
        If Not txtLastNItems.Enabled Then
            txtLastNItems.Text = "1"
        Else
            txtLastNItems.SelectionStart = 0
            txtLastNItems.SelectionLength = Len(txtLastNItems.Text) + 1
        End If
        txtFlatRate.Enabled = optFlatRate.Checked
        If Not txtFlatRate.Enabled Then
            txtFlatRate.Text = "$0.00"
        Else
            txtFlatRate.SelectionStart = 0
            txtFlatRate.SelectionLength = Len(txtFlatRate.Text) + 1
            txtDiscountAmount.Text = "0"
        End If
    End Sub

    Private Sub InitDiscountTypes()
        LoadDiscountTypesIntoComboBox(cmbDiscountType, , C_NOTHINGSELECTED)
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        If SetupIndex = 0 Then
            'Unload Me
            Me.Close()
        Else
            Hide()
        End If
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        If SetupIndex = 0 Then
            AddDiscount()
            'Unload Me
            Me.Close()
        Else
            SetupResult = BuildDiscountDef()
            Hide()
        End If
    End Sub

    Private Sub SelectDiscountType(ByVal I As Integer)
        Dim D As String, N As String, T As String, P As String, E As String

        D = DiscountType(I, N, T, P, E)
        txtDiscountType.Text = N
        txtDiscountAmount.Text = Val(P)
        Select Case T
            Case "F" ' flat rate of $<extra>
                optFlatRate.Checked = True
                txtFlatRate.Text = FormatCurrency(GetPrice(E))
            Case "A" ' all items discoutned <percent>%
                optAllItems.Checked = True
            Case "L" ' Last <extra> Items discounted <percent>%
                optLastNItems.Checked = True
                txtLastNItems.Text = Val(E)
            Case "C" ' Current Item discounted <percent>%
                optCurrentItem.Checked = True
            Case Else
                optAllItems.Checked = True
        End Select
        UpdateControls()
    End Sub

    Private Sub AddDiscount()
        Dim SubTotal As Double, Rate As Double, Dscnt As Double
        Dim I As Integer, J As Integer, N As Integer, LL As Integer
        Dim Loc As String, Msg As String

        SubTotal = 0#
        If optCurrentItem.Checked = True Then      ' only current item
            LL = BillOSale.LastLineWithItem
            If LL >= 0 Then ' if any items, record it!
                If Val(BillOSale.QueryPrice(LL)) <> 0 Then SubTotal = BillOSale.QueryPrice(LL) Else SubTotal = 0#
                Loc = BillOSale.QueryLoc(LL)
            End If
            Msg = " PRIOR ITEM"
        ElseIf optAllItems.Checked = True Then     ' all items
            GenerateSubTotalLineToBOS2()

            For I = 0 To BillOSale.UGridIO1.MaxRows - 1
                If Trim(BillOSale.QueryStyle(I)) = "" Then Exit For
                If IsItem(BillOSale.QueryStyle(I)) Or IsNote(BillOSale.QueryStyle(I)) Then
                    If Val(BillOSale.QueryPrice(I)) <> 0 Then SubTotal = SubTotal + GetPrice(BillOSale.QueryPrice(I))
                    If Loc = "" Then Loc = Trim(BillOSale.QueryLoc(I))
                End If
            Next
            Msg = " ALL ITEMS"
        ElseIf optLastNItems.Checked = True Then                              ' last n items
            N = Val(txtLastNItems.Text)
            LL = BillOSale.LastLineWithItem
            If N > 0 Then
                J = 0
                For I = LL To 0 Step -1
                    If N >= 0 Then
                        If IsItem(BillOSale.QueryStyle(I)) Then
                            If Val(BillOSale.QueryPrice(I)) <> 0 Then SubTotal = SubTotal + BillOSale.QueryPrice(I)
                            If Loc = "" Then Loc = Trim(BillOSale.QueryLoc(I))
                            J = J + 1
                            If J >= N Then Exit For
                        End If
                    End If
                Next
            End If
            Msg = " LAST " & N & " ITEM(S)"
        ElseIf optFlatRate.Checked = True Then
            Dim Amt As Decimal
            LL = BillOSale.LastLineWithItem
            If LL >= 0 Then ' if any items, record it!
                SubTotal = txtFlatRate.Text
                Loc = StoresSld
            End If
        Else
            MessageBox.Show("Please select an option.", "Wait!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Rate = GetDouble(txtDiscountAmount.Text) / 100.0#
        Dscnt = IIf(optFlatRate.Checked, SubTotal, SubTotal * Rate)
        Msg = IIf(optFlatRate.Checked, txtFlatRate.Text & " DISCOUNT", FormatPercent(Rate) & " DISCOUNT" & Msg)
        If cmbDiscountType.Text <> C_NOTHINGSELECTED Then
            Msg = cmbDiscountType.Text & ": " & Msg
        End If

        If Dscnt <= 0# Then
            MessageBox.Show("Nothing to discount!", "", MessageBoxButtons.OK)
        Else
            AddDiscountLine(Dscnt, Loc, Msg)
        End If
    End Sub

    Public Function BuildDiscountDef() As String
        Dim F As String
        If txtDiscountType.Visible Then
            F = txtDiscountType.Text
        Else
            F = cmbDiscountType.Text
        End If

        If Trim(F) = "" Then F = "DISCOUNT"
        F = F & ":"

        If optAllItems.Checked = True Then         ' all items discoutned <percent>%
            F = F & "A:" & txtDiscountAmount.Text & ":"
        ElseIf optCurrentItem.Checked = True Then  ' Current Item discounted <percent>%
            F = F & "C:" & txtDiscountAmount.Text & ":"
        ElseIf optLastNItems.Checked = True Then   ' Last <extra> Items discounted <percent>%
            F = F & "L:" & txtDiscountAmount.Text & ":" & txtLastNItems.Text
        ElseIf optFlatRate.Checked = True Then     ' flat rate of $<extra>
            F = F & "F::" & GetPrice(txtFlatRate.Text)
        End If
        BuildDiscountDef = F
    End Function

    Public Function GenerateSubTotalLineToBOS2() As Double
        Dim X As Integer

        X = BillOSale.LastLineUsed + 1
        BillOSale.RowClear(X)
        BillOSale.SetStyle(X, "SUB")
        BillOSale.SetDesc(X, "               Sub Total =")
        GenerateSubTotalLineToBOS2 = BillOSale.SubTotal(X - 1)
        BillOSale.SetPrice(X, CurrencyFormat(GenerateSubTotalLineToBOS2))
        BillOSale.NewStyleLine = X + 1
    End Function

    Private Sub frmBOSDiscount_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'RemoveCustomFrame Me -> This line is not required. It is to rollback font and color values using modNeoCaption module.
    End Sub

    Private Sub txtDiscountAmount_Enter(sender As Object, e As EventArgs) Handles txtDiscountAmount.Enter
        SelectContents(txtDiscountAmount)
    End Sub

    Private Sub txtDiscountAmount_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtDiscountAmount.Validating
        'txtDiscountAmount.Text = FormatPercent(txtDiscountAmount, , vbTrue, vbFalse, vbFalse)
    End Sub

    Private Sub optAllItems_Click(sender As Object, e As EventArgs) Handles optAllItems.Click
        UpdateControls()
    End Sub

    Private Sub optCurrentItem_Click(sender As Object, e As EventArgs) Handles optCurrentItem.Click
        UpdateControls()
    End Sub

    Private Sub optLastNItems_Click(sender As Object, e As EventArgs) Handles optLastNItems.Click
        UpdateControls()
    End Sub

    Private Sub optFlatRate_Click(sender As Object, e As EventArgs) Handles optFlatRate.Click
        UpdateControls()
    End Sub

    Private Sub txtLastNItems_Enter(sender As Object, e As EventArgs) Handles txtLastNItems.Enter
        SelectContents(txtLastNItems)
    End Sub

    Private Sub txtFlatRate_Enter(sender As Object, e As EventArgs) Handles txtFlatRate.Enter
        SelectContents(txtFlatRate)
    End Sub

    Private Sub txtFlatRate_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtFlatRate.Validating
        On Error Resume Next
        txtFlatRate.Text = FormatCurrency(txtFlatRate.Text)
    End Sub

    Private Sub cmbDiscountType_Click(sender As Object, e As EventArgs) Handles cmbDiscountType.Click
        'SelectDiscountType cmbDiscountType.itemData(cmbDiscountType.ListIndex)
        'SelectDiscountType(CType(cmbDiscountType.SelectedItem, ItemDataClass).ItemData)
    End Sub

    Private Sub cmbDiscountType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDiscountType.SelectedIndexChanged
        SelectDiscountType(CType(cmbDiscountType.SelectedItem, ItemDataClass).ItemData)
    End Sub

    Private Sub AddDiscountLine(ByVal Amount As Double, ByVal Loc As String, Optional ByVal Description As String = "")
        Dim X As Integer
        X = BillOSale.NewStyleLine
        BillOSale.X = X
        BillOSale.RowClear(X)
        If True Then ' IsDevelopment
            BillOSale.SetStyle(X, "DISCOUNT")
        Else
            BillOSale.SetStyle(X, "NOTES")
        End If
        BillOSale.SetDesc(X, Description)
        BillOSale.SetPrice(X, FormatCurrency(-Amount))
        BillOSale.SetLoc(X, "" & Loc)
        BillOSale.StyleEnabled = False
        BillOSale.MfgEnabled = False
        BillOSale.LocEnabled = False
        BillOSale.StatusEnabled = False
        BillOSale.QuanEnabled = False
        BillOSale.PriceEnabled = True
        BillOSale.DescEnabled = True
        BillOSale.DescFocus()

        BillOSale.StyleAddEnd(False, 1)
    End Sub
End Class