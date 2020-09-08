Public Class frmSalesList
    Dim Sm As Object
    Private Shared mSalesCode As String
    Private medotclose As Boolean

    Public Shared Property SalesCode() As String
        Get
            SalesCode = mSalesCode
        End Get
        Set(value As String)
            mSalesCode = value
        End Set
    End Property

    Private Sub SelectAndApply(Optional ByVal CloseForm As Boolean = True)
        'If Trim(lstSalesmen.List(lstSalesmen.ListIndex)) = "" Then Exit Sub
        If lstSalesmen.SelectedItem.ToString = "" Then Exit Sub

        If BillOSale.Sales1.Text = "" Then
            'BillOSale.Sales1.Text = lstSalesmen.List(lstSalesmen.ListIndex)
            BillOSale.Sales1.Text = lstSalesmen.SelectedItem.ToString
            GetNo()
            BillOSale.SalesSplit1.Text = "100%"
            BillOSale.SalesSplit2.Text = "0%"
            BillOSale.SalesSplit3.Text = "0%"
        ElseIf BillOSale.Sales2.Text = "" Then
            'If BillOSale.Sales1.Text = lstSalesmen.List(lstSalesmen.ListIndex) Then Exit Sub
            If BillOSale.Sales1.Text = lstSalesmen.SelectedItem.ToString Then Exit Sub
            'BillOSale.Sales2.Text = lstSalesmen.List(lstSalesmen.ListIndex)
            BillOSale.Sales2.Text = lstSalesmen.SelectedItem.ToString
            GetNo()
            BillOSale.SalesSplit1.Text = "50%"
            BillOSale.SalesSplit2.Text = "50%"
            BillOSale.SalesSplit3.Text = "0%"
        ElseIf BillOSale.Sales3.Text = "" Then
            'If BillOSale.Sales1.Text = lstSalesmen.List(lstSalesmen.ListIndex) Then Exit Sub
            'If BillOSale.Sales2.Text = lstSalesmen.List(lstSalesmen.ListIndex) Then Exit Sub
            If BillOSale.Sales1.Text = lstSalesmen.SelectedItem.ToString Then Exit Sub
            If BillOSale.Sales2.Text = lstSalesmen.SelectedItem.ToString Then Exit Sub
            'BillOSale.Sales3.Text = lstSalesmen.List(lstSalesmen.ListIndex)
            BillOSale.Sales3.Text = lstSalesmen.SelectedItem.ToString
            BillOSale.SalesSplit1.Text = "33.33%"
            BillOSale.SalesSplit2.Text = "33.33%"
            BillOSale.SalesSplit3.Text = "33.33%"
            GetNo
        Else
            ' they're all filled already!!
        End If

        If CloseForm Then
            'Unload Me
            Me.Close()
            BillOSale.cmdApplyBillOSale.Select()
        End If
    End Sub

    Public Sub GetNo()
        ' Add current selection's sales code to the end of SalesCode.
        Dim Z As Integer

        'Z = lstSalesmen.ListIndex
        Z = lstSalesmen.SelectedIndex

        If Z >= LBound(Sm, 1) And Z <= UBound(Sm, 1) Then
            If Sm(Z, 2) = "" Then
                MsgBox("Salesman Error #2: Added salesman (" & lstSalesmen.Text & ") with a blank salesman number!", vbCritical)
            Else
                SalesCode = Trim(SalesCode & " " & Sm(Z, 2))
                BillOSale.SalesCode = SalesCode
            End If
            'Debug.Print "Selected salesman " & Sm(Z, 2) & " (" & Sm(Z, 1) & ")"
        Else
            MsgBox("Salesman Error #3: Added salesman (" & lstSalesmen.Text & ") without a salesman number!", vbCritical)
        End If
    End Sub

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        'Unload Me
        medotclose = True  'To avoid formclosing event execution for Me.Close().
        Me.Close()
        BillOSale.cmdApplyBillOSale.Select()
    End Sub

    Private Sub cmdClear_Click(sender As Object, e As EventArgs) Handles cmdClear.Click
        BillOSale.Sales1.Text = ""
        BillOSale.Sales2.Text = ""
        BillOSale.Sales3.Text = ""
        BillOSale.SalesCode = ""
        BillOSale.SalesSplit1.Text = "100%"
        BillOSale.SalesSplit2.Text = "0%"
        BillOSale.SalesSplit3.Text = "0%"
        SalesCode = ""
        On Error Resume Next
        lstSalesmen.Select()
    End Sub

    Private Sub frmSalesList_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        On Error Resume Next
        lstSalesmen.Select()
    End Sub

    Private Sub frmSalesList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage(cmdApply)
        'SetButtonImage(cmdClear)
        SetButtonImage(cmdApply, 2)
        SetButtonImage(cmdClear, 22)
        'SetCustomFrame(Me, ncBasicDialog)

        On Error GoTo HandleErr
        'cmdClear.Value = True
        cmdClear.PerformClick()

        Sm = GetSalesmanDatabase(StoresSld, True)

        'lstSalesmen.Clear
        lstSalesmen.Items.Clear()

        Dim EE As Integer

        For EE = LBound(Sm, 1) To UBound(Sm, 1)
            'lstSalesmen.AddItem Sm(EE, 1), EE  ' - 1
            'lstSalesmen.itemData(lstSalesmen.NewIndex) = Sm(EE, 2)
            'lstSalesmen.Items.Insert(EE, Sm(EE, 0))
            lstSalesmen.Items.Insert(EE, New ItemDataClass(Sm(EE, 0), Sm(EE, 1)))
        Next

        lstSalesmen.SetSelected(0, True)  'added this line here instead of enter event.
        Exit Sub
HandleErr:
    End Sub

    Private Sub lstSalesmen_DoubleClick(sender As Object, e As EventArgs) Handles lstSalesmen.DoubleClick
        SelectAndApply(False)         ' bfh20050803: changed dblclick and clicking apply to do same thing
    End Sub

    Private Sub frmSalesList_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        '--------> This event is replacement for form unload and queryunload events of vb 6.0 <----------'
        'If UnloadMode = vbFormControlMenu Then cmdClear.Value = True
        If medotclose = True Then Exit Sub
        If e.CloseReason = CloseReason.UserClosing Then
            cmdClear.PerformClick()
        End If
    End Sub

    'Note: This event is not required. Code will be moved to load event.
    'Private Sub lstSalesmen_Enter(sender As Object, e As EventArgs) Handles lstSalesmen.Enter
    'On Error Resume Next
    'If lstSalesmen.ListIndex < 0 Then lstSalesmen.ListIndex = 0
    'lstSalesmen.Selected(lstSalesmen.ListIndex) = True
    'If lstSalesmen.SelectedIndex < 0 Then lstSalesmen.SelectedIndex = 0
    'lstSalesmen.SetSelected(0, True)
    'End Sub

    Private Sub lstSalesmen_KeyDown(sender As Object, e As KeyEventArgs) Handles lstSalesmen.KeyDown
        'If KeyCode = vbKeyReturn Then
        'SelectAndApply False
        '    If Shift And vbCtrlMask > 0 Then
        '      SelectAndApply False
        '      lstSalesmen.SetFocus
        '    Else
        '      SelectAndApply True
        '    End If
        'End If

        If e.KeyCode = Keys.Enter Then
            SelectAndApply(False)
        End If
    End Sub
End Class