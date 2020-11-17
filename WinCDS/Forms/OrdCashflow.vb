Public Class OrdCashflow
    Private Function CashRegImg() As String
        CashRegImg = FXFile("Cash1.JPG")
    End Function

    Private Function CashRegImg2() As String
        CashRegImg2 = FXFile("Cash2.JPG")
    End Function

    Private Function CashRegSnd() As String
        CashRegSnd = FXFile("Cash03.WAV")
    End Function

    Private Sub cboCashIn_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCashIn.SelectedIndexChanged
        'cashin
        On Error GoTo HandleErr
        Dim Acct As Integer
        If cboCashIn.SelectedIndex = -1 Then
            Acct = 0
        Else
            'Acct = cboCashIn.itemData(cboCashIn.ListIndex)
            Acct = CType(cboCashIn.Items(cboCashIn.SelectedIndex), ItemDataClass).ItemData
        End If

        If Trim(Acct) = "99900" Then
            txtAuditNote.Visible = True
            lblAuditNote.Visible = True
        End If

        If Acct = 0 Then
            txtAccount.Text = ""
        Else
            txtAccount.Text = Acct
            txtAmount.Select()
            'pic.Picture = LoadPictureStd(CashRegImg)
            pic.Image = Image.FromFile(CashRegImg)
            PlayIt(CashRegSnd)
        End If

HandleErr:
        Resume Next
    End Sub

    Private Sub cboCashOut_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCashOut.SelectedIndexChanged
        'cashout
        On Error GoTo HandleErr
        Dim Acct As Integer
        If cboCashOut.SelectedIndex = -1 Then
            Acct = 0
        Else
            'Acct = cboCashOut.itemData(cboCashOut.ListIndex)
            Acct = CType(cboCashOut.Items(cboCashOut.SelectedIndex), ItemDataClass).ItemData
        End If

        If Trim(Acct) = "99800" Then
            txtAuditNote.Visible = True
            lblAuditNote.Visible = True
        End If

        If Acct = 0 Then
            txtAccount.Text = ""
        Else
            txtAccount.Text = Acct
            txtAmount.Select()
            'pic.Picture = LoadPictureStd(CashRegImg)
            pic.Image = Image.FromFile(CashRegImg)
            PlayIt(CashRegSnd)
        End If

HandleErr:
        Resume Next
    End Sub

    Private Sub cboCashIn_DropDown(sender As Object, e As EventArgs) Handles cboCashIn.DropDown
        On Error GoTo HandleErr
        cboCashOut.SelectedIndex = 0
        'pic.Picture = LoadPictureStd(CashRegImg2)
        pic.Image = Image.FromFile(CashRegImg2)
        PlayIt(CashRegSnd)
        Exit Sub
HandleErr:
        Resume Next
    End Sub

    Private Sub cboCashOut_DropDown(sender As Object, e As EventArgs) Handles cboCashOut.DropDown
        On Error GoTo HandleErr
        cboCashIn.SelectedIndex = 0
        'pic.Picture = LoadPictureStd(CashRegImg2)
        pic.Image = Image.FromFile(CashRegImg2)
        PlayIt(CashRegSnd)
        Exit Sub
HandleErr:
        Resume Next
    End Sub

    Private Sub cmdPost_Click(sender As Object, e As EventArgs) Handles cmdPost.Click
        Dim R As String, A As Decimal, Memo As String
        Dim StoreNum As Integer
        Dim postToBank As Boolean

        'Post entry
        On Error GoTo HandleErr
        A = GetPrice(txtAmount.Text)
        txtAmount.Text = CurrencyFormat(A)

        If A = 0 Then
            MessageBox.Show("Amount Missing", "WinCDS")
            txtAmount.Select()
            Exit Sub
        End If

        ' wants to open the drawer manually here instead of 5 times
        '    OpenCashDrawer

        postToBank = StoreSettings.bBankManagerPost
        If postToBank And IsIn(Trim(txtAccount.Text), "10200", "10300", "10400", "10500", "10600", "10650", "10250") Then
            If StoreSettings.bPostToLoc1 Then StoreNum = 1 Else StoreNum = StoresSld   ' If the store uses a single bank account for all locations, only write to database 1.

            ' BFH20090824 - Now is one or the other... if they have post to bank but use QB, it does not post to Bank1.mdb
            If Not UseQB(R) Then
                SetBankAccount(GetDatabaseBK(StoreNum), txtAccount.Text, txtAmount.Text, cboCashOut.Text, DDate.Value, StoresSld)
            End If

            ' only cash outs go to QB... rest will eventually go through a cash out
            If UseQB(R) Then
                Memo = Trim(cboCashOut.Text) & " Loc" & StoresSld
                If A > 0 Then
                    If Not QBCreateDeposit(DDate.Value,
                        , QueryGLQBAccountMap("10200"), , QueryGLQBAccountMap("01200"), Memo, ,
                        , , A, QBCustomerDepositsListID, QBCustomerDepositsName,
                        , QBLocationClassID(StoresSld, True)) Then
                        MessageBox.Show("Cash drawer posting to Quickbooks failed.", "Warning")
                    End If
                Else
                    If Not QBCreateJournalEntry(DDate.Value, , ,
                        , , QueryGLQBAccountMap("01200"), Math.Abs(GetPrice(A)), Memo, , QBCustomerDepositsName, , , QBLocationClassID(StoresSld, True),
                        , , QueryGLQBAccountMap("10200"), Math.Abs(GetPrice(A)), Memo, , , , QBLocationClassID(StoresSld, True)
                        ) Then
                        MessageBox.Show("Cash drawer posting to Quickbooks failed.", "Warning")
                    End If
                    '            MsgBox "Negative deposits not allowed by Quickbooks.", vbExclamation, "Not Allowed"
                End If
            Else
                If QBWanted() Then MessageBox.Show("Could not post cash drawer entry to Quickbooks." & vbCrLf & R, "QB Support Selected but not Available")
            End If    ' end UseQB

        End If ' end postToBank

        Cash()

        'cmdNext.Value = True   ' Clear the amount, to avoid re-posting.
        cmdNext_Click(cmdNext, New EventArgs)
        Exit Sub

HandleErr:
        Resume Next
    End Sub

    Private Sub Cash()
        Dim Note As String

        If Trim(cboCashIn.Items(cboCashIn.SelectedIndex).ToString) <> "" Then
            Note = cboCashIn.Items(cboCashIn.SelectedIndex).ToString
        Else
            Note = cboCashOut.Items(cboCashOut.SelectedIndex).ToString
        End If
        On Error GoTo HandleErr

        If Trim(txtAccount.Text) = "99900" And Trim(txtAuditNote.Text) <> "" Then
            Note = txtAuditNote.Text
        End If
        If Trim(txtAccount.Text) = "99800" And Trim(txtAuditNote.Text) <> "" Then
            Note = txtAuditNote.Text
        End If

        If cboCashIn.SelectedIndex = 8 Then txtAmount.Text = -txtAmount.Text   'Incoming resale
        AddNewCashJournalRecord(Trim(txtAccount.Text), GetPrice(txtAmount.Text), "", Trim(Note), Date.Parse(DateFormat(DDate.Value), Globalization.CultureInfo.InvariantCulture))
        Exit Sub

HandleErr:
        MessageBox.Show("ERROR: OrdCashflow.Cash New Cash Database!  " & Err.Description & ", " & Err.Source & Err.Number, "WinCDS")
        Resume Next
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        ' Next item
        txtAccount.Text = ""
        txtAmount.Text = ""
        cboCashIn.SelectedIndex = 0
        cboCashOut.SelectedIndex = 0
        txtAuditNote.Visible = False
        lblAuditNote.Visible = False
        On Error Resume Next
        txtAccount.Select()
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        'Cancel
        'Unload Me
        Me.Close()
    End Sub

    Private Sub AddListItem(ByRef Cbo As ComboBox, ByVal Description As String, Optional ByVal itemData As Integer = 0)
        'Cbo.AddItem Description
        'Cbo.itemData(Cbo.NewIndex) = itemData
        Cbo.Items.Add(New ItemDataClass(Description, itemData))
    End Sub

    Private Sub OrdCashflow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdPost, 2)
        SetButtonImage(cmdNext, 6)
        SetButtonImage(cmdCancel, 3)

        DDate.Value = Today

        ' Cash In
        AddListItem(cboCashIn, "                     ", 0)
        ' AddListItem cboCashIn, "Master/Visa Check    ", 10700
        AddListItem(cboCashIn, "Merch Credit/Forfeit ", 41500)
        AddListItem(cboCashIn, "Medical Co-Pay       ", 61600)
        ' AddListItem cboCashIn, "Finance Co. Rebate   ", 70000
        AddListItem(cboCashIn, "Beginning Cash       ", 90000)
        AddListItem(cboCashIn, "Misc. Cash In        ", 99900)
        AddListItem(cboCashIn, "Resale Payments      ", 11500)
        '  AddListItem cboCashIn, "Check Refund Adj.    ", 21500
        ' AddListItem cboCashIn, "Resale               ", 50200


        ' Cash Out
        AddListItem(cboCashOut, "                     ", 0)
        AddListItem(cboCashOut, "Bank Deposit         ", 10200)

        '   AddListItem cboCashOut, "M/C-Visa Deposit     ", 10200
        '   AddListItem cboCashOut, "Discover Deposit     ", 10200
        '   AddListItem cboCashOut, "American Exp Dep.    ", 10200
        '   AddListItem cboCashOut, "Debit Card Deposit   ", 10200
        '   AddListItem cboCashOut, "Store Credit Card Dep", 10200
        '
        AddListItem(cboCashOut, "M/C-Visa Deposit     ", 10300)
        AddListItem(cboCashOut, "Discover Deposit     ", 10400)
        AddListItem(cboCashOut, "American Exp Dep.    ", 10500)
        AddListItem(cboCashOut, "Debit Card Deposit   ", 10600)
        AddListItem(cboCashOut, "Store Credit Card Dep", 10650)

        AddListItem(cboCashOut, "Electronic Check     ", 10250)
        AddListItem(cboCashOut, "Petty Cash Out       ", 10000)
        AddListItem(cboCashOut, "Freight Bill         ", 50500)
        AddListItem(cboCashOut, "Discount On Financing", 50600)
        AddListItem(cboCashOut, "Gas & Oil            ", 60100)
        AddListItem(cboCashOut, "Maintaince           ", 62300)
        AddListItem(cboCashOut, "Credit Card Expense  ", 60500)
        AddListItem(cboCashOut, "Repair & Refinishing ", 62400)
        AddListItem(cboCashOut, "Warehouse Supplies   ", 63500)
        AddListItem(cboCashOut, "Office Supplies      ", 64100)
        AddListItem(cboCashOut, "Casual Labor         ", 65200)
        AddListItem(cboCashOut, "Meals & Entertainment", 67500)
        AddListItem(cboCashOut, "Misc Cash Out        ", 99800)
        '   AddListItem cboCashOut, "Purchases & Resales  ", 50200

        On Error GoTo HandleErr
        'pic.Picture = LoadPictureStd(CashRegImg)
        pic.Image = Image.FromFile(CashRegImg)

        Exit Sub
HandleErr:
        Resume Next
    End Sub

    Private Sub OrdCashflow_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        MainMenu.Show()
        'MMControl1.Command = "Close"
    End Sub

    Private Sub txtAmount_Enter(sender As Object, e As EventArgs) Handles txtAmount.Enter
        SelectContents(txtAmount)
    End Sub

    Private Sub txtAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAmount.KeyPress
        'If KeyAscii = Asc(",") Then KeyAscii = Asc(" ") :         ' change , to ;
        If e.KeyChar = "," Then
            e.KeyChar = " "
        End If
    End Sub

    Private Sub txtAmount_Leave(sender As Object, e As EventArgs) Handles txtAmount.Leave
        txtAmount.Text = CurrencyFormat(GetPrice(txtAmount.Text))
    End Sub
End Class