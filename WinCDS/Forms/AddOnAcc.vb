Public Class AddOnAcc
    Public Typee As String
    Public ServiceNo As String

    Private ArNo As String
    Private mDisallowNew As Boolean  ' Property to disable Add feature.
    Private ShowMode as integer
    Private SelectedValue As String
    Private mRevolved As Boolean

    Private Const Sp As String = "   "
    Private Const Sp2 As String = "  "
    Private Const Sp1 As String = " "
    Public Function GetSaleNumber(ByVal MailIndex As String, Optional ByRef frmParent As Form = Nothing) As String
        ' Load sale numbers into box by MailIndex.
        ' Return the selected sale.
        DisallowNew = True
        ShowMode = 1

        ' Load the sales.
        ' If we can't load anything, the caller handle it.
        Dim tHold As New cHolding
        If Not tHold.Load(MailIndex, "#Index") Then Exit Function
        Do Until tHold.DataAccess.Record_EOF
            If ArNoIsAddOnRecord(tHold.LeaseNo) Then GoTo SkipItem
            'lstAccounts.AddItem ArrangeString(tHold.LeaseNo, 10, vbAlignRight) & Sp1 & ArrangeString(DescribeHoldingStatus(tHold.Status), 6, vbAlignRight) & Sp1 & ArrangeString(FormatCurrency(tHold.Sale), 12, vbAlignRight) & Sp1 & ArrangeString(FormatCurrency(tHold.Sale - tHold.Deposit), 12, vbAlignRight)
            lstAccounts.Items.Add(ArrangeString(tHold.LeaseNo, 10, ContentAlignment.MiddleRight) & Sp1 & ArrangeString(DescribeHoldingStatus(tHold.Status), 6, ContentAlignment.MiddleRight) & Sp1 & ArrangeString(FormatCurrency(tHold.Sale), 12, ContentAlignment.MiddleRight) & Sp1 & ArrangeString(FormatCurrency(tHold.Sale - tHold.Deposit), 12, ContentAlignment.MiddleRight))

            tHold.DataAccess.Records_MoveNext()
SkipItem:
        Loop
        DisposeDA(tHold)

        'If lstAccounts.ListCount = 1 Then
        If lstAccounts.Items.Count = 1 Then
                SelectEntry(0)
            Else
            lblHeadings.Text = ArrangeString("Sale", 10, ContentAlignment.MiddleRight) & Sp1 & ArrangeString("Status", 6) & Sp1 & ArrangeString("Total", 12, ContentAlignment.MiddleRight) & Sp1 & ArrangeString("Balance", 12, ContentAlignment.MiddleRight)
            cmdAdd.Text = "Select Sale"
            cmdNew.Text = "New Sale"
            If Not cmdNew.Enabled Then cmdNew.Visible = False
            cmdAddToNew.Visible = False
            If frmParent Is Nothing Then
                'Me.Show vbModal
                Me.ShowDialog()
            Else
                'Me.Show vbModal, frmParent
                Me.ShowDialog(frmParent)
            End If
        End If
        GetSaleNumber = SelectedValue  ' Set by command buttons..
    End Function

    Public Property DisallowNew() As Boolean
        Get
            DisallowNew = mDisallowNew
        End Get
        Set(value As Boolean)
            mDisallowNew = value
            cmdNew.Enabled = Not value
        End Set
    End Property
    Private Sub SelectEntry(ByVal SelInd as integer)
        If SelInd < 0 Then
            MsgBox("You must select an item from the list.", vbExclamation, "Error")
            Exit Sub
        End If
        'SelectedValue = ExtractArNoFromList(lstAccounts.List(SelInd))
        SelectedValue = ExtractArNoFromList(lstAccounts.GetItemText(SelInd))
        'If cmdRevolving.Value = True Then SelectedValue = AddRevolvingSuffix(SelectedValue)

        'Note:  FOR THE ABOVE ERROR, REFER BOOKMARK BUTTON.VALUE

        ' Replaced by the above; kept for posterity
        '  Dim tmpItem As String, indTab as integer
        '  tmpItem = lstAccounts.List(SelInd)
        '  indTab = InStr(tmpItem, vbTab)
        '  If indTab > 0 Then
        '    tmpItem = Trim(Left(tmpItem, indTab - 1))
        '  End If
        '  SelectedValue = tmpItem
        ' And hide the modal form so the process can continue.
        If Me.Visible Then Me.Hide()
    End Sub
    Private Function ExtractArNoFromList(ByVal ListVal As String) As String
        Dim indTab as integer
        ExtractArNoFromList = ListVal
        ListVal = LTrim(ListVal)
        indTab = InStr(ListVal, vbTab)
        If indTab = 0 Then indTab = InStr(ListVal, Sp1)
        If indTab > 0 Then
            ExtractArNoFromList = Trim(Microsoft.VisualBasic.Left(ListVal, indTab - 1))
        Else
            indTab = InStr(ListVal, "        ")
            If indTab > 0 Then
                ExtractArNoFromList = Trim(Microsoft.VisualBasic.Left(ListVal, indTab - 1))
            End If
        End If
    End Function

    Public ReadOnly Property Revolved() As Boolean
        Get
            Revolved = mRevolved
        End Get
    End Property

End Class