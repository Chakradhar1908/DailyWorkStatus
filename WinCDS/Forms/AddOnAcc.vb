Imports VBRUN
Public Class AddOnAcc
    Public Typee As String
    Public ServiceNo As String
    Private ArNo As String
    Private mDisallowNew As Boolean  ' Property to disable Add feature.
    Private ShowMode As Integer
    Private SelectedValue As String
    Private mRevolved As Boolean
    Private Const Sp As String = "   "
    Private Const Sp2 As String = "  "
    Private Const Sp1 As String = " "
    Private NoFormLoad As Boolean

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
            lstAccounts.Items.Add(ArrangeString(tHold.LeaseNo, 10, AlignConstants.vbAlignRight) & Sp1 & Space(10) & ArrangeString(DescribeHoldingStatus(tHold.Status), 6, AlignConstants.vbAlignRight) & Sp1 & Space(5) & ArrangeString(FormatCurrency(tHold.Sale), 12, AlignConstants.vbAlignRight) & Sp1 & Space(8) & ArrangeString(FormatCurrency(tHold.Sale - tHold.Deposit), 12, AlignConstants.vbAlignRight))
            tHold.DataAccess.Records_MoveNext()
            tHold.cDataAccess_GetRecordSet(tHold.DataAccess.RS)
SkipItem:
        Loop
        DisposeDA(tHold)

        'If lstAccounts.ListCount = 1 Then
        If lstAccounts.Items.Count = 1 Then
            SelectEntry(0)
        Else
            lblHeadings.Text = ArrangeString("Sale", 8, AlignConstants.vbAlignRight) & Sp1 & Space(13) & ArrangeString("Status", 6) & Sp1 & Space(8) & ArrangeString("Total", 12, AlignConstants.vbAlignRight) & Sp1 & Space(11) & ArrangeString("Balance", 12, AlignConstants.vbAlignRight)
            cmdAdd.Text = "Select Sale"
            cmdNew.Text = "New Sale"
            If Not cmdNew.Enabled Then
                cmdNew.Visible = False
            End If
            cmdAddToNew.Visible = False
            If frmParent Is Nothing Then
                'Me.Show vbModal
                Me.ShowDialog()
            Else
                'Me.Show vbModal, frmParent
                NoFormLoad = True
                Me.ShowDialog(OnScreenReport)
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
            AddOnAcc_Load(Me, New EventArgs) 'IMP NOTE: Called form load explicitly here to override its default execution. In the default execution, it will firing after loading data in to listview using GetSaleNumber function and clearing all the loaded data.
        End Set
    End Property

    Private Sub SelectEntry(ByVal SelInd As Integer)
        If SelInd < 0 Then
            MessageBox.Show("You must select an item from the list.")
            Exit Sub
        End If
        'SelectedValue = ExtractArNoFromList(lstAccounts.List(SelInd))
        SelectedValue = ExtractArNoFromList(lstAccounts.Items(SelInd).ToString)
        'If cmdRevolving.Value = True Then SelectedValue = AddRevolvingSuffix(SelectedValue) ---> IMP NOTE: This a button in vb6.0. But here in vb.net, it will be a checkbox. Read below Note.
        If cmdRevolving.Checked = True Then
            SelectedValue = AddRevolvingSuffix(SelectedValue)
        End If

        'Note:  FOR THE ABOVE ERROR, REFER BOOKMARK BUTTON.VALUE
        'SOLUTION FOR THE ERROR, READ THE BELOW FOUR LINES.
        'The VB6 code used a button that stayed clicked when pressed it until it was pressed again and it unclicked.  
        'Therefore it had a 'Value' which indicated whether it was pressed or not pressed.  In .Net it is the Checkbox, not the button, that works like this.  
        'You make the Checkbox look Like a button Using the appearance Property, And your use the Checked property to test whether it is pressed or not instead of 
        'the Value property. 


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
        Dim indTab As Integer
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

    Public Function GetDetailLine(ByVal Style As String, Optional ByVal frmParent As Form = Nothing) As Integer
        ' What do we have.. only style number.
        DisallowNew = True
        ShowMode = 2
        Dim InvDetail As New CInventoryDetail
        InvDetail.DataAccess.Records_OpenSQL("SELECT * FROM Detail WHERE Style=""" & ProtectSQL(Style) & """ AND (SaleNo='' or SaleNo is null) ORDER BY DetailID")
        Do While InvDetail.DataAccess.Records_Available
            'lstAccounts.AddItem InvDetail.DetailID & vbTab & InvDetail.Style & vbTab & InvDetail.DDate1
            lstAccounts.Items.Add(InvDetail.DetailID & vbTab & InvDetail.Style & vbTab & InvDetail.DDate1)
            InvDetail.DataAccess.Records_MoveNext()
        Loop
        'Select Case lstAccounts.ListCount
        Select Case lstAccounts.Items.Count
            Case 0 : SelectedValue = 0 ' Error, no margin records found.
            Case 1 : SelectEntry(0)     ' Select the only result found.
            Case Else
                If MessageBox.Show("Automatically select the most recent order?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                    SelectEntry(0)
                Else
                    'lblHeadings.Caption = "Record    Style      Date"
                    lblHeadings.Text = "Record    Style      Date"
                    'cmdAdd.Caption = "Select Item"
                    cmdAdd.Text = "Select Item"
                    'cmdNew.Caption = "New Item"
                    cmdNew.Text = "New Item"
                    If frmParent Is Nothing Then
                        'Show vbModal
                        ShowDialog()
                    Else
                        'Show vbModal, frmParent
                        ShowDialog(frmParent)
                    End If
                End If
        End Select
        GetDetailLine = SelectedValue
        DisposeDA(InvDetail)
        'Unload Me
        Me.Close()
    End Function

    Public Function GetMarginLine(ByVal ServiceCallNo As Long, Optional ByRef frmParent As Form = Nothing) As Integer
        ' All we have is a service call number.
        Dim MailIndex As Integer

        MailIndex = GetMailIndexByServiceCallNo(ServiceCallNo)
        If MailIndex = 0 Then Exit Function

        DisallowNew = True
        ShowMode = 2
        Dim Margin As New CGrossMargin
        Margin.DataAccess.Records_OpenSQL("SELECT * FROM GrossMargin WHERE MailIndex=" & MailIndex & "  AND Style NOT IN (" & NonItemStyleString() & ") AND Left(Style,4)<>'KIT-' ORDER BY MarginLine DESC")
        Do While Margin.DataAccess.Records_Available
            'lstAccounts.AddItem Margin.MarginLine & vbTab & Margin.Style & vbTab & Margin.DDelDat
            lstAccounts.Items.Add(Margin.MarginLine & vbTab & Margin.Style & vbTab & Margin.DDelDat)
        Loop
        'Select Case lstAccounts.ListCount
        Select Case lstAccounts.Items.Count
            Case 0
                SelectedValue = 0 ' Error, no margin records found.
                MessageBox.Show("Customer has no recorded item purchases.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Case 1 : SelectEntry(0)      ' Select the only result found.
            Case Else
                If MessageBox.Show("Automatically select the most recent purchase?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                    SelectEntry(0)
                Else
                    'lblHeadings.Caption = "Record    Style      Date"
                    lblHeadings.Text = "Record    Style      Date"
                    'cmdAdd.Caption = "Select Item"
                    cmdAdd.Text = "Select Item"
                    'cmdNew.Caption = "New Item"
                    cmdNew.Text = "New Item"
                    If frmParent Is Nothing Then
                        'Show vbModal
                        ShowDialog()
                    Else
                        'Show vbModal, frmParent
                        ShowDialog(frmParent)
                    End If
                End If
        End Select
        GetMarginLine = SelectedValue
        DisposeDA(Margin)
        'Unload Me
        Me.Close()
    End Function

    Private Sub AddOnAcc_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If NoFormLoad = False Then 'IMP NOTE: Added this If block to override its default execution which clearing the data in the listview for Name option.
            'SetCustomFrame Me, ncBasicDialog
            mRevolved = False
            AdjustForm()
        End If
    End Sub

    Private Sub AddOnAcc_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        cmdNew.Enabled = Not DisallowNew
    End Sub

    Private Sub AdjustForm()
        If OrderMode("S") Then  'service
            Me.Text = "Found Existing Service Calls"
            cmdAdd.Text = "Add to Old Service"
            cmdNew.Text = "New Service Call"
            fraControls.Text = ""
            lblHeadings.Text = "Service No  Name             Telephone"
            cmdRevolving.Visible = False

            '' Removed 20140223 before ever getting used
            '  ElseIf OrderMode("A") And ModifiedRevolvingChargeEnabled() And False Then
            '    ' Expand the button frame, make R button visible
            '    cmdRevolving.Visible = True
            '    fraControls.Height = cmdRevolving.Top + cmdRevolving.Height + cmdAdd.Top
            '    lstAccounts.Height = fraControls.Height - (lstAccounts.Top - fraControls.Top)
            '    Height = fraControls.Top * 2 + fraControls.Height + 240
        Else
            'AddOnAcc.Caption = "Existing Account:"
            Me.Text = "Existing Account:"
            lblHeadings.Text = "Account No" & Sp & " Telephone           Balance"
            cmdAdd.Enabled = (StoreSettings.ExperianAcctNo = "" And StoreSettings.TransUnionAcctNo = "")
            cmdRevolving.Visible = False
        End If
    End Sub

    Private Sub AddOnAcc_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Query unload event code
        'If UnloadMode = vbFormControlMenu Then Cancel = True
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
        End If

        'Unload event code
        'RemoveCustomFrame Me
    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        'add on
        If ShowMode >= 1 Then
            'SelectEntry lstAccounts.ListIndex
            SelectEntry(lstAccounts.SelectedIndex)
            Exit Sub
        End If

        Typee = ArAddOn_Add

        If Not OrderMode("S") Then    ' Installment
            'If lstAccounts.ListIndex < 0 Then
            If lstAccounts.SelectedIndex < 0 Then
                MessageBox.Show("You must select an account to add on to!")
                Exit Sub
            End If
            'ArNo = AddOnAcc.lstAccounts.List(lstAccounts.ListIndex)
            ArNo = lstAccounts.Items(lstAccounts.SelectedIndex).ToString
            Hide()
        Else                          ' Service
            If lstAccounts.SelectedIndex < 0 Then
                MessageBox.Show(" You must select a service call to add on to! ")
                Exit Sub
            End If
            'ServiceNo = Val(Left(AddOnAcc.lstAccounts.List(lstAccounts.ListIndex), 14))
            ServiceNo = Val(Microsoft.VisualBasic.Left(lstAccounts.Items(lstAccounts.SelectedIndex).ToString, 14))
            Hide()
        End If
    End Sub

    Private Sub lstAccounts_DoubleClick(sender As Object, e As EventArgs) Handles lstAccounts.DoubleClick
        'cmdAdd.Value = True
        cmdAdd_Click(cmdAdd, New EventArgs)
    End Sub

End Class