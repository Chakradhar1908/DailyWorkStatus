Public Class frmEditCash
    Private LastStore As Integer, LastD1 As Date, LastD2 As Date
    Private ReturnStatus As Boolean

    ' Need a way to load up out-of-date records...
    Public Function OutOfDateCashReport(ByVal Store As Integer, ByVal D1 As Date, ByVal D2 As Date) As Boolean
        If Not LoadOutOfDateRecords(Store, D1, D2) Then
            OutOfDateCashReport = True
            'Unload Me
            Me.Close()
            Exit Function
        End If
        Text = "Out of Date Cash Journal Entries"
        cmdContinueAudit.Visible = True
        cmdCancelAudit.Visible = True
        'Show vbModal
        ShowDialog()
        OutOfDateCashReport = ReturnStatus
    End Function

    Private Sub SetupGrid()
        grdCashJournal.Rows = 1
        'grdCashJournal.TextMatrix(0, 0) = "Index"
        grdCashJournal.set_TextMatrix(0, 0, "Index")
        'grdCashJournal.TextMatrix(0, 1) = "SaleNo"
        grdCashJournal.set_TextMatrix(0, 1, "SaleNo")
        'grdCashJournal.TextMatrix(0, 2) = "Amount"
        grdCashJournal.set_TextMatrix(0, 2, "Amount")
        'grdCashJournal.TextMatrix(0, 3) = "Account"
        grdCashJournal.set_TextMatrix(0, 3, "Account")
        'grdCashJournal.TextMatrix(0, 4) = "Note"
        grdCashJournal.set_TextMatrix(0, 4, "Note")
        'grdCashJournal.TextMatrix(0, 5) = "Cashier"
        grdCashJournal.set_TextMatrix(0, 5, "Cashier")
        'grdCashJournal.TextMatrix(0, 6) = "Trans Date"
        grdCashJournal.set_TextMatrix(0, 6, "Trans Date")
    End Sub

    Private Function LoadOutOfDateRecords(ByVal Store As Integer, ByVal D1 As Date, ByVal D2 As Date) As Boolean
        ' Load out-of-date records into the grid, allow edit..
        LastStore = Store
        LastD1 = D1
        LastD2 = D2
        SetupGrid()
        ' True=continue with Audit Report.

        Dim MinCash As Integer, MaxCash As Integer, ShowLots As Boolean
        Dim SQL As String, RS As ADODB.Recordset
        Dim Cj As CashJournalNew, CJMax As CashJournalNew

        SQL = "SELECT Min(CashId) as NC, Max(CashID) as XC FROM Cash WHERE ((Val(Account)>=1 and Val(Account)<=9) or Val(Account)=12 or Val(Account)=13) AND TransDate BETWEEN #" & D1 & "# AND #" & D2 & "#"
        RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(Store))
        If RS Is Nothing Then Exit Function
        If RS.EOF Then RS = Nothing : Exit Function

        If IsNothing(RS("NC").Value) Or IsNothing(RS("XC").Value) Then Exit Function
        MinCash = RS("NC").Value
        MaxCash = RS("XC").Value
        RS.Close()
        RS = Nothing

        SQL = "SELECT * FROM Cash WHERE ((Val(Account)>=1 and Val(Account)<=9) or Val(Account)=12 or Val(Account)=13) AND (CashId BETWEEN " & MinCash & " AND " & MaxCash & ") AND NOT (TransDate BETWEEN #" & D1 & "# AND #" & D2 & "#) ORDER BY CashID"
        RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(Store))
        If RS Is Nothing Then Exit Function
        If RS.EOF Then RS = Nothing : Exit Function

        If RS.RecordCount > 100 Then
            ' A large number of out of date transactions probably means one of the boundary records is out of date.
            ShowLots = (MsgBox("There are a large number of out-of-date transactions." & vbCrLf &
           "The reported records are most likely not at fault; rather, either the first or last listed records was wrongly reported as to be within this date range." & vbCrLf &
           "Do you want to see the " & RS.RecordCount & " records that are probably correct?", vbYesNo + vbQuestion) = vbYes)
            Dim RSMinMax As ADODB.Recordset
            RSMinMax = GetRecordsetBySQL("SELECT * FROM Cash WHERE CashID in (" & MinCash & ", " & MaxCash & ") ORDER BY CashID", False, GetDatabaseAtLocation(Store))
            CashJournalNew_RecordSet_Set(Cj, RSMinMax)
            AddGridLine(Cj)
            RSMinMax.MoveNext()
            CashJournalNew_RecordSet_Set(CJMax, RSMinMax)  ' Save for the end of the report...
            RSMinMax.Close()
        End If
        If RS.RecordCount <= 100 Or ShowLots Then
            Do Until RS.EOF
                CashJournalNew_RecordSet_Set(Cj, RS)
                AddGridLine(Cj)
                RS.MoveNext()
            Loop
        End If
        RS.Close()
        If CJMax.CashID <> 0 Then AddGridLine(CJMax)

        RS = Nothing
        If grdCashJournal.Rows > 1 Then SelectRow(1)
        LoadOutOfDateRecords = True
    End Function

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        ' Save the current row's altered date.
        SaveRow(grdCashJournal.Row)
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        ' Throw away the row's altered date.
        SelectRow(grdCashJournal.Row)
    End Sub

    Private Sub cmdCancelAudit_Click(sender As Object, e As EventArgs) Handles cmdCancelAudit.Click
        ReturnStatus = False
        'Unload Me
        Me.Close()
    End Sub

    Private Sub cmdContinueAudit_Click(sender As Object, e As EventArgs) Handles cmdContinueAudit.Click
        ReturnStatus = True
        'Unload Me
        Me.Close()
    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        SelectRow(0)
        LoadOutOfDateRecords(LastStore, LastD1, LastD2)
    End Sub

    Private Sub frmEditCash_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdApply, 2)
        SetButtonImage(cmdCancel, 3) ' "clear"
        SetButtonImage(cmdContinueAudit, 1) '"forward"
        SetButtonImage(cmdCancelAudit, 3) ' "cancel"
        SetButtonImage(cmdRefresh)
    End Sub

    Private Sub grdCashJournal_RowColChange(sender As Object, e As EventArgs) Handles grdCashJournal.RowColChange
        SelectRow(grdCashJournal.Row)
    End Sub

    Private Sub AddGridLine(ByRef Cj As CashJournalNew)
        grdCashJournal.Rows = grdCashJournal.Rows + 1
        'grdCashJournal.TextMatrix(grdCashJournal.Rows - 1, 0) = Cj.CashID
        grdCashJournal.set_TextMatrix(grdCashJournal.Rows - 1, 0, Cj.CashID)
        'grdCashJournal.TextMatrix(grdCashJournal.Rows - 1, 1) = Cj.LeaseNo
        grdCashJournal.set_TextMatrix(grdCashJournal.Rows - 1, 1, Cj.LeaseNo)
        'grdCashJournal.TextMatrix(grdCashJournal.Rows - 1, 2) = Cj.Money
        grdCashJournal.set_TextMatrix(grdCashJournal.Rows - 1, 2, Cj.Money)
        'grdCashJournal.TextMatrix(grdCashJournal.Rows - 1, 3) = Cj.Account
        grdCashJournal.set_TextMatrix(grdCashJournal.Rows - 1, 3, Cj.Account)
        'grdCashJournal.TextMatrix(grdCashJournal.Rows - 1, 4) = Cj.Note
        grdCashJournal.set_TextMatrix(grdCashJournal.Rows - 1, 4, Cj.Note)
        'grdCashJournal.TextMatrix(grdCashJournal.Rows - 1, 5) = Cj.Cashier
        grdCashJournal.set_TextMatrix(grdCashJournal.Rows - 1, 5, Cj.Cashier)
        'grdCashJournal.TextMatrix(grdCashJournal.Rows - 1, 6) = Cj.TransDate
        grdCashJournal.set_TextMatrix(grdCashJournal.Rows - 1, 6, Cj.TransDate)
    End Sub

    Private Sub SelectRow(ByVal Row As Integer)
        If Row = 0 Or Row >= grdCashJournal.Rows Then
            fraEditControls.Visible = False
        Else
            fraEditControls.Visible = True
            ' lblIndex.Caption = grdCashJournal.TextMatrix(Row, 0)
            'lblIndex.Text = "Index: " & grdCashJournal.TextMatrix(Row, 0)
            lblIndex.Text = "Index: " & grdCashJournal.get_TextMatrix(Row, 0)
            lblIndex.Text = lblIndex.Text & " This record is out of date sequence.  If you want to Change date click date dropdown, set correct date, click Apply.  Then click Continue"

            On Error Resume Next
            dteTransDate.Value = grdCashJournal.get_TextMatrix(Row, 6)
        End If
    End Sub

    Private Sub SaveRow(ByVal Row As Integer)
        If Row = 0 Then Exit Sub
        'grdCashJournal.TextMatrix(Row, 6) = dteTransDate.Value
        grdCashJournal.set_TextMatrix(Row, 6, dteTransDate.Value)

        Dim Cj As CashJournalNew, RS As ADODB.Recordset
        RS = GetRecordsetBySQL("SELECT * FROM Cash WHERE CashID=" & grdCashJournal.get_TextMatrix(Row, 0), , GetDatabaseAtLocation(LastStore))
        CashJournalNew_RecordSet_Set(Cj, RS)
        Cj.TransDate = dteTransDate.Value
        CashJournalNew_RecordSet_Get(Cj, RS)
        SetRecordsetByTableLabelIndex(RS, CashJournal_TABLE, CashJournal_INDEX, CStr(Cj.CashID), File:=GetDatabaseAtLocation(LastStore))
    End Sub

End Class