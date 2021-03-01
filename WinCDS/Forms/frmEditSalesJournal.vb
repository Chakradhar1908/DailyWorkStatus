Public Class frmEditSalesJournal
    Private LastStore As Integer, LastD1 As Date, LastD2 As Date
    Private ReturnStatus As Boolean

    ' Need a way to load up out-of-date records...
    Public Function OutOfDateSalesReport(ByVal Store As Integer, ByVal D1 As Date, ByVal D2 As Date) As Boolean
        If Not LoadOutOfDateRecords(Store, D1, D2) Then
            OutOfDateSalesReport = True
            'Unload Me
            Me.Close()
            Exit Function
        End If
        Text = "Out of Date Sales Journal Entries"
        cmdContinueAudit.Visible = True
        cmdCancelAudit.Visible = True
        'Show vbModal
        ShowDialog()
        OutOfDateSalesReport = ReturnStatus
    End Function

    Private Function LoadOutOfDateRecords(ByVal Store As Integer, ByVal D1 As Date, ByVal D2 As Date) As Boolean
        ' Load out-of-date records into the grid, allow edit..
        LastStore = Store
        LastD1 = D1
        LastD2 = D2
        SetupGrid()
        ' True=continue with Audit Report.

        Dim MinAudit As Integer, MaxAudit As Integer, ShowLots As Boolean
        Dim SQL As String, RS As ADODB.Recordset
        Dim Sj As SalesJournalNew, SJMax As SalesJournalNew

        SQL = "SELECT Min(AuditId) as NS, Max(AuditID) as XS FROM Audit WHERE TransDate BETWEEN #" & D1 & "# AND #" & D2 & "#"
        RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(Store))
        If RS Is Nothing Then Exit Function
        If RS.EOF Then RS = Nothing : Exit Function

        If IsNothing(RS("NS").Value) Or IsNothing(RS("XS").Value) Then Exit Function
        MinAudit = RS("NS").Value
        MaxAudit = RS("XS").Value
        RS.Close()
        RS = Nothing

        SQL = "SELECT * FROM Audit WHERE (AuditID BETWEEN " & MinAudit & " AND " & MaxAudit & ") AND NOT (TransDate BETWEEN #" & D1 & "# AND #" & D2 & "#) ORDER BY AuditID"
        RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(Store))
        If RS Is Nothing Then Exit Function
        If RS.EOF Then Exit Function

        If RS.RecordCount > 100 Then
            ' A large number of out of date transactions probably means one of the boundary records is out of date.
            ShowLots = (MsgBox("There are a large number of out-of-date transactions." & vbCrLf &
           "The reported records are most likely not at fault; rather, either the first or last listed records was wrongly reported as to be within this date range." & vbCrLf &
           "Do you want to see the " & RS.RecordCount & " records that are probably correct?", vbYesNo + vbQuestion) = vbYes)
            Dim RSMinMax As ADODB.Recordset
            RSMinMax = GetRecordsetBySQL("SELECT * FROM Audit WHERE AuditID in (" & MinAudit & ", " & MaxAudit & ") ORDER BY AuditID", False, GetDatabaseAtLocation(Store))
            SalesJournalNew_RecordSet_Set(Sj, RSMinMax)
            AddGridLine(Sj)
            RSMinMax.MoveNext()
            SalesJournalNew_RecordSet_Set(SJMax, RSMinMax)  ' Save for the end of the report...
            RSMinMax.Close()
        End If
        If RS.RecordCount <= 100 Or ShowLots Then
            Do Until RS.EOF
                SalesJournalNew_RecordSet_Set(Sj, RS)
                AddGridLine(Sj)
                RS.MoveNext()
            Loop
        End If
        RS.Close()
        If SJMax.AuditID <> 0 Then AddGridLine(SJMax)

        RS = Nothing
        If grdSalesJournal.Rows > 1 Then SelectRow(1)
        LoadOutOfDateRecords = True
    End Function

    Private Sub SetupGrid()
        grdSalesJournal.Rows = 1
        grdSalesJournal.set_TextMatrix(0, 0, "Index")
        grdSalesJournal.set_TextMatrix(0, 1, "SaleNo")
        grdSalesJournal.set_TextMatrix(0, 2, "Name")
        grdSalesJournal.set_TextMatrix(0, 3, "TransDate")
        grdSalesJournal.set_TextMatrix(0, 4, "Written")
        grdSalesJournal.set_TextMatrix(0, 5, "Tax Charged")
        grdSalesJournal.set_TextMatrix(0, 6, "ArCashSls")
        grdSalesJournal.set_TextMatrix(0, 7, "Control")
        grdSalesJournal.set_TextMatrix(0, 8, "Undelivered")
        grdSalesJournal.set_TextMatrix(0, 9, "Delivered")
        grdSalesJournal.set_TextMatrix(0, 10, "Tax Recv")
        grdSalesJournal.set_TextMatrix(0, 11, "Tax Code")
        grdSalesJournal.set_TextMatrix(0, 12, "Salesman")
    End Sub

    Private Sub AddGridLine(ByRef Sj As SalesJournalNew)
        grdSalesJournal.Rows = grdSalesJournal.Rows + 1
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 0, Sj.AuditID)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 1, Sj.SaleNo)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 2, Sj.Name1)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 3, Sj.TransDate)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 4, Sj.Written)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 5, Sj.TaxCharged1)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 6, Sj.ArCashSls)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 7, Sj.Control)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 8, Sj.UndSls)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 9, Sj.DelSls)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 10, Sj.TaxRec1)
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 11, IIf(Sj.TaxCode = 0, 1, Sj.TaxCode))
        grdSalesJournal.set_TextMatrix(grdSalesJournal.Rows - 1, 12, Sj.Salesman)
    End Sub

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        ' Save the current row's altered date.
        SaveRow(grdSalesJournal.Row)
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        ' Throw away the row's altered date.
        SelectRow(grdSalesJournal.Row)
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

    Private Sub frmEditSalesJournal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdApply, 2)
        SetButtonImage(cmdCancel, 3) '"clear"
        SetButtonImage(cmdContinueAudit, 1) ' "forward"
        SetButtonImage(cmdCancelAudit, 3) ' "cancel"
        SetButtonImage(cmdRefresh)
    End Sub

    Private Sub grdSalesJournal_RowColChange(sender As Object, e As EventArgs) Handles grdSalesJournal.RowColChange
        SelectRow(grdSalesJournal.Row)
    End Sub

    Private Sub SelectRow(ByVal Row As Integer)
        If Row = 0 Or Row >= grdSalesJournal.Rows Then
            fraEditControls.Visible = False
        Else
            fraEditControls.Visible = True
            lblIndex.Text = "Index: " & grdSalesJournal.get_TextMatrix(Row, 0)
            lblIndex.Text = lblIndex.Text & " This record is out of date sequence.  If you want to Change date click date dropdown, set correct date, click Apply.  Then click Continue"

            On Error Resume Next
            dteTransDate.Value = grdSalesJournal.get_TextMatrix(Row, 3)
        End If
    End Sub

    Private Sub SaveRow(ByVal Row As Integer)
        If Row = 0 Then Exit Sub
        grdSalesJournal.set_TextMatrix(Row, 3, dteTransDate.Value)
        Dim Sj As SalesJournalNew, RS As ADODB.Recordset
        RS = GetRecordsetBySQL("SELECT * FROM Audit WHERE AuditID=" & grdSalesJournal.get_TextMatrix(Row, 0), , GetDatabaseAtLocation(LastStore))
        SalesJournalNew_RecordSet_Set(Sj, RS)
        Sj.TransDate = dteTransDate.Value
        SalesJournalNew_RecordSet_Get(Sj, RS)
        SetRecordsetByTableLabelIndex(RS, SalesJournal_TABLE, SalesJournal_INDEX, CStr(Sj.AuditID), File:=GetDatabaseAtLocation(LastStore))
    End Sub

End Class