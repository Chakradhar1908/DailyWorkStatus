Public Class frmTransferSetup
    Private InputRow As Integer, InputCol As Integer
    Private Working As Boolean, Viewing As Boolean
    Private TransferIDs() As Object

    Public Sub QuickShowTransfer(ByVal TransferNo As String)
        LoadTransfer(TransferNo)
        cmdTransfer0.Visible = False
        cmdTransfer1.Visible = False
        cmdTransfer2.Visible = False
        cmdTransfer3.Visible = False
        cmdGo.Visible = False
        updTr.Visible = False
        'txtTransferNo.Locked = True
        txtTransferNo.ReadOnly = True
        grd.Enabled = False
        txtInput.Enabled = False
        cmdOK.Visible = False
        'Show 1
        ShowDialog()
    End Sub

    Public Sub LoadTransfer(ByVal TN As String)
        Dim X As String
        InitGrid("SELECT * FROM Detail WHERE Trans IN ('TP','TR','TV') AND [Misc]='" & TN & "' ORDER BY DetailID", True)
    End Sub

    Private Sub InitGrid(ByVal SQL As String, Optional ByVal View As Boolean = False)
        Dim I As Integer, N As Integer, C As CInvRec, X As Integer, Y As Integer
        Dim R As ADODB.Recordset

        Viewing = View

        Working = True
        'MousePointer = vbHourglass
        Cursor = Cursors.WaitCursor
        grd.Visible = False
        Application.DoEvents()

        If Not View Then
            C = New CInvRec
            C.DataAccess.Records_OpenSQL(SQL)
            Y = C.DataAccess.Record_Count
        Else
            R = GetRecordsetBySQL(SQL, , GetDatabaseInventory)
            Y = R.RecordCount
            If Y = 0 Then
                MessageBox.Show("Transfer Number not found.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                grd.Visible = False
                cmdOK.Visible = False
                'MousePointer = vbDefault
                Cursor = Cursors.Default
                Exit Sub
            End If
            'ReDim TransferIDs(3 To Y + 2)
            ReDim TransferIDs(0 To Y - 1)
        End If

        X = NoOfActiveLocations

        grd.Top = IIf(Not View, 240, 1200) '840
        grd.Height = IIf(Not View, 3495, 2700)
        fraView.Visible = View

        grd.Clear()
        grd.Rows = 3 + IIf(Y = 0, 1, Y)
        grd.FixedRows = 3
        grd.Cols = 1 + X
        grd.FixedCols = 1

        'grd.ColWidth(0) = 2000
        grd.set_ColWidth(0, 2000)
        grd.Col = 0
        grd.Row = 2
        grd.CellFontBold = True
        grd.CellForeColor = Color.Blue
        For I = 1 To X
            'grd.TextMatrix(0, I) = "Loc " & I
            grd.set_TextMatrix(0, 1, "Loc " & I)
            'grd.TextMatrix(1, I) = StoreSettings(I).Name
            grd.set_TextMatrix(1, I, StoreSettings(I).Name)
            grd.Row = 2
            grd.Col = I
            grd.Text = ""
            grd.CellFontBold = True
            grd.CellForeColor = Color.Blue
            'grd.ColWidth(I) = 1000
            grd.set_ColWidth(I, 1000)
        Next

        N = 2
        If Not View Then
            Do While C.DataAccess.Records_Available
                N = N + 1
                'grd.TextMatrix(0, 0) = GetVendorByStyle(C.Style)
                grd.set_TextMatrix(0, 0, GetVendorByStyle(C.Style))
                'grd.TextMatrix(N, 0) = C.Style
                grd.set_TextMatrix(N, 0, C.Style)
                For I = 1 To X
                    'grd.TextMatrix(N, I) = "0"
                    grd.set_TextMatrix(N, I, "0")
                    'grd.RowHeight(N) = 300
                    grd.set_RowHeight(N, 300)
                Next
            Loop
            DisposeDA(C)
        Else
            Dim SD As Date, ST As String, TN As String, vN As String, NN As String
            SD = #1/1/2000#

            lblSchedule.Visible = True
            dtpSchedule.Visible = True
            lblNote.Visible = True
            txtNote.Visible = True
            chkCompleted.Visible = True

            Do While Not R.EOF
                N = N + 1
                'TransferIDs(N) = IfNullThenZero(R("DetailID").Value)
                TransferIDs(N - 3) = IfNullThenZero(R("DetailID").Value)
                'grd.TextMatrix(N, 0) = R("Style")
                grd.set_TextMatrix(N, 0, R("Style").Value)
                For I = 1 To X
                    'grd.TextMatrix(N, I) = IfNullThenZeroDouble(R("Loc" & I))
                    grd.set_TextMatrix(N, 1, IfNullThenZeroDouble(R("Loc" & I).Value))
                    'grd.RowHeight(N) = 300
                    grd.set_RowHeight(N, 300)
                    grd.Row = N
                    grd.Col = 0
                    grd.CellForeColor = TransferViewRowColor(IfNullThenNilString(R("Trans").Value))
                    If IsDate(R("ddate1").Value) Then
                        If DateDiff("d", SD, DateValue(R("ddate1").Value)) > 0 Then SD = DateValue(R("ddate1").Value)
                    End If
                    If vN = "" Then vN = GetVendorByStyle(R("Style").Value)
                    If ST = "" Then   ' if any are open, whole is
                        ST = IfNullThenNilString(R("Trans").Value)
                    ElseIf ST = "TV" And IfNullThenNilString(R("Trans").Value) <> "TV" Then
                        ST = IfNullThenNilString(R("Trans").Value)
                    ElseIf ST = "TR" And IfNullThenNilString(R("Trans").Value) = "TP" Then
                        ST = IfNullThenNilString(R("Trans").Value)
                    End If
                    If TN = "" Then TN = IfNullThenNilString(R("Misc").Value)
                    On Error Resume Next
                    NN = IfNullThenNilString(R("Notes").Value)
                Next

                R.MoveNext()
            Loop
            txtSetupDate.Text = SD
            txtStatus.Text = DescribeTransferStatus(ST)
            txtTransferNo.Text = TN
            txtVendor.Text = vN
            txtDisplayNote.Text = NN

            lblSchedule.Visible = False
            dtpSchedule.Visible = False
            lblNote.Visible = False
            txtNote.Visible = False
            chkCompleted.Visible = False
        End If

        grd.Row = 3
        grd.Col = 1
        'LoadItemIntoDisplayRow grd.TextMatrix(3, 0)
        LoadItemIntoDisplayRow(grd.get_TextMatrix(3, 0))
        LoadCurrentRowStatus()    ' only for view, but checked internally
        AllowInput()

        Working = True
        grd.Visible = True
        Application.DoEvents()
        Working = False
        'MousePointer = vbDefault
        Cursor = Cursors.Default
    End Sub

    Private Sub LoadItemIntoDisplayRow(ByVal Style As String)
        Dim C As CInvRec, I As Integer
        C = New CInvRec
        C.Load(Style, "Style")
        'grd.TextMatrix(2, 0) = Style
        grd.set_TextMatrix(2, 0, Style)
        For I = 1 To NoOfActiveLocations
            'grd.TextMatrix(2, I) = C.QueryStock(I)
            grd.set_TextMatrix(2, I, C.QueryStock(I))
        Next
        DisposeDA(C)
    End Sub

    Private Sub LoadCurrentRowStatus()
        Dim R As ADODB.Recordset, S As String, SS As String
        If Not Viewing Then Exit Sub
        On Error GoTo None
        'R = GetTransferRSByTID(TransferIDs(grd.Row))
        R = GetTransferRSByTID(TransferIDs(grd.Row - 3))
        S = IfNullThenNilString(R("Trans").Value)
        If S = "TP" Then SS = DescribeTransferStatus(S) & ": Schd " & IfNullThenNilString(R("ddate1").Value)
        If S = "TR" Then SS = DescribeTransferStatus(S) & ": " & IfNullThenNilString(R("ddate1").Value)
        If S = "TV" Then SS = DescribeTransferStatus(S) & ": " & IfNullThenNilString(R("ddate1").Value)
        txtLineStatus.Text = S
        cmdTransfer0.Enabled = (S = "TP")
        cmdTransfer1.Enabled = (S = "TP")
        cmdTransfer2.Enabled = S = "TP"
        cmdTransfer3.Enabled = S = "TP"
None:
    End Sub

    Private Sub AllowInput(Optional ByVal Show As Boolean = True)
        If Inven = "View Transfer" Then Show = False
        If Not Show Then txtInput.Visible = False : InputRow = -1 : InputCol = -1 : Exit Sub
        On Error Resume Next
        InputRow = grd.Row
        InputCol = grd.Col
        'If IsDevelopment And (InputRow = -1 Or InputCol = -1) Then Stop
        'txtInput.Move grd.CellLeft + grd.Left, grd.CellTop + grd.Top, grd.CellWidth, grd.CellHeight
        txtInput.Location = New Point(grd.CellLeft + grd.Left, grd.CellTop + grd.Top)
        txtInput.Size = New Size(grd.CellWidth, grd.CellHeight)
        txtInput.Text = grd.get_TextMatrix(grd.Row, grd.Col)
        SelectContents(txtInput)
        txtInput.Visible = True
        If txtInput.Visible = True Then txtInput.Select()
    End Sub
End Class