Public Class frmBarcode
    Private PortOpen As Boolean
    Private ClearUponAcquire As Boolean
    Private FormActivated As Boolean
    Public FromBarCodeReader As Boolean

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        'Unload Me
        Me.Close()
        If ReportsMode("Mini-Barcode Scanner") Then MainMenu.Show()
    End Sub

    Private Sub cmdClear_Click(sender As Object, e As EventArgs) Handles cmdClear.Click
        lstBarcodes.Items.Clear()
        cmdGetBarcodes.Select()
        cmdClear.Enabled = False
        cmdTransfer.Enabled = False
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        'lstBarcodes.RemoveItem lstBarcodes.ListIndex
        lstBarcodes.Items.RemoveAt(lstBarcodes.SelectedIndex)
        If lstBarcodes.Items.Count = 0 Then
            cmdGetBarcodes.Select()
            cmdDelete.Enabled = False
            cmdTransfer.Enabled = False
        Else
            lstBarcodes.SelectedIndex = 0
            cmdTransfer.Select()
        End If
    End Sub

    Private Sub cmdOptions_Click(sender As Object, e As EventArgs) Handles cmdOptions.Click
        frmBarcodeOptions.ShowDialog(Me)
        Application.DoEvents()
        If Tag = "" Then Exit Sub
        If Tag Then
            csp2Restore
            PortOpen = False
            'cmdGetBarcodes_Click() 'If Tag=False then Cancel was pressed
            cmdGetBarcodes_Click(cmdGetBarcodes, New EventArgs)
        End If
    End Sub

    Private Sub cmdGetBarcodes_Click(sender As Object, e As EventArgs) Handles cmdGetBarcodes.Click
        'MousePointer = vbHourglass
        Cursor = Cursors.WaitCursor
        Application.DoEvents()

        If Not PortOpen Then
            If Not OpenPort() Then 'Attempt to open the port
                'MousePointer = vbDefault
                Cursor = Cursors.Default
                Exit Sub
            End If
        End If

        Dim nRC As Integer
        Dim NumofBarcodes As Integer
        Dim X As Integer
        Dim Y As Integer

        Dim arrbyteBarcode(99) As Byte '100 elements
        Dim nBytesRead As Integer
        Dim bstrBarcode As String
        'Dim bstrTmp As String * 50
        Dim BarCodes As String

        lblStatus.Text = "Reading from Scanner"
        NumofBarcodes = csp2ReadData

        If NumofBarcodes > 0 Then
            Application.DoEvents()
            'Check to see that we are in ASCII mode...
            If csp2GetASCIIMode = PARAM_ON Then
                For X = 0 To (NumofBarcodes - 1)
                    lblStatus.Text = "Acquiring packet"
                    nBytesRead = csp2GetPacket(arrbyteBarcode(0), X, 100)
                    If nBytesRead > 0 Then
                        'Display the Barcode type
                        'nRC = csp2GetCodeType(arrbyteBarcode(1), bstrTmp, Len(bstrTmp))

                        'DisplayInBCWindow bstrTmp
                        bstrBarcode = ""

                        ' display the barcode in ascii
                        ' skip the length, type, .... timestamp
                        lblStatus.Text = "Converting barcode"
                        For Y = 2 To (nBytesRead - 5)
                            bstrBarcode = bstrBarcode & Chr(arrbyteBarcode(Y))
                        Next

                        'Display the timestamp
                        'nRC = csp2TimeStamp2Str(arrbyteBarcode(nBytesRead - 4), bstrTmp, Len(bstrTmp))

                        lblStatus.Text = "Preparing barcode"
                        If BarCodes = "" Then
                            BarCodes = InterpretBarcode(bstrBarcode)
                        Else
                            BarCodes = BarCodes & vbCrLf & InterpretBarcode(bstrBarcode)
                        End If
                        'lstBarcodes.AddItem InterpretBarcode(bstrBarcode)
                    End If
                Next

                lblStatus.Text = "Displaying Barcodes"
                Dim El As Object
                For Each El In Split(BarCodes, vbCrLf)
                    lstBarcodes.Items.Add(El)
                Next
                lblStatus.Text = "Clearing Scanner"
                'If ClearUponAcquire Then cmdClearCS1504.Value = True 'Clear barcodes so they aren't retreived twice (only if ClearUponAcquire = True).
                If ClearUponAcquire Then cmdClearCS1504.PerformClick()  'Clear barcodes so they aren't retreived twice (only if ClearUponAcquire = True).
                lblStatus.Text = "Port Open"
                'If lstBarcodes.ListCount > 0 Then lstBarcodes.ListIndex = 0
                If lstBarcodes.Items.Count > 0 Then lstBarcodes.SelectedIndex = 0
                cmdTransfer.Enabled = True
                cmdTransfer.Select()
                cmdDelete.Enabled = True
                cmdClear.Enabled = True
            Else
                'Add binary mode packets handling here..
                MessageBox.Show("Warning: Binary Mode ON" & vbCrLf & "CS1504 Must be in ASCII mode.", "Binary Mode", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            MessageBox.Show("No Barcodes in CS1504.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            cmdGetBarcodes.Select()
        End If
        'MousePointer = vbDefault
        Cursor = Cursors.Default

    End Sub

    Private Function OpenPort() As Boolean
        Dim ErrorCode As Integer, ErrorName As String
        Do
            On Error GoTo ErrorHandler 'Capture the "CSP2.DLL Not Found" error
            ErrorCode = csp2Init(CInt(GetCDSSetting("COM Port", CommPorts.COM2, "Barcode")))
            Select Case ErrorCode
                Case STATUS_OK '0
                    If csp2WakeUp() = STATUS_OK Then
                        lblStatus.Text = "COM" & GetCDSSetting("COM Port", CommPorts.COM2, "Barcode") + 1 & " Open"
                        PortOpen = True
                        OpenPort = True 'Port opened--return True
                        Exit Function
                    Else 'COMMUNICATIONS_ERROR
                        ErrorCode = -1
                        ErrorName = "Communication Error"
                    End If
                Case COMMUNICATIONS_ERROR '-1
                    ErrorName = "Communication Error"
                Case BAD_PARAM '-2
                    ErrorName = "Bad Paramater"
                Case SETUP_ERROR '-3
                    ErrorName = "Setup Error"
                Case INVALID_COMMAND_NUMBER ' -4
                    ErrorName = "Invalid Command Number"
                Case COMMAND_LRC_ERROR '-7
                    ErrorName = "Command LRC Error"
                Case RECEIVED_CHARACTER_ERROR '-8
                    ErrorName = "Received Character Error"
                Case GENERAL_ERROR '-9
                    ErrorName = "General Error"
                Case FILE_NOT_FOUND '2
                    ErrorName = "File Not Found"
                Case ACCESS_DENIED '5
                    ErrorName = "Access Denied"
            End Select
            If ErrorCode <> STATUS_OK Then
                lblStatus.Text = "Error: " & ErrorName
                'If MsgBox("Error: " & ErrorName, vbExclamation + vbRetryCancel, "Error in CS 1504: " & ErrorCode) = vbCancel Then
                If MessageBox.Show("Error: " & ErrorName, "Error in CS 1504: " & ErrorCode, MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation) = DialogResult.Cancel Then
                    lblStatus.Text = "Error"
                    Exit Function
                End If
            End If
        Loop
        Exit Function

ErrorHandler:
        If Err.Description = "File not found: csp2.dll" Then
            'MsgBox Error & ":  Please contact " & companyname & " to implement the" & vbCrLf & "Keychain Barcode Reader feature in your software.", vbInformation
            MessageBox.Show("The cs1504 should be plugged in, with scanned items, to check inventory quantity!", "Check Scanner", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            'Unload Me
            Me.Close()
        Else
            MessageBox.Show(Err.Description, "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Resume Next
        End If
    End Function

    Private Sub lstBarcodes_DoubleClick(sender As Object, e As EventArgs) Handles lstBarcodes.DoubleClick
        cmdTransfer_Click(cmdTransfer, New EventArgs)
    End Sub

    Private Sub cmdClearCS1504_Click(sender As Object, e As EventArgs) Handles cmdClearCS1504.Click
        Dim ErrorCode As Integer, ErrorName As String
        'MousePointer = vbHourglass
        Cursor = Cursors.WaitCursor
        If Not PortOpen Then
            If Not OpenPort() Then
                'MousePointer = vbDefault
                Cursor = Cursors.Default
                Exit Sub
            End If
        End If
        ErrorCode = csp2ClearData()
        Select Case ErrorCode
            Case STATUS_OK '0
                PortOpen = True
                '            cmdClear.value = True
                'If lstBarcodes.ListCount = 0 Then
                If lstBarcodes.Items.Count = 0 Then
                    cmdGetBarcodes.Select()
                Else
                    lstBarcodes.SelectedIndex = 0
                    cmdTransfer.Enabled = True
                    cmdTransfer.Select()
                End If
                'MousePointer = vbDefault
                Cursor = Cursors.Default
                Exit Sub
            Case COMMUNICATIONS_ERROR '-1
                ErrorName = "Communication Error"
            Case BAD_PARAM '-2
                ErrorName = "Bad Paramater"
            Case SETUP_ERROR '-3
                ErrorName = "Setup Error"
            Case INVALID_COMMAND_NUMBER ' -4
                ErrorName = "Invalid Command Number"
            Case COMMAND_LRC_ERROR '-7
                ErrorName = "Command LRC Error"
            Case RECEIVED_CHARACTER_ERROR '-8
                ErrorName = "Received Character Error"
            Case GENERAL_ERROR '-9
                ErrorName = "General Error"
            Case FILE_NOT_FOUND '2
                ErrorName = "File Not Found"
            Case ACCESS_DENIED '5
                ErrorName = "Access Denied"
        End Select
        If ErrorCode <> STATUS_OK Then
            lblStatus.Text = "Error: " & ErrorName
            If MessageBox.Show("Error: " & ErrorName, "Error Clearing CS 1504: " & ErrorCode, MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation) = DialogResult.Cancel Then
                lblStatus.Text = "Error"
                Exit Sub
            End If
        End If
        'MousePointer = vbDefault
        Cursor = Cursors.Default
    End Sub

    Private Sub cmdTransfer_Click(sender As Object, e As EventArgs) Handles cmdTransfer.Click
        Select Case BarcodeFormType
            Case 0
                If lstBarcodes.SelectedIndex = -1 Or lstBarcodes.SelectedIndex >= lstBarcodes.Items.Count Then Exit Sub
                'modBarcode.Barcode = lstBarcodes.List(lstBarcodes.ListIndex)
                modBarcode.Barcode = lstBarcodes.SelectedItem
                'lstBarcodes.RemoveItem(lstBarcodes.ListIndex)
                lstBarcodes.Items.RemoveAt(lstBarcodes.SelectedIndex)
            Case 1
                BarcodeFormType = 0 'Signal modBarcode that user has clicked OK
        End Select
        FromBarCodeReader = True
        'Unload Me
        Me.Close()
    End Sub

    Private Sub frmBarcode_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        'Wait until the form is visible before calling the GetBarcodes_Click event, because
        'the GetBarcodes_Click event causes the program to stop responding for several seconds.
        'This way, it doesn't appear to the user that the program is frozen.
        If FormActivated Then Exit Sub 'Don't repeat if the form has already been activated.  This usually occurs after the Options button has been clicked.
        FormActivated = True
        ClearUponAcquire = True 'Clear CS1504 after acquiring barcodes.
        If BarcodeListQty <> 0 Then 'If the Listbox already contains barcodes, don't try to Acquire again.
            Dim X As Integer
            For X = LBound(BarcodeList) To UBound(BarcodeList)
                lstBarcodes.Items.Add(BarcodeList(X))
            Next
            lstBarcodes.SelectedIndex = 0
            cmdTransfer.Enabled = True
            cmdTransfer.Select()
            cmdDelete.Enabled = True
            cmdClear.Enabled = True
        Else 'Otherwise, Acquire barcodes.
            'cmdGetBarcodes_Click()
            cmdGetBarcodes_Click(cmdGetBarcodes, New EventArgs)
        End If
        'frmBarcode.HelpContextID = 59710
    End Sub

    Private Sub frmBarcode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetCustomFrame Me, ncBasicTool
        Select Case BarcodeFormType
            Case 0 'Retreive a list of barcodes and return them one at a time.
                chkShowCost.Visible = False
                'lblStatus.Top = 200
                lblStatus.Location = New Point(lblStatus.Left, 220)
                lstBarcodes.Height = 200
                'Height = 280
                Me.Size = New Size(Me.Width, 280)
            Case 1 'Retreive a list of barcodes and return them all at once.
                chkShowCost.Visible = True
                cmdTransfer.Text = "OK"
                lblStatus.Top = 220
                lstBarcodes.Height = 200
                Height = 330
                'frmBarcode.HelpContextID = 59710
        End Select
    End Sub

    Private Sub frmBarcode_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'this event is replacement for form queryunload and unload events of vb6.0
        'RemoveCustomFrame Me
        If PortOpen Then
            'MousePointer = vbHourglass
            Cursor = Cursors.WaitCursor
            If PortOpen Then csp2Restore
            lblStatus.Text = "Port Closed"
            PortOpen = False
        End If
        'MousePointer = vbDefault
        Cursor = Cursors.Default
        If lstBarcodes.Items.Count <> 0 Then
            Dim X As Integer
            ReDim BarcodeList(0 To lstBarcodes.Items.Count - 1)
            For X = 0 To lstBarcodes.Items.Count - 1
                'BarcodeList(X) = lstBarcodes.List(X)
                BarcodeList(X) = lstBarcodes.SelectedItem
            Next
        End If
        BarcodeListQty = lstBarcodes.Items.Count
        FormActivated = False
    End Sub
End Class
