Public Class frmEmail
    Public Mode As EmailMode
    Private Const EMAIL_SETUP_INST As String = "Goto Store Settings and enter a valid store email address."

    Public Enum EmailMode
        emSimple = 0
        emPO = 1
        emSale = 2
        emPartOrder = 3
        emChargeBack = 4
    End Enum
    Private Results() As EmailResult, ResultCount As Integer
    Private mPO As cPODetail
    Const FRM_H As Integer = 4485
    Const FRM_W1 As Integer = 6315
    Const FRM_W2 As Integer = 10020

    'Public Sub EmailSale(ByVal SaleNo As String, Optional ByVal StoreNo as integer = 0)
    '    Dim X As String, E As String, En As String
    '    Dim Cust As Boolean
    '    Mode = EmailMode.emSale

    '    Cust = True
    '    '  If MsgBox("Customer Copy (No Style Numbers)?", vbYesNo, "Customer Copy") = vbNo Then Cust = False
    '    If IsBFMyer Then Cust = False

    '    X = SaleToHTML(SaleNo, StoreNo, E, En, Cust)

    '    If Trim(E) = "" Then
    '        MsgBox("No email address in customer information!")
    '    ElseIf Trim(txtFromAddr.Text) = "" Then
    '        MsgBox("Store Email Address not specified." & vbCrLf & EMAIL_SETUP_INST, vbExclamation, "No Sender Email Address")
    '    Else
    '        E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, E, En, "Sale #" & SaleNo & " - " & txtFromName.Text, X)
    '        MsgBox("Email Sale: " & IIf(E = "", "Success!", "FAILURE - " & E))
    '    End If
    'End Sub

    Public ReadOnly Property SendMailURL() As String
        Get
            SendMailURL = WebUpdateURL & "vbSendMail.dll"
        End Get
    End Property

    Public ReadOnly Property SendMailChilkatURL() As String
        Get
            SendMailChilkatURL = WebUpdateURL & "ChilkatAx-9.5.0-win32.dll"
        End Get
    End Property

    Public ReadOnly Property Busy() As Boolean
        Get
            Busy = Not (mPO Is Nothing)
        End Get
    End Property

    Private Sub ClearResults()
        Dim R As EmailResult
        lstResults.Items.Clear()
        ResultCount = 0
        On Error Resume Next
        Results(0) = R
    End Sub

    Private Sub AddResult(ByVal PoNo As String, Optional ByVal Value As Boolean = False)
        Dim AddedItem As Integer
        'lstResults.AddItem PoNo
        'lstResults.itemData(lstResults.NewIndex) = Val(PoNo)
        'lstResults.Selected(lstResults.NewIndex) = Value
        AddedItem = lstResults.Items.Add(New ItemDataClass(PoNo, Val(PoNo)))
        'lstResults.SetSelected(AddedItem, True)
        lstResults.SetSelected(AddedItem, Value)
    End Sub

    Private Sub AddEmailResult(ByVal PoNo As String)
        ResultCount = ResultCount + 1
        ReDim Preserve Results(ResultCount - 1)
        Results(ResultCount - 1).PoNo = PoNo
    End Sub

    Public Function EmailPO(ByVal PO As cPODetail, ByVal Msg As String) As Boolean
        Dim Res As String, Addr As String, AddrName As String, Subject As String
        Dim Found As Boolean

        Mode = EmailMode.emPO
        mPO = PO
        'AddResult PO.PoNo
        AddEmailResult(PO.PoNo)
        lblStatus.Text = ("Processing PoNo " & mPO.PoNo)
        lblStatus.Refresh()

        If Trim(txtFromAddr.Text) = "" Then
            MessageBox.Show("Please enter a valid sending address for replies.", "No From Address", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            mPO = Nothing
            Exit Function
        End If

        Results(ResultCount - 1).VendorName = PO.Vendor
        Found = GetVendorFactEmail(PO.Vendor, AddrName, Addr)
        If Not Found Or Not ValidEmailAddress(Addr) Then
            '    MsgBox "Vendor Factory Email address not found for " & PO.Vendor, vbExclamation, "Could not send PO " & PO.PoNo
            mPO = Nothing
            Exit Function
        End If
        Results(ResultCount - 1).VendorAddress = Addr
        Results(ResultCount - 1).VendorName = AddrName
        Subject = "Purchase Order #" & PO.PoNo

        '  If True Then ' development
        '    Addr = "simplifiedpos@yahoo.com"
        '    AddrName = "simplifiedpos@yahoo.com"
        '  End If

        Results(ResultCount - 1).SendTime = TimeValue(Now)
        Res = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, Addr, AddrName, Subject, Msg)
        If Res <> "" Then
            MessageBox.Show("Could not Email PO: " & PO.PoNo & vbCrLf & Res)
            mPO = Nothing
        Else
            EmailPO = True

            If Inven = "EPO" Then
                AddResult(Results(ResultCount - 1).PoNo, True)
                Results(ResultCount - 1).Success = True
            End If
            mPO = Nothing
        End If
    End Function

    Public Sub EmailSale(ByVal SaleNo As String, Optional ByVal StoreNo As Integer = 0)
        Dim X As String, E As String, En As String
        Dim Cust As Boolean

        frmEmail_Load(Me, New EventArgs)
        Mode = EmailMode.emSale

        Cust = True
        '  If MsgBox("Customer Copy (No Style Numbers)?", vbYesNo, "Customer Copy") = vbNo Then Cust = False
        If IsBFMyer Then Cust = False

        X = SaleToHTML(SaleNo, StoreNo, E, En, Cust)

        If Trim(E) = "" Then
            MessageBox.Show("No email address in customer information!")
        ElseIf Trim(txtFromAddr.Text) = "" Then
            MessageBox.Show("Store Email Address not specified." & vbCrLf & EMAIL_SETUP_INST, "No Sender Email Address", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, E, En, "Sale #" & SaleNo & " - " & txtFromName.Text, X)
            MessageBox.Show("Email Sale: " & IIf(E = "", "Success!", "FAILURE - " & E))
        End If
    End Sub

    Public Sub EmailPartOrder(ByVal PartOrderNo As String, Optional ByVal StoreNo As Integer = 0, Optional ByVal EmailAddr As String = "")
        Dim X As String, E As String, En As String, Attach As String
        Mode = EmailMode.emPartOrder

        X = PartOrderToHTML(PartOrderNo, StoreNo, E, En, , Attach)
        If E = "" And EmailAddr <> "" Then E = EmailAddr
        If X = "" Then Exit Sub

        If Trim(E) = "" Then
            MessageBox.Show("No email address in vendor information!")
        ElseIf Trim(txtFromAddr.Text) = "" Then
            frmEmail_Load(Me, New EventArgs)
            If txtFromAddr.Text = "" Then
                MessageBox.Show("Store Email Address not specified." & vbCrLf & EMAIL_SETUP_INST, "No Sender Email Address", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, E, En, "Part Order #" & PartOrderNo & " - " & txtFromName.Text, X, , , Attach)
                MessageBox.Show("Email Parts Order: " & IIf(E = "", "Success!", "FAILURE - " & E))
            End If
        Else
            E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, E, En, "Part Order #" & PartOrderNo & " - " & txtFromName.Text, X, , , Attach)
            MessageBox.Show("Email Parts Order: " & IIf(E = "", "Success!", "FAILURE - " & E))
        End If
    End Sub

    Public Function EmailChargeBackLetter(ByVal PON As Integer, ByVal LetterType As Integer, ByVal StoreNum As Integer, ByVal Amount As Decimal, ByVal InvoiceNo As String) As Boolean
        Dim S As String, vEm As String, vNm As String, E As String, Attach As String

        Mode = EmailMode.emChargeBack
        S = ChargeBackLetterHTML(PON, LetterType, StoreNum, Amount, InvoiceNo, vEm, vNm, Attach)
        If S = "" Then Exit Function
        If Trim(vEm) = "" Then
            MessageBox.Show("No email address in vendor information!")
        ElseIf Trim(txtFromAddr.Text) = "" Then
            MessageBox.Show("Store Email Address not specified." & vbCrLf & EMAIL_SETUP_INST, "No Sender Email Address", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, vEm, vNm, "Charge Back - " & StoreSettings.Name, S, , , Attach)
            MessageBox.Show("Email Charge Back: " & IIf(E = "", "Success!", "FAILURE - " & E))
        End If
        EmailChargeBackLetter = (E = "")
    End Function

    Public Function EmailOrderNotAcknowledged(ByVal PoNo As String) As Boolean
        Dim S As String, C As cPODetail, Em As String, N As String, E As String
        C = New cPODetail
        If Not C.Load(PoNo, "#PoNo") Then
            DisposeDA(C)
            Exit Function
        End If

        GetVendorName(C.Vendor, N, , , , , , , , Em)
        S = OrderNotAcknowledgedLetterHTML(PoNo, C.PoDate, C.Vendor)
        If S = "" Then Exit Function
        If Trim(Em) = "" Then
            MessageBox.Show("No email address in vendor information!")
        ElseIf Trim(txtFromAddr.Text) = "" Then
            MessageBox.Show("Store Email Address not specified." & vbCrLf & EMAIL_SETUP_INST, "No Sender Email Address", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, Em, N, "Order Not Acknowledged (PO: " & PoNo & ") - " & StoreSettings.Name, S)
        End If
        EmailOrderNotAcknowledged = (E = "")
        DisposeDA(C)
    End Function

    Private Function OrderNotAcknowledgedLetterHTML(ByVal vPoNo As String, ByVal vPoDate As String, ByVal vVendor As String) As String
        Dim S As String
        S = ""
        S = S & ""
        S = S & "<html>" & vbCrLf
        S = S & "<head>" & vbCrLf
        S = S & "</head>" & vbCrLf
        S = S & "<body>" & vbCrLf
        S = S & "<b>Attn: <u>ORDER DEPARTMENT</u><br/>" & vbCrLf
        S = S & "PoNo: <font size=+1>" & vPoNo & "</font><br/>" & vbCrLf
        S = S & "Order Date: " & vPoDate & "<br/>" & vbCrLf
        S = S & "Manufacturer: " & vVendor & "<br/>" & vbCrLf
        S = S & "</b><br/>" & vbCrLf
        S = S & EmailFactOrdNotAckBodyHTML()
        S = S & "<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & StoreSettings.Name & "<br/>" & vbCrLf
        S = S & StoreSettings.Email & "<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & StoreSettings.Address & "<br/>" & vbCrLf
        S = S & StoreSettings.City & "<br/>" & vbCrLf
        S = S & StoreSettings.Phone & "<br/>" & vbCrLf
        S = S & "</body>" & vbCrLf
        S = S & "</html>" & vbCrLf
        OrderNotAcknowledgedLetterHTML = S
    End Function

    Public Function EmailOverdueOrder(ByVal PoNo As String) As Boolean
        Dim S As String, C As cPODetail, Em As String, N As String, E As String
        C = New cPODetail
        If Not C.Load(PoNo, "#PoNo") Then
            DisposeDA(C)
            Exit Function
        End If

        GetVendorName(C.Vendor, N, , , , , , , , Em)
        S = OverdueOrderLetterHTML(PoNo, C.AckInv, C.DueDate, C.Vendor, C.PoDate)
        If S = "" Then Exit Function
        If Trim(Em) = "" Then
            MessageBox.Show("No email address in vendor information!")
        ElseIf Trim(txtFromAddr.Text) = "" Then
            MessageBox.Show("Store Email Address not specified." & vbCrLf & "Goto File | Purchase Orders | Email POs and enter a valid email address.", "No Sender Email Address", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, Em, N, "OverDue Order (PO: " & PoNo & ") - " & StoreSettings.Name, S)
        End If
        EmailOverdueOrder = (E = "")
        DisposeDA(C)
    End Function

    Private Sub cmdMail_Click(sender As Object, e As EventArgs) Handles cmdMail.Click
        On Error Resume Next
        If Not ValidEmailAddress(txtFromAddr.Text) Then
            MessageBox.Show("You must enter a valid From email address.")
            txtFromAddr.Select()
            SelectContents(txtFromAddr)
            Exit Sub
        End If
        If optByPoNo.Checked = True Then
            If Val(txtFromPO.Text) = 0 Or Val(txtToPO.Text) = 0 Then
                MessageBox.Show("Please enter a valid PO number range.")
                FocusSelect(txtFromPO)
                Exit Sub
            End If
        Else
            If DateBefore(dtpToDate.Value, dtpFromDate.Value, False) Then
                MessageBox.Show("Please enter a valid PO date range.")
                dtpFromDate.Select()
                Exit Sub
            End If
        End If

        If optByDate.Checked = True Then
            Dim R As ADODB.Recordset
            R = GetRecordsetBySQL("SELECT Min([PoNo]) FROM [PO] WHERE [PoDate] BETWEEN #" & dtpFromDate.Value & " # AND #" & dtpToDate.Value & "#", , GetDatabaseInventory)
            If R.RecordCount = 0 Then
                MessageBox.Show("Could not find valid POs in date range " & dtpFromDate.Value & " to " & dtpToDate.Value & ".", "Invalid PO Date Range", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
            txtFromPO = R(0)
            DisposeDA(R)

            R = GetRecordsetBySQL("SELECT Max([PoNo]) FROM [PO] WHERE [PoDate] BETWEEN #" & dtpFromDate.Value & " # AND #" & dtpToDate.Value & "#", , GetDatabaseInventory)
            If R.RecordCount = 0 Then
                MessageBox.Show("Could not find valid POs in date range " & dtpFromDate.Value & " to " & dtpToDate.Value & ".", "Invalid PO Date Range", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
            txtToPO = R(0)
            DisposeDA(R)

            If txtFromPO.Text = "" Or txtToPO.Text = "" Then
                MessageBox.Show("Could not obtain PO #'s frmo date range " & dtpFromDate.Value & " to " & dtpToDate.Value & ".", "Invalid PO Date Range", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
            Debug.Print("PoNo Range from Dates: " & txtFromPO.Text & "-" & txtToPO.Text)
        End If

        cmdMail.Enabled = False
        cmdOK.Enabled = False

        AdjustForm(True)
        ClearResults()
        InvPoPrint.ReprintPO(txtFromPO.Text, txtToPO.Text, , Not chkReprint.Checked = True)


        lblStatus.Text = "Printing Report..." : lblStatus.Refresh()

        PrintEmailReport()

        If chkPrintPO.Checked = True Then
            modProgramState.Inven = ""
            InvPoPrint.ReprintPO(txtFromPO.Text, txtToPO.Text, chkReprint.Checked = True)
            modProgramState.Inven = "EPO"
        End If

        lblStatus.Text = ""
        cmdMail.Enabled = True
        cmdOK.Enabled = True

    End Sub

    Private Function OverdueOrderLetterHTML(ByVal vPoNo As String, ByVal vAckNo As String, ByVal vDueDate As String, ByVal vVendor As String, ByVal vPoDate As String) As String
        Dim S As String
        S = ""
        S = S & "<html>" & vbCrLf
        S = S & "<head>" & vbCrLf
        S = S & "</head>" & vbCrLf
        S = S & "<body>" & vbCrLf
        S = S & "<b>Attn: <u>ORDER DEPARTMENT</u><br/>" & vbCrLf
        S = S & "Acknowledgement No: <i>" & vAckNo & "</i><br/>" & vbCrLf
        S = S & "PoNo: <u><font size=+1>" & vPoNo & "</font></u><br/>" & vbCrLf
        S = S & "Order Date: " & vPoDate & "<br/>" & vbCrLf
        S = S & "Manufacturer: " & vVendor & "<br/>" & vbCrLf
        S = S & "Anticipated Due Date: <i>" & vDueDate & "</i><br/>" & vbCrLf
        S = S & "</b><br/>" & vbCrLf
        S = S & EmailOverdueOrdersBodyHTML() & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "Thank you,<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & StoreSettings.Name & "<br/>" & vbCrLf
        S = S & StoreSettings.Email & "<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & StoreSettings.Address & "<br/>" & vbCrLf
        S = S & StoreSettings.City & "<br/>" & vbCrLf
        S = S & StoreSettings.Phone & "<br/>" & vbCrLf
        S = S & "</body>" & vbCrLf
        S = S & "</html>" & vbCrLf
        OverdueOrderLetterHTML = S
    End Function

    Public Sub AdjustForm(Optional ByVal Wide As Boolean = False)
        Width = IIf(Not Wide, FRM_W1, FRM_W2)
    End Sub

    Public Sub PrintEmailReport()
        Dim I As Integer, A As String, Y As Integer
        On Error Resume Next
        OutputToPrinter = True
        OutputObject = Printer

        CommonReportAddColumn("PoNo", 800, True, "[RIGHT]")
        CommonReportAddColumn("Vendor", 2000)
        CommonReportAddColumn("Vendor Email", 4000)
        CommonReportAddColumn("Send Time", 1500, , "[RIGHT]")
        CommonReportAddColumn("Success", 1000)

        CommonReportHeader("Email PO Result")

        For I = 0 To ResultCount - 1
            Y = OutputObject.CurrentY
            CommonReportPrintColumn(1, CStr(Results(I).PoNo), Y)
            CommonReportPrintColumn(2, Microsoft.VisualBasic.Left(Results(I).VendorName, 18), Y)
            A = Results(I).VendorAddress
            If A <> "" Then
                CommonReportPrintColumn(3, Microsoft.VisualBasic.Left(A, 38), Y)
                CommonReportPrintColumn(4, Results(I).SendTime, Y)
                CommonReportPrintColumn(5, YesNo(Results(I).Success), Y)
            Else
                CommonReportPrintColumn(3, "No Email Address.", Y)
            End If
        Next

        If OutputToPrinter Then Printer.EndDoc() Else frmPrintPreviewDocument.DataEnd()
    End Sub

    Private Sub frmEmail_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'unload
        Select Case Mode
            Case EmailMode.emPO : MainMenu.Show()
                '    Case Else: MainMenu.Show
        End Select
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Inven = ""
        'Unload Me
        Me.Close()
    End Sub

    Private Sub txtFromPO_Leave(sender As Object, e As EventArgs) Handles txtFromPO.Leave
        If Val(txtFromPO.Text) > Val(txtToPO.Text) Then txtToPO.Text = txtFromPO.Text
    End Sub

    Private Sub txtToPO_Leave(sender As Object, e As EventArgs) Handles txtToPO.Leave
        If Val(txtFromPO.Text) > Val(txtToPO.Text) Then txtFromPO.Text = txtToPO.Text
    End Sub

    Private Sub txtFromPO_Enter(sender As Object, e As EventArgs) Handles txtFromPO.Enter
        SelectContents(txtFromPO)
    End Sub

    Private Sub txtToPO_Enter(sender As Object, e As EventArgs) Handles txtToPO.Enter
        SelectContents(txtToPO)
    End Sub

    Private Sub optByDate_Click(sender As Object, e As EventArgs) Handles optByDate.Click
        SelectRange()
    End Sub

    Private Sub optByPoNo_Click(sender As Object, e As EventArgs) Handles optByPoNo.Click
        SelectRange()
    End Sub

    Private Sub SelectRange(Optional ByVal Field As Integer = 0)
        Const HideOrDisable = True
        txtFromPO.Text = ""
        txtToPO.Text = ""
        dtpFromDate.Value = Today 'WeeksAgo(-1, NextWeekDay(vbMonday, , -1))
        dtpToDate.Value = Today 'NextDay(WeeksAgo(1, dtpFromDate), -1)
        Select Case Field
            Case 1
                If HideOrDisable Then
                    fraByPoNo.Visible = True
                    fraByDate.Visible = False
                Else
                    EnableFrame(Me, fraByPoNo, True)
                    EnableFrame(Me, fraByDate, False)
                End If
                FocusControl(txtFromPO)
            Case Else
                If HideOrDisable Then
                    fraByPoNo.Visible = False
                    fraByDate.Visible = True
                Else
                    EnableFrame(Me, fraByPoNo, False)
                    EnableFrame(Me, fraByDate, True)
                End If
                FocusControl(dtpFromDate)
        End Select
    End Sub

    Private Sub frmEmail_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage cmdOK
        'SetButtonImage cmdMail, "forward"
        SetButtonImage(cmdOK, 2)
        SetButtonImage(cmdMail, 1)
        ColorDatePicker(dtpFromDate)
        ColorDatePicker(dtpToDate)

        optByDate.Checked = True

        If Not HasSendMail() Then
            cmdMail.Enabled = False
            If MessageBox.Show("You do not have a required DLL to do email POs yet." & vbCrLf & "In order to email POs, you must install this file." & vbCrLf2 & "Would you like instructions for downloading and installing this DLL?", "Missing DLL, get instructions?", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                MessageBox.Show("Visit dthe URL:  " & SendMailURL & vbCrLf2 & "Place that file in your Windows System directory (usually C:\Windows\System32\)." & vbCrLf & "Then From the Start Menu, click 'Run'." & vbCrLf & "Enter this command: regsvr32 /s vbSendMail.dll", "DLL Download and Installation", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
        ClearResults()
        AdjustForm(False)

        txtFromName.Text = GetEmailSetting("FromName")
        txtFromAddr.Text = GetEmailSetting("FromAddr")
        lblStatus.Text = ""
    End Sub

    Public Sub LOG(Optional ByVal Msg As String = "", Optional ByVal PC As Integer = -101)
        ActiveLog("frmEmail::Log(Msg=" & Msg & ", PC=" & PC & ")", 7)
        If PC <> -101 Then
            If PC < 0 Then PC = 0
            If PC > 100 Then PC = 100
            prg.Value = PC
        End If

        If Msg <> "" Then
            txt.Text = txt.Text & vbCrLf & Msg
            txt.SelectionStart = Len(txt.Text)
        End If
    End Sub

    Public Function DoSendSimpleEmail(ByVal From As String, ByVal FromName As String, ByVal T As String, ByVal TName As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachments As String = "") As String
        DoSendSimpleEmail = SendSimpleEmail(From, FromName, T, TName, Subject, Body, CC, BCC, Attachments)
        If Inven = "EPO" Then
            mPO = Nothing
        End If
    End Function
End Class