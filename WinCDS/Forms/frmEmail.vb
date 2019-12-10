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
    Private Results() As EmailResult, ResultCount As Long
    Private mPO As cPODetail
    Const FRM_H As Long = 4485
    Const FRM_W1 As Long = 6315
    Const FRM_W2 As Long = 10020

    Public Sub EmailSale(ByVal SaleNo As String, Optional ByVal StoreNo as integer = 0)
        Dim X As String, E As String, En As String
        Dim Cust As Boolean
        Mode = EmailMode.emSale

        Cust = True
        '  If MsgBox("Customer Copy (No Style Numbers)?", vbYesNo, "Customer Copy") = vbNo Then Cust = False
        If IsBFMyer Then Cust = False

        X = SaleToHTML(SaleNo, StoreNo, E, En, Cust)

        If Trim(E) = "" Then
            MsgBox("No email address in customer information!")
        ElseIf Trim(txtFromAddr.Text) = "" Then
            MsgBox("Store Email Address not specified." & vbCrLf & EMAIL_SETUP_INST, vbExclamation, "No Sender Email Address")
        Else
            E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, E, En, "Sale #" & SaleNo & " - " & txtFromName.Text, X)
            MsgBox("Email Sale: " & IIf(E = "", "Success!", "FAILURE - " & E))
        End If
    End Sub

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
        lstResults.SetSelected(AddedItem, True)
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

    Public Sub EmailSale(ByVal SaleNo As String, Optional ByVal StoreNo As Long = 0)
        Dim X As String, E As String, En As String
        Dim Cust As Boolean
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

    Public Sub EmailPartOrder(ByVal PartOrderNo As String, Optional ByVal StoreNo As Long = 0, Optional ByVal EmailAddr As String = "")
        Dim X As String, E As String, En As String, Attach As String
        Mode = EmailMode.emPartOrder

        X = PartOrderToHTML(PartOrderNo, StoreNo, E, En, , Attach)
        If E = "" And EmailAddr <> "" Then E = EmailAddr
        If X = "" Then Exit Sub

        If Trim(E) = "" Then
            MessageBox.Show("No email address in vendor information!")
        ElseIf Trim(txtFromAddr.Text) = "" Then
            MessageBox.Show("Store Email Address not specified." & vbCrLf & EMAIL_SETUP_INST, "No Sender Email Address", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, E, En, "Part Order #" & PartOrderNo & " - " & txtFromName.Text, X, , , Attach)
            MessageBox.Show("Email Parts Order: " & IIf(E = "", "Success!", "FAILURE - " & E))
        End If
    End Sub

    Public Function EmailChargeBackLetter(ByVal PON As Long, ByVal LetterType As Long, ByVal StoreNum As Long, ByVal Amount As Decimal, ByVal InvoiceNo As String) As Boolean
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

End Class