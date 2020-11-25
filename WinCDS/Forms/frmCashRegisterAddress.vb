Public Class frmCashRegisterAddress
    Private mMailIndex As Integer, MailRec As clsMailRec, MailShip As MailNew2
    Private mAtype As Integer
    Dim UserClose As Boolean

    Public Property MailIndex() As Integer
        Get
            MailIndex = mMailIndex
        End Get
        Set(value As Integer)
            Dim EmptyMailShip As MailNew2
            mMailIndex = value
            MailRec = New clsMailRec
            MailShip = EmptyMailShip
            If MailIndex > 0 Then
                If Not MailRec.Load(value, "#Index") Then
                    mMailIndex = 0
                    DisposeDA(MailRec)
                Else
                    Mail2_GetAtIndex(MailIndex, MailShip)
                End If
            End If

            LoadAddressFromMailRec()
        End Set
    End Property

    Private Sub LoadAddressFromMailRec()
        If MailRec Is Nothing Then ClearForm() : Exit Sub
        If AddressType = 0 Then
            'chkBusiness.checked = IIf(MailRec.Business, 1, 0)
            chkBusiness.Checked = IIf(MailRec.Business, True, False)
            txtFirstName.Text = MailRec.First
            txtLastName.Text = MailRec.Last
            txtAdd1.Text = MailRec.Address
            txtAdd2.Text = MailRec.AddAddress
            txtCityST.Text = MailRec.City
            txtZip.Text = MailRec.Zip
            txtPhone1.Text = DressAni(MailRec.Tele)
            txtPhone2.Text = DressAni(MailRec.Tele2)
            txtEmail.Text = MailRec.Email
        Else
            'chkBusiness.Value = 0
            chkBusiness.Checked = False
            txtFirstName.Text = MailShip.ShipToFirst
            txtLastName.Text = MailShip.ShipToLast
            txtAdd1.Text = MailShip.Address2
            txtCityST.Text = MailShip.City2
            txtZip.Text = MailShip.Zip2
            txtPhone1.Text = DressAni(MailShip.Tele3)
            txtPhone2.Text = ""
            txtEmail.Text = ""
        End If
    End Sub

    Private Sub ClearForm()
        'chkBusiness.Value = 0
        chkBusiness.Checked = False
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtAdd1.Text = ""
        txtAdd2.Text = ""
        txtCityST.Text = ""
        txtZip.Text = ""
        txtPhone1.Text = ""
        txtPhone2.Text = ""
        txtEmail.Text = ""
    End Sub

    Public Property AddressType() As Integer
        Get
            AddressType = mAtype
        End Get
        Set(value As Integer)
            mAtype = value
            UpdateControls()
        End Set
    End Property

    Private Sub frmCashRegisterAddress_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'If UnloadMode = vbFormControlMenu Then MailIndex = 0    ' X on top corner is same as cancel
        If UserClose = True Then
            If e.CloseReason = CloseReason.UserClosing Then
                MailIndex = 0
            End If
        End If
        DisposeDA(MailRec)
    End Sub

    Private Sub frmCashRegisterAddress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UpdateControls()
    End Sub

    Private Sub chkBusiness_CheckedChanged(sender As Object, e As EventArgs) Handles chkBusiness.CheckedChanged
        MailRec.Business = chkBusiness.Checked = True
        UpdateControls()
    End Sub

    Private Sub StoreChanges()
        If AddressType = 0 Then
            MailRec.Business = (chkBusiness.Checked = True)
            MailRec.First = Trim(txtFirstName.Text)
            MailRec.Last = Trim(txtLastName.Text)
            MailRec.Address = Trim(txtAdd1.Text)
            MailRec.AddAddress = Trim(txtAdd2.Text)
            MailRec.City = Trim(txtCityST.Text)
            MailRec.Zip = Trim(txtZip.Text)
            MailRec.Tele = CleanAni(txtPhone1.Text)
            MailRec.Tele2 = CleanAni(txtPhone2.Text)
            MailRec.Email = Trim(txtEmail.Text)
        Else
            MailShip.Index = MailRec.Index
            MailShip.ShipToFirst = Trim(txtFirstName.Text)
            MailShip.ShipToLast = Trim(txtLastName.Text)
            MailShip.Address2 = Trim(txtAdd1.Text)
            MailShip.City2 = Trim(txtCityST.Text)
            MailShip.Zip2 = Trim(txtZip.Text)
            MailShip.Tele3 = CleanAni(txtPhone1.Text)
        End If
    End Sub

    Private Sub UpdateControls()
        Dim ShipOnly As Boolean, Bus As Boolean
        Select Case AddressType
            Case 0 : ShipOnly = True
            Case 1 : ShipOnly = False
        End Select
        'Bus = chkBusiness.Value = 1
        Bus = chkBusiness.Checked = True

        chkBusiness.Visible = ShipOnly

        txtFirstName.Visible = Not Bus
        txtLastName.Left = IIf(Bus, txtFirstName.Left, 157)
        txtLastName.Width = IIf(Bus, 200, 100)

        txtAdd2.Visible = ShipOnly
        txtPhone2.Visible = ShipOnly
        lblEmail.Visible = ShipOnly
        txtEmail.Visible = ShipOnly

        'cmdShipTo.Caption = IIf(ShipOnly, "&Shipping Address >>", "<< Billing Addres&s")
        cmdShipTo.Text = IIf(ShipOnly, "&Shipping Address >>", "<< Billing Addres&s")

        LoadAddressFromMailRec()
    End Sub

    Private Sub cmdShipTo_Click(sender As Object, e As EventArgs) Handles cmdShipTo.Click
        StoreChanges()
        AddressType = IIf(AddressType = 0, 1, 0)
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        StoreChanges()
        If Not UpdateCustomerInfo() Then Exit Sub
        '<CT>
        UserClose = False
        '</CT>
        'Unload Me
        Me.Close()
        UserClose = True
    End Sub

    Private Sub txtFirstName_Enter(sender As Object, e As EventArgs) Handles txtFirstName.Enter
        SelectContents(txtFirstName)
    End Sub

    Private Sub txtLastName_Enter(sender As Object, e As EventArgs) Handles txtLastName.Enter
        SelectContents(txtLastName)
    End Sub

    Private Sub txtAdd1_Enter(sender As Object, e As EventArgs) Handles txtAdd1.Enter
        SelectContents(txtAdd1)
    End Sub

    Private Sub txtAdd2_Enter(sender As Object, e As EventArgs) Handles txtAdd2.Enter
        SelectContents(txtAdd2)
    End Sub

    Private Sub txtCityST_Enter(sender As Object, e As EventArgs) Handles txtCityST.Enter
        SelectContents(txtCityST)
    End Sub

    Private Sub txtZip_Enter(sender As Object, e As EventArgs) Handles txtZip.Enter
        SelectContents(txtZip)
    End Sub

    Private Sub txtPhone1_Enter(sender As Object, e As EventArgs) Handles txtPhone1.Enter
        SelectContents(txtPhone1)
    End Sub

    Private Sub txtPhone2_Enter(sender As Object, e As EventArgs) Handles txtPhone2.Enter
        SelectContents(txtPhone2)
    End Sub

    Private Sub txtEmail_Enter(sender As Object, e As EventArgs) Handles txtEmail.Enter
        SelectContents(txtEmail)
    End Sub

    Private Sub Modify(ByRef txt As TextBox)
        Dim A As Object, B As Object
        A = txt.SelectionStart
        B = txt.SelectionLength
        txt.Text = UCase(txt.Text)
        txt.SelectionStart = A
        txt.SelectionLength = B
    End Sub

    Private Sub txtFirstName_TextChanged(sender As Object, e As EventArgs) Handles txtFirstName.TextChanged
        Modify(txtFirstName)
    End Sub

    Private Sub txtLastName_TextChanged(sender As Object, e As EventArgs) Handles txtLastName.TextChanged
        Modify(txtLastName)
    End Sub

    Private Sub txtAdd1_TextChanged(sender As Object, e As EventArgs) Handles txtAdd1.TextChanged
        Modify(txtAdd1)
    End Sub

    Private Sub txtAdd2_TextChanged(sender As Object, e As EventArgs) Handles txtAdd2.TextChanged
        Modify(txtAdd2)
    End Sub

    Private Sub txtCityST_TextChanged(sender As Object, e As EventArgs) Handles txtCityST.TextChanged
        Modify(txtCityST)
    End Sub

    Private Sub txtZip_TextChanged(sender As Object, e As EventArgs) Handles txtZip.TextChanged
        Modify(txtZip)
    End Sub

    Private Sub txtPhone1_TextChanged(sender As Object, e As EventArgs) Handles txtPhone1.TextChanged
        FormatAniTextBox(txtPhone1)
    End Sub

    Private Sub txtPhone2_TextChanged(sender As Object, e As EventArgs) Handles txtPhone2.TextChanged
        FormatAniTextBox(txtPhone2)
    End Sub

    Private Function UpdateCustomerInfo() As Boolean
        Dim RS As ADODB.Recordset
        If MailIndex = 0 Then
            If MailRec.Last = "" Or MailRec.Tele = "" Then
                MessageBox.Show("You must enter a last name and a telephone number to enter this customer." & vbCrLf & "Either add supply the missing information or press ESC to exit this screen.", "No Customer Name or Telephone")
                Exit Function
            End If
        End If

        MailRec.Save()

        If Val(MailShip.Index) <> 0 And (MailShip.ShipToFirst <> "" Or MailShip.ShipToLast <> "" Or MailShip.Address2 <> "" Or MailShip.City2 <> "" Or MailShip.Zip2 <> "" Or MailShip.Tele3 <> "") Then
            Mail2_SetAtIndex(MailIndex, MailShip)
        End If

        MailIndex = MailRec.Index ' mostly for new records, but shouldn't hurt..
        UpdateCustomerInfo = True
    End Function

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        MailIndex = 0
        'Unload Me
        Me.Close()
    End Sub
End Class