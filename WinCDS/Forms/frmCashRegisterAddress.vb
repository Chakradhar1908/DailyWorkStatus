Public Class frmCashRegisterAddress
    Private mMailIndex As Integer, MailRec As clsMailRec, MailShip As MailNew2
    Private mAtype As Integer

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
        txtLastName.Left = IIf(Bus, txtFirstName.Left, 2040)
        txtLastName.Width = IIf(Bus, 2415, 1215)

        txtAdd2.Visible = ShipOnly
        txtPhone2.Visible = ShipOnly
        lblEmail.Visible = ShipOnly
        txtEmail.Visible = ShipOnly

        'cmdShipTo.Caption = IIf(ShipOnly, "&Shipping Address >>", "<< Billing Addres&s")
        cmdShipTo.Text = IIf(ShipOnly, "&Shipping Address >>", "<< Billing Addres&s")

        LoadAddressFromMailRec()
    End Sub

End Class