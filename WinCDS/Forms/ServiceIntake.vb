Public Class ServiceIntake
    Private mSONO As Integer
    Private PartsOrderNumber As Integer
    Private mStore As Integer
    Private mVendor As String
    Private CBType As Integer
    Private Amount As Decimal
    Private InvoiceNo As String
    Private Updating As Boolean   ' internal semaphore

    Public Property ServiceOrderNumber() As Integer
        Get
            ServiceOrderNumber = mSONO
        End Get
        Set(value As Integer)
            mSONO = value
            txtServiceOrderNumber.Text = value
        End Set
    End Property

    Public Property Store() As Integer
        Get
            If mStore = 0 Then mStore = StoresSld
            Store = mStore
        End Get
        Set(value As Integer)
            If value = 0 Then value = StoresSld
            mStore = value
            txtLocation.Text = "Loc " & Store
        End Set
    End Property

    Public Property Vendor() As String
        Get
            Vendor = mVendor
        End Get
        Set(value As String)
            mVendor = value
            txtVendor.Text = Vendor
        End Set
    End Property

    Public Sub InitForm(ByVal nServiceOrderNumber As Integer, ByVal PartsOrderID As Integer, ByVal nCBType As Integer, ByVal CBAmount As String, ByVal nInvoiceNo As String, ByVal nStore As Integer, ByVal nVendor As String)
        If nServiceOrderNumber <> 0 Then
            txtServiceOrderNumber.Visible = True
            lblServiceOrderNumber.Visible = True
            ServiceOrderNumber = nServiceOrderNumber
        Else
            txtServiceOrderNumber.Visible = False
            lblServiceOrderNumber.Visible = False
            ServiceOrderNumber = 0
        End If

        PartsOrderNumber = PartsOrderID
        Store = nStore
        Vendor = nVendor

        CBType = nCBType
        Amount = CBAmount
        InvoiceNo = nInvoiceNo

        Select Case CBType
            Case 0
                txtMode.Text = "Charging Back " & FormatCurrency(Amount)
            Case 1
                txtMode.Text = "Deducting " & FormatCurrency(Amount) & " From Invoice #" & InvoiceNo
            Case 2
                txtMode.Text = "Requesting a Credit of " & FormatCurrency(Amount)
            Case Else
                MessageBox.Show("Unknown letter type!!", "Stop!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                'Unload Me
                Me.Close()
        End Select

        LoadImages(PartsOrderID)
    End Sub

    Private Sub LoadImages(ByVal No As String)
        Dim RS As ADODB.Recordset
        cboImage.Items.Clear()
        cboImage.Items.Add("No Image")
        cboImage.Enabled = False
        RS = GetRecordsetBySQL("SELECT PictureID, Caption FROM Pictures WHERE PictureType=2 AND PictureRef='" & ProtectSQL(No) & "' ORDER BY PictureID", , GetDatabaseAtLocation())
        Do While Not RS.EOF
            cboImage.Enabled = True
            'cboImage.AddItem IfNullThenNilString(RS("Caption"))
            'cboImage.itemData(cboImage.NewIndex) = IfNullThenZero(RS("PictureID"))
            cboImage.Items.Add(New ItemDataClass(IfNullThenNilString(RS("Caption").Value), IfNullThenZero(RS("PictureID").Value)))
            RS.MoveNext
        Loop
        'cboImage.ListIndex = IIf(cboImage.Enabled, 1, 0)
        cboImage.SelectedIndex = IIf(cboImage.Enabled, 1, 0)
    End Sub

End Class