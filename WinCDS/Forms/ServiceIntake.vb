Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
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
            RS.MoveNext()
        Loop
        'cboImage.ListIndex = IIf(cboImage.Enabled, 1, 0)
        cboImage.SelectedIndex = IIf(cboImage.Enabled, 1, 0)
    End Sub

    Private Sub cboImage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboImage.SelectedIndexChanged
        Dim X As Integer

        'X = cboImage.itemData(cboImage.ListIndex)
        Try
            X = CType(cboImage.Items(cboImage.SelectedIndex), ItemDataClass).ItemData
            If X > 0 Then
                'datPicture.DatabaseName = GetDatabaseAtLocation()
                'datPicture.RecordSource = "SELECT Picture FROM Pictures WHERE PictureID=" & X
                'datPicture.Refresh
                'datPicture.DataBase.Close

                Dim Rs As ADODB.Recordset
                Rs = GetRecordsetBySQL("SELECT Picture FROM Pictures WHERE PictureID=" & X,, GetDatabaseAtLocation)
                Rs.Close()
            End If
        Catch ic As System.InvalidCastException
            'If no items in cboImage except "No Image" item, this error will thrown because of using ItemDataClass in the top line.
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ServiceIntake_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdPrint, 2)
        SetButtonImage(cmdCancel, 3)
        SetButtonImage(cmdEditTemplate, 25)
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        ServiceParts.PrintedChargeBack(False)
        'Unload Me
        Me.Close()
    End Sub

    Private Sub cmdEditTemplate_Click(sender As Object, e As EventArgs) Handles cmdEditTemplate.Click
        'frmEmailEdit.Show vbModal
        frmEmailEdit.ShowDialog()
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Dim OK As String, Em As Boolean
        If optDelivery0.Checked = True Then
            OK = True
            Em = False
            PrintChargeBackLetter(CBType, Store, Amount, InvoiceNo, imgPicture)
        Else
            OK = frmEmail.EmailChargeBackLetter(PartsOrderNumber, CBType, Store, Amount, InvoiceNo)
            Em = True
        End If
        ServiceParts.PrintedChargeBack(True, Em, CBType)
        'Unload Me
        Me.Close()
    End Sub

    Private Sub PrintChargeBackLetter(ByVal LetterType As Integer, ByVal StoreNum As Integer, ByVal Amount As Decimal, ByVal InvoiceNo As String, ByRef MyPic As PictureBox)
        Dim P As Printer, Oper As String
        Dim Ex As String

        OutputObject = Printer
        OutputToPrinter = True

        If PartsOrderNumber = 0 Then
            MessageBox.Show("No Parts Order Number Available!", "Insufficient Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Oper = ChargeBackLetterOperationDesc(LetterType, Amount, InvoiceNo)
        P = Printer

        P.FontName = "Arial New"
        P.FontSize = 16 : PrintAligned("CREDIT DEPARTMENT", VBRUN.AlignmentConstants.vbCenter, 0, 100, True)
        P.FontSize = 12
        P.DrawWidth = 2

        PrintAutoMailingLetterHeader(StoreSettings.Name, StoreSettings.Address, StoreSettings.City, StoreSettings.Phone, ServiceParts.txtVendorName.Text, ServiceParts.txtVendorAddress.Text, ServiceParts.txtVendorCity.Text, ServiceParts.txtVendorTele.Text)

        P.FontBold = True
        If Len(InvoiceNo) > 0 Or mSONO <> 0 Then
            Ex = ""
            Ex = Ex & " ("
            If mSONO <> 0 Then Ex = Ex & "Service Order #" & mSONO
            If Len(InvoiceNo) > 0 Then
                Ex = Ex & IIf(Len(Ex) > 2, ", ", "")
                Ex = Ex & "Invoice #" & InvoiceNo
            End If
            Ex = Ex & ")"
        End If
        P.Print("RE: Service Parts Order #" & PartsOrderNumber & Ex)
        P.FontBold = False

        P.Print("")

        P.Print("Attention: Accounts Receivable Department")

        P.Print("")

        P.Print("Dear Sir:")

        P.Print("")

        P.Print("", "As per the attached service order, we are " & Oper & ".")

        P.Print("")
        'If MyPic.Image <> 0 Then
        If Not IsNothing(MyPic.Image) Then
            MaintainPictureRatio(MyPic, 7000, 5000, True)
            Printer.PaintPicture(MyPic.Image, (P.ScaleWidth - MyPic.Width) / 2, P.CurrentY, MyPic.Width, MyPic.Height)
            P.CurrentY = P.CurrentY + 5000
        End If

        P.Print("")

        P.Print("Thank You,")
        P.Print("")
        '    P.Print "", SI.ContactName  ' it'd be nice if we had something like this...
        P.Print("", StoreSettings.Name)
        P.EndDoc()
    End Sub

End Class