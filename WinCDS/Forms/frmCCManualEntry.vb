Public Class frmCCManualEntry
    Private Const O As String = "[OMIT]"
    Private mCancelled As Boolean
    Private mSwipe As String

    Public Function GetManualCCEntry(Optional ByRef CCNumber As String = "", Optional ByRef ExpDate As String = "0", Optional ByRef CardHolderName As String = "O", Optional ByRef CVV2 As String = "O", Optional ByRef ZipCode As String = "O", Optional ByRef Swipe As String = "") As Boolean
        SetButtonImage(cmdOK)
        SetButtonImage(cmdCancel)

        mCancelled = False
        mSwipe = ""

        cmdOK.Top = 3000

        If ZipCode = O Then
            lblZipCode.Visible = False
            txtZipCode.Visible = False
        Else
            txtZipCode.Text = DetectCustomerZipCode()
            txtCardHolderName.Text = DetectCustomerName()
        End If
        txtZipCode.Tag = IIf(ZipCode = O, "-", "")
        If CVV2 = O Then
            lblCVV2.Visible = False
            txtCVV2.Visible = False
        End If
        txtCVV2.Tag = IIf(CVV2 = O, "-", "")

        If ZipCode = O And CVV2 = O Then
            cmdOK.Top = cmdOK.Top - 720
        End If

        If CardHolderName = O Then
            lblCardHolderName.Visible = False
            txtCardHolderName.Visible = False
            cmdOK.Top = cmdOK.Top - 720
        End If
        txtCardHolderName.Tag = IIf(CardHolderName = O, "-", "")

        If ExpDate = O Then
            lblExpDate.Visible = False
            txtExpDate.Visible = False
            cmdOK.Top = cmdOK.Top - 720
        End If
        txtExpDate.Tag = IIf(ExpDate = O, "-", "")

        cmdCancel.Top = cmdOK.Top
        fraCC.Height = cmdOK.Top + cmdOK.Height + 120
        'Height = Height - ScaleHeight + fraCC.Top + fraCC.Height + 120
        Height = Height - Me.ClientSize.Height + fraCC.Top + fraCC.Height + 120

        'Show vbModal
        ShowDialog()

        GetManualCCEntry = Not mCancelled

        If GetManualCCEntry Then
            CCNumber = txtCCNumber.Text
            ExpDate = txtExpDate.Text
            CardHolderName = txtCardHolderName.Text
            CVV2 = txtCVV2.Text
            ZipCode = txtZipCode.Text
            Swipe = mSwipe
        Else
            CCNumber = ""
            ExpDate = ""
            CardHolderName = ""
            CVV2 = ""
            ZipCode = ""
            Swipe = ""
        End If
    End Function

End Class