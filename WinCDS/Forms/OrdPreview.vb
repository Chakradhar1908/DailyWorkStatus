Public Class OrdPreview
    Public RN As Integer
    Private InfoBase As Integer

    Private mShowingLabels As Boolean

    Private Const H_FraInfoShowInfo As Integer = 2055
    Private Const H_FraInfoHideInfo As Integer = 855

    Public Function LoadItemByRN(ByVal nRN As Integer) As Boolean
        Dim DataObj As New CInvRec, I As Integer

        If Not ItemIsActive(nRN) Then Exit Function

        On Error Resume Next
        RN = nRN

        If DataObj.Load(CStr(RN), "#Rn") Then

            OnSale.Text = Format(DataObj.OnSale, "$###,##0.00")
            Desc.Text = ""
            Desc.Text = DataObj.Desc
            Comments.Text = ""
            Comments.Text = DataObj.Comments

            txtStyle.Text = DataObj.Style

            For I = 1 To Setup_MaxStores
                'lblStore(I).ToolTipText = StoreSettings(I).Address
                ToolTip1.SetToolTip(lblStore1, StoreSettings(1).Address)
                ToolTip1.SetToolTip(lblStore2, StoreSettings(2).Address)
                ToolTip1.SetToolTip(lblStore3, StoreSettings(3).Address)
                ToolTip1.SetToolTip(lblStore4, StoreSettings(4).Address)
                ToolTip1.SetToolTip(lblStore5, StoreSettings(5).Address)
                ToolTip1.SetToolTip(lblStore6, StoreSettings(6).Address)
                ToolTip1.SetToolTip(lblStore7, StoreSettings(7).Address)
                ToolTip1.SetToolTip(lblStore8, StoreSettings(8).Address)

                'lblStock(I).ToolTipText = StoreSettings(I).Address
                ToolTip1.SetToolTip(lblStock1, StoreSettings(1).Address)
                ToolTip1.SetToolTip(lblStock2, StoreSettings(2).Address)
                ToolTip1.SetToolTip(lblStock3, StoreSettings(3).Address)
                ToolTip1.SetToolTip(lblStock4, StoreSettings(4).Address)
                ToolTip1.SetToolTip(lblStock5, StoreSettings(5).Address)
                ToolTip1.SetToolTip(lblStock6, StoreSettings(6).Address)
                ToolTip1.SetToolTip(lblStock7, StoreSettings(7).Address)
                ToolTip1.SetToolTip(lblStock8, StoreSettings(8).Address)


                'lblStock(I) = DataObj.QueryOnOrder(I)
                lblStock1.Text = DataObj.QueryOnOrder(1)
                lblStock2.Text = DataObj.QueryOnOrder(2)
                lblStock3.Text = DataObj.QueryOnOrder(3)
                lblStock4.Text = DataObj.QueryOnOrder(4)
                lblStock5.Text = DataObj.QueryOnOrder(5)
                lblStock6.Text = DataObj.QueryOnOrder(6)
                lblStock7.Text = DataObj.QueryOnOrder(7)
                lblStock8.Text = DataObj.QueryOnOrder(8)

                'lblOrder(I) = DataObj.QueryOnOrder(I)
                lblOrder1.Text = DataObj.QueryOnOrder(1)
                lblOrder2.Text = DataObj.QueryOnOrder(2)
                lblOrder3.Text = DataObj.QueryOnOrder(3)
                lblOrder4.Text = DataObj.QueryOnOrder(4)
                lblOrder5.Text = DataObj.QueryOnOrder(5)
                lblOrder6.Text = DataObj.QueryOnOrder(6)
                lblOrder7.Text = DataObj.QueryOnOrder(7)
                lblOrder8.Text = DataObj.QueryOnOrder(8)

            Next
        End If

        DisposeDA(DataObj)

        LoadPix()
    End Function

    Private Function ItemIsActive(ByVal RN As Integer) As Boolean
        Dim RS As ADODB.Recordset
        RS = GetRecordsetBySQL("SELECT * FROM Search WHERE RN = " & RN, , GetDatabaseInventory)
        ItemIsActive = Not RS.EOF
        RS = Nothing
    End Function

    Private Sub LoadPix()
        On Error GoTo HandleErr
        'imgPicture.Picture = LoadPictureStd(ItemPXByRN(RN))
        imgPicture.Image = LoadPictureStd(ItemPXByRN(RN))
        'Form_Resize
        OrdPreview_Resize(Me, New EventArgs)
        Exit Sub

HandleErr:
        'imgPicture.Picture = LoadPictureStd("")
        imgPicture.Image = LoadPictureStd("")
        Err.Clear()
        Resume Next
    End Sub

    Private Sub OrdPreview_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        Dim MaxPicX As Integer, MaxPicY As Integer
        'MaxPicX = ScaleWidth - imgPicture.Left - 240
        MaxPicX = Me.ClientSize.Width - imgPicture.Left - 240
        'MaxPicY = ScaleHeight - imgPicture.Top - fraInformation.Height - 100 - 240
        MaxPicY = Me.ClientSize.Height - imgPicture.Top - fraInformation.Height - 100 - 240

        MaintainPictureRatio(imgPicture, MaxPicX, MaxPicY)

        StoreName.Width = imgPicture.Width
        fraInformation.Width = imgPicture.Width
        Comments.Width = IIf(fraInformation.Width - 240 > 0, fraInformation.Width - 240, 10)
        Desc.Width = Comments.Width

        fraInformation.Top = imgPicture.Top + imgPicture.Height + 100
    End Sub
End Class
