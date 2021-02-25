Imports System.ComponentModel
Public Class OrdPreview
    Public RN As Integer
    Private InfoBase As Integer
    Private mShowingLabels As Boolean
    Private Const H_FraInfoShowInfo As Integer = 120
    Private Const H_FraInfoHideInfo As Integer = 60

    Public Sub Navigate(ByVal Forward As Boolean, ByVal Absolute As Boolean)
        Dim SQL As String, Rest As String, RS As ADODB.Recordset
        Dim GetDeptValue As Integer

        Rest = " ASC"
        SQL = "SELECT"
        If Absolute Then
            SQL = SQL & " TOP 1"
            If Forward Then
                Rest = " DESC"
            End If
        ElseIf Not Forward Then
            Rest = " DESC"
        End If
        SQL = SQL & " *"
        ' bfh20050616 - limited this stock preview to non-deleted items..
        '  SQL = SQL & " FROM [2Data] WHERE (TRUE=TRUE)"
        SQL = SQL & " FROM [2Data] WHERE Rn IN (SELECT Rn From Search) AND (TRUE=TRUE)"
        If GetVendor() <> "" Then SQL = SQL & " AND (Vendor = """ & ProtectSQL(GetVendor) & """)"
        'If GetDept() <> 0 Then SQL = SQL & " AND (Dept = '" & GetDept() & "')"
        If GetDeptValue = GetDept() <> 0 Then SQL = SQL & " AND (Dept = '" & GetDeptValue & "')"
        If Not Absolute Then SQL = SQL & " AND (Style " & IIf(Forward, ">", "<") & " '" & txtStyle.Text & "')"
        SQL = SQL & " ORDER BY STYLE"
        SQL = SQL & Rest

        RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory)
        If Not RS.EOF Then
            LoadItemByRN(RS("RN").Value)
        End If
        RS.Close()
        RS = Nothing
    End Sub

    Private Function GetVendor() As String
        GetVendor = cboVendor.Text
    End Function

    Private Function GetDept() As Integer
        On Error Resume Next
        Dim I As Integer
        For I = 0 To cboDepartment.Items.Count - 1
            'If cboDepartment.Text = cboDepartment.List(I) Then GetDept = cboDepartment.itemData(I) : Exit For
            If cboDepartment.Text = cboDepartment.Items(cboDepartment.SelectedIndex).ToString Then GetDept = CType(cboDepartment.Items(cboDepartment.SelectedIndex), ItemDataClass).ItemData : Exit For
        Next
    End Function

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
        MaxPicX = Me.ClientSize.Width - imgPicture.Left - 24
        'MaxPicY = ScaleHeight - imgPicture.Top - fraInformation.Height - 100 - 240
        MaxPicY = Me.ClientSize.Height - imgPicture.Top - fraInformation.Height - 10 - 24

        MaintainPictureRatio(imgPicture, MaxPicX, MaxPicY)

        StoreName.Width = imgPicture.Width
        fraInformation.Width = imgPicture.Width
        Comments.Width = IIf(fraInformation.Width - 24 > 0, fraInformation.Width - 24, 1)
        Desc.Width = Comments.Width

        fraInformation.Top = imgPicture.Top + imgPicture.Height + 10
    End Sub

    Private Sub fraControl_Enter(sender As Object, e As EventArgs) Handles fraControl.Enter

    End Sub

    Private Sub cboDepartment_Leave(sender As Object, e As EventArgs) Handles cboDepartment.Leave
        txtStyle.Text = ""
        cboVendor.Text = ""
        Navigate(False, True)
    End Sub

    Private Sub cboDepartment_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDepartment.SelectedIndexChanged
        txtStyle.Text = ""
        cboVendor.Text = ""
        Navigate(False, True)
    End Sub

    Private Sub cboVendor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboVendor.SelectedIndexChanged
        txtStyle.Text = ""
        cboDepartment.Text = ""
        Navigate(False, True)
    End Sub

    Private Sub cboVendor_Leave(sender As Object, e As EventArgs) Handles cboVendor.Leave
        txtStyle.Text = ""
        cboDepartment.Text = ""
        Navigate(False, True)
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        'Unload Me
        Me.Close()
        modProgramState.Order = ""
        MainMenu.Show()
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        InvCkStyle.Owned = True
        InvCkStyle.Width = 3345
        InvCkStyle.lstStyles.Visible = False
        'InvCkStyle.Show vbModal, Me
        InvCkStyle.ShowDialog(Me)
        LoadPix()
    End Sub

    Private Sub cmdShow_Click(sender As Object, e As EventArgs) Handles cmdShow.Click
        ShowLabels(Not mShowingLabels)
    End Sub

    Private Sub ShowLabels(ByVal Vis As Boolean, Optional ByVal DoInit As Boolean = False)
        Dim I As Integer
        mShowingLabels = Vis
        fraInformation.Height = IIf(Vis, H_FraInfoShowInfo, H_FraInfoHideInfo)
        'Form_Resize
        OrdPreview_Resize(Me, New EventArgs)

        If Not DoInit Then Exit Sub

        On Error Resume Next
        For I = 2 To Setup_MaxStores
            'Load lblStore(I)
            'lblStore(I).Caption = "Loc" & I
            'Load lblStock(I)
            'lblStock(I).Caption = ""
            'lblStock(I).Width = lblStock(1).Width
            'lblStock(I).Height = lblStock(1).Height
            'lblStock(I).Alignment = lblStock(1).Alignment
            'Load lblOrder(I)
            'lblOrder(I).Caption = ""
            'lblOrder(I).Width = lblOrder(1).Width
            'lblOrder(I).Height = lblOrder(1).Height
            'lblOrder(I).Alignment = lblOrder(1).Alignment

            Dim lblStore As Label
            Dim lblStock As Label
            Dim lblOrder As Label
            Select Case I
                Case 1
                    lblStore1.Text = "Loc" & I
                    lblStock1.Text = ""
                    lblStock1.Width = lblStock1.Width
                    lblStock1.Height = lblStock1.Height
                    lblStock1.TextAlign = lblStock1.TextAlign
                    lblOrder1.Text = ""
                    lblOrder1.Width = lblOrder1.Width
                    lblOrder1.Height = lblOrder1.Height
                    lblOrder1.TextAlign = lblOrder1.TextAlign
                Case 2
                    lblStore2.Text = "Loc" & I
                    lblStock2.Text = ""
                    lblStock2.Width = lblStock1.Width
                    lblStock2.Height = lblStock1.Height
                    lblStock2.TextAlign = lblStock1.TextAlign
                    lblOrder2.Text = ""
                    lblOrder2.Width = lblOrder1.Width
                    lblOrder2.Height = lblOrder1.Height
                    lblOrder2.TextAlign = lblOrder1.TextAlign
                Case 3
                    lblStore3.Text = "Loc" & I
                    lblStock3.Text = ""
                    lblStock3.Width = lblStock1.Width
                    lblStock3.Height = lblStock1.Height
                    lblStock3.TextAlign = lblStock1.TextAlign
                    lblOrder3.Text = ""
                    lblOrder3.Width = lblOrder1.Width
                    lblOrder3.Height = lblOrder1.Height
                    lblOrder3.TextAlign = lblOrder1.TextAlign
                Case 4
                    lblStore4.Text = "Loc" & I
                    lblStock4.Text = ""
                    lblStock4.Width = lblStock1.Width
                    lblStock4.Height = lblStock1.Height
                    lblStock4.TextAlign = lblStock1.TextAlign
                    lblOrder4.Text = ""
                    lblOrder4.Width = lblOrder1.Width
                    lblOrder4.Height = lblOrder1.Height
                    lblOrder4.TextAlign = lblOrder1.TextAlign
                Case 5
                    lblStore5.Text = "Loc" & I
                    lblStock5.Text = ""
                    lblStock5.Width = lblStock1.Width
                    lblStock5.Height = lblStock1.Height
                    lblStock5.TextAlign = lblStock1.TextAlign
                    lblOrder5.Text = ""
                    lblOrder5.Width = lblOrder1.Width
                    lblOrder5.Height = lblOrder1.Height
                    lblOrder5.TextAlign = lblOrder1.TextAlign
                Case 6
                    lblStore6.Text = "Loc" & I
                    lblStock6.Text = ""
                    lblStock6.Width = lblStock1.Width
                    lblStock6.Height = lblStock1.Height
                    lblStock6.TextAlign = lblStock1.TextAlign
                    lblOrder6.Text = ""
                    lblOrder6.Width = lblOrder1.Width
                    lblOrder6.Height = lblOrder1.Height
                    lblOrder6.TextAlign = lblOrder1.TextAlign
                Case 7
                    lblStore7.Text = "Loc" & I
                    lblStock7.Text = ""
                    lblStock7.Width = lblStock1.Width
                    lblStock7.Height = lblStock1.Height
                    lblStock7.TextAlign = lblStock1.TextAlign
                    lblOrder7.Text = ""
                    lblOrder7.Width = lblOrder1.Width
                    lblOrder7.Height = lblOrder1.Height
                    lblOrder7.TextAlign = lblOrder1.TextAlign
                Case 8
                    lblStore8.Text = "Loc" & I
                    lblStock8.Text = ""
                    lblStock8.Width = lblStock1.Width
                    lblStock8.Height = lblStock1.Height
                    lblStock8.TextAlign = lblStock1.TextAlign
                    lblOrder8.Text = ""
                    lblOrder8.Width = lblOrder1.Width
                    lblOrder8.Height = lblOrder1.Height
                    lblOrder8.TextAlign = lblOrder1.TextAlign
                Case 9
                    lblStore = New Label()
                    lblStore.Name = "lblStore9"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock9"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder9"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 10
                    lblStore = New Label()
                    lblStore.Name = "lblStore10"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock10"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder10"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 11
                    lblStore = New Label()
                    lblStore.Name = "lblStore11"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock11"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder11"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 12
                    lblStore = New Label()
                    lblStore.Name = "lblStore12"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock12"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder12"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 13
                    lblStore = New Label()
                    lblStore.Name = "lblStore13"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock13"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder13"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 14
                    lblStore = New Label()
                    lblStore.Name = "lblStore14"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock14"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder14"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 15
                    lblStore = New Label()
                    lblStore.Name = "lblStore15"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock15"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder15"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 16
                    lblStore = New Label()
                    lblStore.Name = "lblStore16"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock16"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder16"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 17
                    lblStore = New Label()
                    lblStore.Name = "lblStore17"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock17"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder17"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 18
                    lblStore = New Label()
                    lblStore.Name = "lblStore18"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock18"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder18"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 19
                    lblStore = New Label()
                    lblStore.Name = "lblStore19"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock19"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder19"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 20
                    lblStore = New Label()
                    lblStore.Name = "lblStore20"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock20"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder20"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 21
                    lblStore = New Label()
                    lblStore.Name = "lblStore21"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock21"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder21"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 22
                    lblStore = New Label()
                    lblStore.Name = "lblStore22"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock22"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder22"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 23
                    lblStore = New Label()
                    lblStore.Name = "lblStore23"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock23"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder23"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 24
                    lblStore = New Label()
                    lblStore.Name = "lblStore24"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock24"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder24"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 25
                    lblStore = New Label()
                    lblStore.Name = "lblStore25"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock25"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder25"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 26
                    lblStore = New Label()
                    lblStore.Name = "lblStore26"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock26"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder26"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 27
                    lblStore = New Label()
                    lblStore.Name = "lblStore27"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock27"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder27"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 28
                    lblStore = New Label()
                    lblStore.Name = "lblStore28"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock28"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder28"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 29
                    lblStore = New Label()
                    lblStore.Name = "lblStore29"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock29"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder29"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 30
                    lblStore = New Label()
                    lblStore.Name = "lblStore30"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock30"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder30"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 31
                    lblStore = New Label()
                    lblStore.Name = "lblStore31"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock31"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder31"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
                Case 32
                    lblStore = New Label()
                    lblStore.Name = "lblStore32"
                    lblStore.Text = "Loc" & I
                    Me.Controls.Add(lblStore)
                    lblStock = New Label
                    lblStock.Name = "lblStock32"
                    lblStock.Text = ""
                    lblStock.Width = lblStock1.Width
                    lblStock.Height = lblStock1.Height
                    lblStock.TextAlign = lblStock1.TextAlign
                    Me.Controls.Add(lblStock)
                    lblOrder = New Label
                    lblOrder.Name = "lblOrder32"
                    lblOrder.Text = ""
                    lblOrder.Width = lblOrder1.Width
                    lblOrder.Height = lblOrder1.Height
                    lblOrder.TextAlign = lblOrder1.TextAlign
                    Me.Controls.Add(lblOrder)
            End Select
        Next

        InfoBase = 1
        ArrangeInfo()
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Dim vImg As Image, Nx As Integer, nY As Integer
        If txtStyle.Text = "" Then
            MessageBox.Show("Please view an item before attempting to print!", "Not Viewing An Item!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        vImg = imgPicture.Image
        OutputToPrinter = True

        If imgPicture.Image.Width = 0 And imgPicture.Image.Height = 0 Then
            MessageBox.Show("There is not an image available for this item.", "No Image Available", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Nx = Printer.ScaleWidth - 3500 ' - vImg.Width
        nY = imgPicture.Image.Height / imgPicture.Image.Width * Nx

        '    nX = 3.5 * nX
        '    nY = 3.5 * nY

        Printer.FontSize = 16
        Printer.FontName = "Arial"
        PrintAligned(StoreName.Text, VBRUN.AlignmentConstants.vbCenter, , , True)

        If imgPicture.Image.Width > 0 And imgPicture.Image.Height > 0 Then
            Printer.PaintPicture(vImg, 3000, 1000, Nx, nY)
        End If

        'Printer.Line(240, 600)-(2415, 2000), , B
        Printer.Line(240, 600, 2415, 2000, , True)
        PrintAligned("Sale Price", , 400, 800, True)
        PrintAligned(OnSale.Text, , 400, , True)

        Printer.CurrentY = 1000 + nY + 1000
        ' bfh20051026 - style removed from hardcopy
        '    PrintAligned txtStyle, vbCenter, , 1000 + nY + 1000
        PrintAligned(Desc.Text, VBRUN.AlignmentConstants.vbCenter)
        PrintAligned(Comments.Text, VBRUN.AlignmentConstants.vbCenter)

        Printer.EndDoc()
    End Sub

    Private Sub cmdNextStores_Click(sender As Object, e As EventArgs) Handles cmdNextStores.Click
        InfoBase = 9
        ArrangeInfo()
    End Sub

    Private Sub ArrangeInfo()
        Dim Max As Integer, N As Integer, I As Integer

        Max = LicensedNoOfStores()

        cmdNextStores.Visible = LicensedNoOfStores() > 8
        cmdPrevStores.Visible = LicensedNoOfStores() > 8

        For I = 1 To Setup_MaxStores
            'lblStore(I).Visible = False
            'lblStock(I).Visible = False
            'lblOrder(I).Visible = False
            Select Case I
                Case 1
                    lblStore1.Visible = False
                    lblStock1.Visible = False
                    lblOrder1.Visible = False
                Case 2
                    lblStore2.Visible = False
                    lblStock2.Visible = False
                    lblOrder2.Visible = False
                Case 3
                    lblStore3.Visible = False
                    lblStock3.Visible = False
                    lblOrder3.Visible = False
                Case 4
                    lblStore4.Visible = False
                    lblStock4.Visible = False
                    lblOrder4.Visible = False
                Case 5
                    lblStore5.Visible = False
                    lblStock5.Visible = False
                    lblOrder5.Visible = False
                Case 6
                    lblStore6.Visible = False
                    lblStock6.Visible = False
                    lblOrder6.Visible = False
                Case 7
                    lblStore7.Visible = False
                    lblStock7.Visible = False
                    lblOrder7.Visible = False
                Case 8
                    lblStore8.Visible = False
                    lblStock8.Visible = False
                    lblOrder8.Visible = False
                Case 9
                    StoreVisible("lblStore9")
                    StockVisible("lblStock9")
                    OrderVisible("lblOrder9")
                Case 10
                    StoreVisible("lblStore10")
                    StockVisible("lblStock10")
                    OrderVisible("lblOrder10")
                Case 11
                    StoreVisible("lblStore11")
                    StockVisible("lblStock11")
                    OrderVisible("lblOrder11")
                Case 12
                    StoreVisible("lblStore12")
                    StockVisible("lblStock12")
                    OrderVisible("lblOrder12")
                Case 13
                    StoreVisible("lblStore13")
                    StockVisible("lblStock13")
                    OrderVisible("lblOrder13")
                Case 14
                    StoreVisible("lblStore14")
                    StockVisible("lblStock14")
                    OrderVisible("lblOrder14")
                Case 15
                    StoreVisible("lblStore15")
                    StockVisible("lblStock15")
                    OrderVisible("lblOrder15")
                Case 16
                    StoreVisible("lblStore16")
                    StockVisible("lblStock16")
                    OrderVisible("lblOrder16")
                Case 17
                    StoreVisible("lblStore17")
                    StockVisible("lblStock17")
                    OrderVisible("lblOrder17")
                Case 18
                    StoreVisible("lblStore18")
                    StockVisible("lblStock18")
                    OrderVisible("lblOrder18")
                Case 19
                    StoreVisible("lblStore19")
                    StockVisible("lblStock19")
                    OrderVisible("lblOrder19")
                Case 20
                    StoreVisible("lblStore20")
                    StockVisible("lblStock20")
                    OrderVisible("lblOrder20")
                Case 21
                    StoreVisible("lblStore21")
                    StockVisible("lblStock21")
                    OrderVisible("lblOrder21")
                Case 22
                    StoreVisible("lblStore22")
                    StockVisible("lblStock22")
                    OrderVisible("lblOrder22")
                Case 23
                    StoreVisible("lblStore23")
                    StockVisible("lblStock23")
                    OrderVisible("lblOrder23")
                Case 24
                    StoreVisible("lblStore24")
                    StockVisible("lblStock24")
                    OrderVisible("lblOrder24")
                Case 25
                    StoreVisible("lblStore25")
                    StockVisible("lblStock25")
                    OrderVisible("lblOrder25")
                Case 26
                    StoreVisible("lblStore26")
                    StockVisible("lblStock26")
                    OrderVisible("lblOrder26")
                Case 27
                    StoreVisible("lblStore27")
                    StockVisible("lblStock27")
                    OrderVisible("lblOrder27")
                Case 28
                    StoreVisible("lblStore28")
                    StockVisible("lblStock28")
                    OrderVisible("lblOrder28")
                Case 29
                    StoreVisible("lblStore29")
                    StockVisible("lblStock29")
                    OrderVisible("lblOrder29")
                Case 30
                    StoreVisible("lblStore30")
                    StockVisible("lblStock30")
                    OrderVisible("lblOrder30")
                Case 31
                    StoreVisible("lblStore31")
                    StockVisible("lblStock31")
                    OrderVisible("lblOrder31")
                Case 32
                    StoreVisible("lblStore32")
                    StockVisible("lblStock32")
                    OrderVisible("lblOrder32")
            End Select
        Next

        For I = 1 To 8     ' we only display 8 at a time
            Select Case InfoBase
                Case 9
                    N = I + 8
                Case Else  ' 1 or whatever else
                    N = I
            End Select
            If N <= Max Then
                'lblStore(N).Visible = True
                'lblStore(N).Top = lblStore(1).Top
                'lblStore(N).Left = 960 + 600 * (I - 1)
                'lblStock(N).Visible = True
                'lblStock(N).Top = lblStock(1).Top
                'lblStock(N).Left = 1080 + 600 * (I - 1)
                'lblOrder(N).Visible = True
                'lblOrder(N).Top = lblOrder(1).Top
                'lblOrder(N).Left = 1080 + 600 * (I - 1)

                Select Case N
                    Case 1
                        lblStore1.Visible = True
                        lblStore1.Top = lblStore1.Top
                        lblStore1.Left = 90 + 60 * (1 - 1)
                        lblStock1.Visible = True
                        lblStock1.Top = lblStock1.Top
                        lblStock1.Left = 100 + 60 * (1 - 1)
                        lblOrder1.Visible = True
                        lblOrder1.Top = lblOrder1.Top
                        lblOrder1.Left = 100 + 50 * (1 - 1)
                    Case 2
                        lblStore2.Visible = True
                        lblStore2.Top = lblStore1.Top
                        'lblStore2.Left = 90 + 60 * (2 - 1)
                        lblStore2.Left = 90 + 50 * (2 - 1)
                        lblStock2.Visible = True
                        lblStock2.Top = lblStock1.Top
                        lblStock2.Left = 100 + 40 * (2 - 1)
                        lblOrder2.Visible = True
                        lblOrder2.Top = lblOrder1.Top
                        lblOrder2.Left = 100 + 40 * (2 - 1)
                    Case 3
                        lblStore3.Visible = True
                        lblStore3.Top = lblStore1.Top
                        'lblStore3.Left = 90 + 60 * (3 - 1)
                        lblStore3.Left = 90 + 100
                        lblStock3.Visible = True
                        lblStock3.Top = lblStock1.Top
                        'lblStock3.Left = 100 + 60 * (3 - 1)
                        lblStock3.Left = 100 + 90
                        lblOrder3.Visible = True
                        lblOrder3.Top = lblOrder1.Top
                        'lblOrder3.Left = 100 + 50 * (3 - 1)
                        lblOrder3.Left = 100 + 90
                    Case 4
                        lblStore4.Visible = True
                        lblStore4.Top = lblStore1.Top
                        'lblStore4.Left = 90 + 60 * (4 - 1)
                        lblStore4.Left = 90 + 150
                        lblStock4.Visible = True
                        lblStock4.Top = lblStock1.Top
                        'lblStock4.Left = 100 + 60 * (4 - 1)
                        lblStock4.Left = 100 + 140
                        lblOrder4.Visible = True
                        lblOrder4.Top = lblOrder1.Top
                        'lblOrder4.Left = 100 + 50 * (4 - 1)
                        lblOrder4.Left = 100 + 140
                    Case 5
                        lblStore5.Visible = True
                        lblStore5.Top = lblStore1.Top
                        'lblStore5.Left = 90 + 60 * (5 - 1)
                        lblStore5.Left = 90 + 200
                        lblStock5.Visible = True
                        lblStock5.Top = lblStock1.Top
                        'lblStock5.Left = 100 + 60 * (5 - 1)
                        lblStock5.Left = 100 + 190
                        lblOrder5.Visible = True
                        lblOrder5.Top = lblOrder1.Top
                        'lblOrder5.Left = 100 + 50 * (5 - 1)
                        lblOrder5.Left = 100 + 190
                    Case 6
                        lblStore6.Visible = True
                        lblStore6.Top = lblStore1.Top
                        'lblStore6.Left = 90 + 60 * (6 - 1)
                        lblStore6.Left = 90 + 250
                        lblStock6.Visible = True
                        lblStock6.Top = lblStock1.Top
                        'lblStock6.Left = 100 + 60 * (6 - 1)
                        lblStock6.Left = 100 + 240
                        lblOrder6.Visible = True
                        lblOrder6.Top = lblOrder1.Top
                        'lblOrder6.Left = 100 + 50 * (6 - 1)
                        lblOrder6.Left = 100 + 240
                    Case 7
                        lblStore7.Visible = True
                        lblStore7.Top = lblStore1.Top
                        'lblStore7.Left = 90 + 60 * (7 - 1)
                        lblStore7.Left = 90 + 300
                        lblStock7.Visible = True
                        lblStock7.Top = lblStock1.Top
                        'lblStock7.Left = 100 + 60 * (7 - 1)
                        lblStock7.Left = 100 + 290
                        lblOrder7.Visible = True
                        lblOrder7.Top = lblOrder1.Top
                        'lblOrder7.Left = 100 + 50 * (7 - 1)
                        lblOrder7.Left = 100 + 290
                    Case 8
                        lblStore8.Visible = True
                        lblStore8.Top = lblStore1.Top
                        'lblStore8.Left = 90 + 60 * (8 - 1)
                        lblStore8.Left = 90 + 350
                        lblStock8.Visible = True
                        lblStock8.Top = lblStock1.Top
                        'lblStock8.Left = 100 + 60 * (8 - 1)
                        lblStock8.Left = 100 + 340
                        lblOrder8.Visible = True
                        lblOrder8.Top = lblOrder1.Top
                        'lblOrder8.Left = 100 + 50 * (8 - 1)
                        lblOrder8.Left = 100 + 340
                End Select
            End If
        Next
    End Sub

    Private Sub StoreVisible(ByVal Lname As String)
        For Each c As Control In Me.Controls
            If c.Name = Lname Then
                c.Visible = False
                Exit For
            End If
        Next
    End Sub

    Private Sub StockVisible(ByVal Lname As String)
        For Each c As Control In Me.Controls
            If c.Name = Lname Then
                c.Visible = False
                Exit For
            End If
        Next
    End Sub

    Private Sub OrderVisible(ByVal Lname As String)
        For Each c As Control In Me.Controls
            If c.Name = Lname Then
                c.Visible = False
                Exit For
            End If
        Next
    End Sub
    Private Sub cmdPrevStores_Click(sender As Object, e As EventArgs) Handles cmdPrevStores.Click
        InfoBase = 1
        ArrangeInfo()
    End Sub

    Private Sub OrdPreview_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim X As Integer
        'SetButtonImage cmdClose, "menu"
        'SetButtonImage cmdPrint, "print"
        SetButtonImage(cmdClose, 9)
        SetButtonImage(cmdPrint, 19)
        WindowState = FormWindowState.Maximized
        ShowLabels(False, True)

        StoreName.Text = StoreSettings.Name
        LoadPix()

        LoadDeptNamesIntoComboBox(cboDepartment, , , True)
        LoadMfgNamesIntoComboBox(cboVendor, , True, True)
    End Sub

    Private Sub SelectVendor(ByVal V As String)
        On Error Resume Next
        cboVendor.Text = V
    End Sub

    Private Sub SelectDept(ByVal N As Integer)
        On Error Resume Next
        Dim I As Integer
        For I = 0 To cboDepartment.Items.Count - 1
            'If cboDepartment.itemData(I) = N Then cboDepartment.Text = cboDepartment.List(I)
            If CType(cboDepartment.Items(cboDepartment.SelectedIndex), ItemDataClass).ItemData = N Then cboDepartment.Text = cboDepartment.Items(cboDepartment.SelectedIndex)
        Next
    End Sub

    Private Function GetRNFromStyle(ByVal Style As String) As Integer
        Dim DataObj As New CInvRec
        GetRNFromStyle = -1
        If DataObj.Load(Style, "Style") Then GetRNFromStyle = DataObj.RN
        DisposeDA(DataObj)
    End Function

    Private Sub txtStyle_Enter(sender As Object, e As EventArgs) Handles txtStyle.Enter
        SelectContents(txtStyle)
    End Sub

    Private Sub txtStyle_Validating(sender As Object, e As CancelEventArgs) Handles txtStyle.Validating
        Dim X As Integer
        If txtStyle.Text = "" Then Exit Sub
        X = GetRNFromStyle(txtStyle.Text)
        If X < 0 Then
            MessageBox.Show("Unable to find style " & txtStyle.Text & ".", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            e.Cancel = True
            Exit Sub
        End If
        LoadItemByRN(X)
    End Sub

    Private Sub cmdMoveFirst_Click(sender As Object, e As EventArgs) Handles cmdMoveFirst.Click
        Navigate(False, True)
    End Sub

    Private Sub cmdMoveLast_Click(sender As Object, e As EventArgs) Handles cmdMoveLast.Click
        Navigate(True, True)
    End Sub

    Private Sub cmdMoveNext_Click(sender As Object, e As EventArgs) Handles cmdMoveNext.Click
        Navigate(True, False)
    End Sub

    Private Sub cmdMovePrevious_Click(sender As Object, e As EventArgs) Handles cmdMovePrevious.Click
        Navigate(False, False)
    End Sub
End Class
