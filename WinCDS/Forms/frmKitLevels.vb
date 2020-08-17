Public Class frmKitLevels
    Private mLocation As Integer
    Public Cancelled As Boolean
    Public AllowPartialKits As Boolean
    Public AllowAdjustedQuantities As Boolean
    Public AllowStatusChange As Boolean
    Public AllowItemStatusChange As Boolean
    Public AllowItemLocChange As Boolean
    Public lblItemNumCount As Integer
    Public IsfrmKitLevelsHide As Boolean
    Dim N As Integer
    'Dim TxtItemQty As TextBox

    Public Function KitCost(Optional ByVal vType As String = "Landed", Optional ByVal Line As Integer = 0, Optional ByVal Style As String = "") As Decimal
        Dim I As Integer, S As String, C As CInvRec
        For I = 1 To ItemCount
            If (Line = 0 And Style = "") Or Line = I Or Style = ItemStyle(I) Then
                S = ItemStyle(I)
                C = New CInvRec
                If C.Load(S, "Style") Then
                    Select Case UCase(vType)
                        Case "LANDED" : KitCost = KitCost + C.Landed * ItemQuantityByStyle(S)
                        Case "COST" : KitCost = KitCost + C.Cost * ItemQuantityByStyle(S)
                        Case "ONSALE" : KitCost = KitCost + C.OnSale * ItemQuantityByStyle(S)
                        Case "LIST" : KitCost = KitCost + C.List * ItemQuantityByStyle(S)
                        Case Else
                            Err.Raise(-1, "", "Invalid vType in frmKitLevels.KitCost()")
                    End Select
                End If
                DisposeDA(C)
            End If
        Next
    End Function

    Public ReadOnly Property ItemQuantityByStyle(ByVal vStyle As String) As Double
        Get
            Dim I As Integer
            For I = 1 To ItemCount
                If ItemStyle(I) = vStyle Then ItemQuantityByStyle = ItemQuantity(I) : Exit Property
            Next
            ItemQuantityByStyle = Quantity
        End Get
    End Property

    Public ReadOnly Property ItemLocByStyle(ByVal vStyle As String) As Double
        Get
            Dim I As Integer
            For I = 1 To ItemCount
                If ItemStyle(I) = vStyle Then ItemLocByStyle = ItemLoc(I) : Exit Property
            Next
            ItemLocByStyle = Locations   '--> In vb 6, this property name is Location. Changed to Locations because Location is a keyword in vb.net
        End Get
    End Property

    Public ReadOnly Property ItemStatusByStyle(ByVal vStyle As String) As String
        Get
            Dim I As Integer
            For I = 1 To ItemCount
                If ItemStyle(I) = vStyle Then ItemStatusByStyle = ItemStatus(I) : Exit Property
            Next
            ItemStatusByStyle = status
        End Get
    End Property

    Public Sub LoadKit(ByVal vLoc As Integer, ByVal vStat As String, ByVal KitStyle As String, Optional ByVal Quan As Double = 1)
        Dim CI As cInvKit, I As Integer
        On Error Resume Next
        ClearItems()

        Locations = vLoc
        status = vStat

        CI = New cInvKit
        If CI.Load(KitStyle, "KitStyleNo") Then
            lblStyle.Text = KitStyle
            txtKitQuantity.Text = Quan

            For I = 1 To Setup_MaxKitItems
                'For I = 1 To 2
                If CI.Item(I) <> "" Then
                    AddItem(CI.Item(I), Quan * CI.Quantity(I), Locations, CI.Quantity(I))
                End If
            Next
        End If
        DisposeDA(CI)
    End Sub

    Public Property status As String
        Get
            status = cmdStatus.Text
            If status = "" Then status = "ST"
        End Get
        Set(value As String)
            cmdStatus.Text = value
        End Set
    End Property

    Public Property Quantity As Double
        Get
            Quantity = Val(Trim(txtKitQuantity.Text))
        End Get
        Set(value As Double)
            txtKitQuantity.Text = value
        End Set
    End Property

    Public ReadOnly Property ItemCount() As Integer
        Get
            Dim cmdItemStatusCount As Integer
            For Each ctrl As Control In Me.fraItems.Controls
                If Mid(ctrl.Name, 1, 13) = "cmdItemStatus" Then
                    cmdItemStatusCount = cmdItemStatusCount + 1
                End If
            Next
            'ItemCount = cmdItemStatus.UBound
            ItemCount = cmdItemStatusCount
        End Get
    End Property

    '    Public Property Get ItemStyle(ByVal Index as integer) As String
    '    If Index <= 0 Then Exit Property
    '  If Index > ItemCount Then Exit Property
    '  ItemStyle = lblItem(Index)
    'End Property
    '    Public Property Let ItemStyle(ByVal Index as integer, ByVal vData As String)
    '    If Index <= 0 Then Exit Property
    '  If Index > ItemCount Then Exit Property
    '  lblItem(Index).Caption = vData
    'End Property


    Public ReadOnly Property ItemQuantity(ByVal Index As Integer) As Double
        Get
            If Index < 0 Then Exit Property
            If Index > ItemCount Then Exit Property

            'ItemQuantity = Val(txtItemQuan(Index))
            'ItemQuantity = txtItemQuan & Index & .text
            If Index = 1 Then
                ItemQuantity = txtItemQuan.Text
            Else
                Dim T As TextBox
                For Each ctrl As Control In Me.fraItems.Controls
                    If ctrl.Name = "txtItemQuan" & Index Then
                        T = ctrl
                        ItemQuantity = T.Text
                        Exit For
                    End If
                Next
            End If
        End Get

    End Property

    Public Property Locations() As Integer
        '--> In vb 6, this property name is Location. Changed to Locations because Location is a keyword in vb.net
        Get
            Locations = mLocation
            If Locations = 0 Then Locations = StoresSld
        End Get
        Set(value As Integer)
            mLocation = value
        End Set
    End Property

    Private Sub ClearItems()
        Dim I As Integer
        lblStyle.Text = ""
        txtKitQuantity.Text = "0"
        'lblItem(1) = ""
        lblItem.Text = ""
        'txtItemQuan(1) = "0"
        txtItemQuan.Text = "0"
        'lblItemLoc(1) = "0"
        lblItemLoc.Text = "0"
        'lblOnOrd(1) = "0"
        lblOnOrd.Text = "0"
        'lblItemAvail(1) = "0"
        lblItemAvail.Text = "0"
        'cmdItemLoc(1).Caption = "L" & Location
        cmdItemLoc.Text = "L" & Locations
        'cmdItemStatus(1).Caption = status
        cmdItemStatus.Text = status

        'On Error Resume Next
        '    If lblItemNum.UBound >= 2 Then
        '        For I = lblItemNum.UBound To 2 Step -1
        '            Unload lblItemNum(I)
        '  Unload lblItem(I)
        '  Unload txtItemQuan(I)
        '  Unload lblItemLoc(I)
        '  Unload lblOnOrd(I)
        '  Unload lblItemAvail(I)
        '  Unload cmdItemLoc(I)
        '  Unload cmdItemStatus(I)
        'Next
        '    End If


        lblItemNumCount = 0
        For Each ctrl As Control In Me.fraItems.Controls
            If Mid(ctrl.Name, 1, 10) = "lblItemNum" And Len(ctrl.Name) <= 13 Then
                lblItemNumCount = lblItemNumCount + 1
            End If

        Next

        If lblItemNumCount >= 2 Then
            'For I = lblItemNumCount To 2 Step -1
            'For I = 2 To lblItemNumCount
            'Me.Controls.Item(lblItemNum.ToString & I).Hide()
            'Me.Controls.Item(lblItem.ToString & I).Hide()
            'Me.Controls.Item(txtItemQuan.ToString & I).Hide()
            'Me.Controls.Item(lblItemLoc.ToString & I).Hide()
            'Me.Controls.Item(lblOnOrd.ToString & I).Hide()
            'Me.Controls.Item(lblItemAvail.ToString & I).Hide()
            'Me.Controls.Item(cmdItemLoc.ToString & I).Hide()
            'Me.Controls.Item(cmdItemStatus.ToString & I).Hide()

            For I = 2 To lblItemNumCount
                For Each c As Control In Me.fraItems.Controls
                    '    'Debug.Print(c.Name)
                    If c.Name = "lblItemNum" & I Then
                        Me.fraItems.Controls.Remove(c)

                    ElseIf c.Name = "lblItem" & I Then
                        Me.fraItems.Controls.Remove(c)

                    ElseIf c.Name = "txtItemQuan" & I Then
                        Me.fraItems.Controls.Remove(c)

                    ElseIf c.Name = "lblItemLoc" & I Then
                        Me.fraItems.Controls.Remove(c)

                    ElseIf c.Name = "lblOnOrd" & I Then
                        Me.fraItems.Controls.Remove(c)

                    ElseIf c.Name = "lblItemAvail" & I Then
                        Me.fraItems.Controls.Remove(c)

                    ElseIf c.Name = "cmdItemLoc" & I Then
                        Me.fraItems.Controls.Remove(c)

                    ElseIf c.Name = "cmdItemStatus" & I Then
                        Me.fraItems.Controls.Remove(c)
                    End If
                Next
            Next
        End If

        'fraItems.Height = cmdItemStatus(1).Top
        fraItems.Height = cmdItemStatus.Top
        fraItems.Visible = False
        fraControls.Top = fraItems.Top + fraItems.Height
        'Height = Height - ScaleHeight + fraControls.Top + fraControls.Height
        Height = Height - Me.ClientSize.Height + fraControls.Top + fraControls.Height
        cmdStatus.Enabled = AllowStatusChange
    End Sub

    Private Sub AddItem(ByVal Style As String, ByRef Q As Double, Optional ByVal vLoc As Integer = 0, Optional ByVal SingleQuantity As Double = 0)
        'Dim N As Integer
        Dim T As Integer
        Dim A As Double, B As Double, C As Double, D As String, E As Double
        Dim lblItemCount As Integer
        Dim labelItemText As String = ""
        Dim lblItemNumber As String

        Dim L1, L2, L3, L4, L5, L6, L7, L8 As Integer

        If vLoc = 0 Then vLoc = Locations
        GetItem(vLoc, Style, A, B, C, D, E)

        lblItemNumCount = 0
        lblItemCount = 1


        For Each ctrl As Control In Me.fraItems.Controls
            If Mid(ctrl.Name, 1, 10) = "lblItemNum" And Len(ctrl.Name) <= 13 Then
                lblItemNumCount = lblItemNumCount + 1
            End If

            lblItemNumber = Mid(ctrl.Name, 8, 1)
            If Mid(ctrl.Name, 1, 7) = "lblItem" And IsNumeric(lblItemNumber) Then
                lblItemCount = lblItemCount + 1
            End If
            'Debug.Print(ctrl.Name)
        Next

        'For Each ctrl As Control In Me.fraItems.Controls
        '    lblItemNumber = Mid(ctrl.Name, 8, 1)
        '    If Mid(ctrl.Name, 1, 7) = "lblItem" And IsNumeric(lblItemNumber) Then
        '        lblItemCount = lblItemCount + 1
        '    Else
        '        lblItemCount = 1
        '    End If
        'Next

        'N = lblItemNum.UBound
        N = lblItemNumCount
        'If N > 1 Or lblItem(N) <> "" Then
        'If N > 1 Or lblItem.ToString & N & ".Text" <> "" Then

        If lblItemCount = 1 Then
            labelItemText = lblItem.Text
        Else
            'labelItemText = Me.Controls.Item(lblItem.ToString & N).Text
            'Note: If the above line will not work, replace it with below three commented lines to get the text and store it in labelItemText variable.
            'Dim Lc As New Label
            'Lc.Name = "lblItem" & N
            'labelItemText = Lc.Text
            Dim L As Label
            For Each ctrl As Control In Me.fraItems.Controls
                If ctrl.Name = "lblItem" & N Then
                    L = ctrl
                    labelItemText = L.Text
                    Exit For
                End If
            Next
        End If

        Dim ctrll As Control
        'If N > 1 Or lblItem.Text <> "" Then
        If N > 1 Or labelItemText <> "" Then
            If lblItemNumCount = 1 Then
                'T = lblItemNum(N).Top + 240
                T = lblItemNum.Top + 18
                L1 = lblItemNum.Left
                L2 = lblItem.Left
                L3 = txtItemQuan.Left
                L4 = lblItemLoc.Left
                L5 = lblOnOrd.Left
                L6 = lblItemAvail.Left
                L7 = cmdItemLoc.Left
                L8 = cmdItemStatus.Left
            Else
                'T = lblItemNum(N).Top + 240
                'T = Me.Controls.Item(lblItemNum.ToString & N).Top + 240
                'Note: If the above line will not work, replace it with the below three commented lines to get top and increment with 240 and assign it to variable T.
                Dim L As Control
                Dim lblItemNumFound, lblItemFound, txtItemQuanFound, lblItemLocFound, lblOnOrdFound, lblItemAvailFound, cmdItemLocfound, cmdItemStatusFound As Boolean
                Dim FoundCount As Integer
                'L.Name = "lblItemNum" & N
                'T = L.Top + 240
                For Each cc As Control In Me.fraItems.Controls

                    If cc.Name = "lblItemNum" & N And lblItemNumFound = False Then
                        L = cc
                        T = L.Top + 18
                        L1 = L.Left
                        lblItemNumFound = True
                        FoundCount = FoundCount + 1
                        'Debug.Print(ctrl.Name)
                    End If
                    If cc.Name = "lblItem" & N And lblItemFound = False Then
                        L = cc
                        L2 = L.Left
                        lblItemFound = True
                        FoundCount = FoundCount + 1
                        'Debug.Print(ctrl.Name)
                    End If
                    If cc.Name = "txtItemQuan" & N And txtItemQuanFound = False Then
                        L = cc
                        L3 = L.Left
                        txtItemQuanFound = True
                        FoundCount = FoundCount + 1
                        'Debug.Print(ctrl.Name)
                    End If
                    If cc.Name = "lblItemLoc" & N And lblItemLocFound = False Then
                        L = cc
                        L4 = L.Left
                        lblItemLocFound = True
                        FoundCount = FoundCount + 1
                        'Debug.Print(ctrl.Name)
                    End If
                    If cc.Name = "lblOnOrd" & N And lblOnOrdFound = False Then
                        L = cc
                        L5 = L.Left
                        lblOnOrdFound = True
                        FoundCount = FoundCount + 1
                        'Debug.Print(ctrl.Name)
                    End If
                    If cc.Name = "lblItemAvail" & N And lblItemAvailFound = False Then
                        L = cc
                        L6 = L.Left
                        lblItemAvailFound = True
                        FoundCount = FoundCount + 1
                        'Debug.Print(ctrl.Name)
                    End If
                    If cc.Name = "cmdItemLoc" & N And cmdItemLocfound = False Then
                        L = cc
                        L7 = L.Left
                        cmdItemLocfound = True
                        FoundCount = FoundCount + 1
                        'Debug.Print(ctrl.Name)
                    End If
                    If cc.Name = "cmdItemStatus" & N And cmdItemStatusFound = False Then
                        L = cc
                        L8 = L.Left
                        cmdItemStatusFound = True
                        FoundCount = FoundCount + 1
                        'Debug.Print(ctrl.Name)
                    End If
                    If FoundCount = 8 Then
                        Exit For
                    End If
                Next
            End If

            N = N + 1
            'Load lblItemNum(N)
            'Me.Controls.Item(lblItemNum.ToString & N).Show()
            'Me.Controls.Item(lblItemNum.ToString & N).Hide()
            'lblItemNum(N).Top = T
            'Me.Controls.Item(lblItemNum.ToString & N).Top = T
            ctrll = New Label
            'ctrll.Name = lblItemNum.ToString & N
            ctrll.Name = "lblItemNum" & N
            'ctrll.Top = T
            ToolTip1.SetToolTip(ctrll, D)
            ctrll.Location = New Point(L1, T)
            ctrll.Size = New Size(12, 11)
            ctrll.Text = N
            ctrll.Font = New Font("Lucida Console", 8, FontStyle.Regular)
            'Me.Controls.Add(ctrll)
            Me.fraItems.Controls.Add(ctrll)
            'DirectCast(Me.fraItems.Controls(ctrll.Name), Label).TextAlign = ContentAlignment.TopLeft
            DirectCast(Me.fraItems.Controls.Item(ctrll.Name), Label).TextAlign = ContentAlignment.TopLeft
            'ctrll.Hide()

            'Load lblItem(N)
            'Me.Controls.Item(lblItem.ToString & N).Show()
            'Me.Controls.Item(lblItem.ToString & N).Hide()
            'lblItem(N).Top = T
            'Me.Controls.Item(lblItem.ToString & N).Top = T
            ctrll = New Label
            ctrll.Name = "lblItem" & N
            'ctrll.Top = T
            ctrll.Text = Style
            ToolTip1.SetToolTip(ctrll, D)
            ctrll.Size = New Size(97, 11)
            ctrll.Location = New Point(L2, T)
            ctrll.Font = New Font("Lucida Console", 8, FontStyle.Regular)
            'Me.Controls.Add(ctrll)
            Me.fraItems.Controls.Add(ctrll)
            DirectCast(Me.fraItems.Controls.Item(ctrll.Name), Label).TextAlign = ContentAlignment.TopRight

            'ctrll.Hide()

            'Load txtItemQuan(N)
            'Me.Controls.Item(txtItemQuan.ToString & N).Show()
            'Me.Controls.Item(txtItemQuan.ToString & N).Hide()
            'txtItemQuan(N).Top = T
            'Me.Controls.Item(txtItemQuan.ToString & N).Top = T
            ctrll = New TextBox
            ctrll.Name = "txtItemQuan" & N
            'ctrll.Top = T
            AddHandler ctrll.TextChanged, AddressOf txtItemQuanTextChanged
            ctrll.Text = Math.Round(Q, 2)
            ctrll.Tag = SingleQuantity
            ctrll.Enabled = Not AllowAdjustedQuantities
            'If ctrll.Enabled = True Then
            'txtItemQuan(N).Appearance = 0  -Appearance property not available in vb.net
            'ctrll.BackColor = Color.White
            'txtItemQuan(N).BorderStyle = 0 -Borderstyle property not available

            'Else
            'txtItemQuan(N).Appearance = 1  -Appearance property not available in vb.net
            'ctrll.BackColor = Color.Gray
            'txtItemQuan(N).BorderStyle = 1 -Borderstyle property not available
            'End If
            ctrll.Location = New Point(L3, T)
            ctrll.Size = New Size(44, 18)
            ctrll.Font = New Font("Lucida Console", 8, FontStyle.Regular)
            'Me.Controls.Add(ctrll)
            Me.fraItems.Controls.Add(ctrll)
            DirectCast(Me.fraItems.Controls.Item(ctrll.Name), TextBox).TextAlign = HorizontalAlignment.Right
            DirectCast(Me.fraItems.Controls.Item(ctrll.Name), TextBox).BorderStyle = BorderStyle.Fixed3D
            'DirectCast(Me.fraItems.Controls.Item(ctrll.Name), TextBox).ReadOnly = Not AllowAdjustedQuantities
            If ctrll.Enabled = True Then
                ctrll.BackColor = Color.White
                'DirectCast(Me.fraItems.Controls.Item(ctrll.Name), TextBox).BorderStyle = BorderStyle.Fixed3D
            Else
                ctrll.BackColor = Color.Gray
                'DirectCast(Me.fraItems.Controls.Item(ctrll.Name), TextBox).BorderStyle = BorderStyle.None
            End If
            'ctrll.Hide()

            'Load lblItemLoc(N)
            'Me.Controls.Item(lblItemLoc.ToString & N).Show()
            'Me.Controls.Item(lblItemLoc.ToString & N).Hide()
            'lblItemLoc(N).Top = T
            'Me.Controls.Item(lblItemLoc.ToString & N).Top = T
            ctrll = New Label
            ctrll.Name = "lblItemLoc" & N
            'ctrll.Top = T
            If ShowST Then
                'lblItemLoc.Visible = True -> This line not required because in the above lblItemLoc is added with default visible is true.
                ctrll.Visible = True
            Else
                lblItemLocCaption.Visible = False
                'lblItemLoc.Visible = False
                ctrll.Visible = False
            End If
            ctrll.Text = A
            ctrll.Location = New Point(L4, T)
            ctrll.Size = New Size(12, 11)
            ctrll.Font = New Font("Lucida Console", 8, FontStyle.Regular)
            'Me.Controls.Add(ctrll)

            Me.fraItems.Controls.Add(ctrll)
            DirectCast(Me.fraItems.Controls.Item(ctrll.Name), Label).TextAlign = ContentAlignment.TopRight
            'ctrll.Hide()

            'Load lblOnOrd(N)
            'Me.Controls.Item(lblOnOrd.ToString & N).Show()
            'Me.Controls.Item(lblOnOrd.ToString & N).Hide()
            'lblOnOrd(N).Top = T
            'Me.Controls.Item(lblOnOrd.ToString & N).Top = T
            ctrll = New Label
            ctrll.Name = "lblOnOrd" & N
            'ctrll.Top = T
            ctrll.Text = B
            ctrll.Tag = E
            ctrll.Location = New Point(L5, T)
            ctrll.Size = New Size(12, 11)
            ctrll.Font = New Font("Lucida Console", 8, FontStyle.Regular)
            'Me.Controls.Add(ctrll)

            Me.fraItems.Controls.Add(ctrll)
            DirectCast(Me.fraItems.Controls.Item(ctrll.Name), Label).TextAlign = ContentAlignment.TopRight
            'ctrll.Hide()

            'Load lblItemAvail(N)
            'Me.Controls.Item(lblItemAvail.ToString & N).Show()
            'Me.Controls.Item(lblItemAvail.ToString & N).Hide()
            'lblItemAvail(N).Top = T
            'Me.Controls.Item(lblItemAvail.ToString & N).Top = T
            ctrll = New Label
            ctrll.Name = "lblItemAvail" & N
            'ctrll.Top = T
            ctrll.Text = C
            ctrll.Location = New Point(L6, T)
            ctrll.Size = New Size(12, 11)
            ctrll.Font = New Font("Lucida Console", 8, FontStyle.Regular)
            'Me.Controls.Add(ctrll)

            Me.fraItems.Controls.Add(ctrll)
            DirectCast(Me.fraItems.Controls.Item(ctrll.Name), Label).TextAlign = ContentAlignment.TopRight
            'ctrll.Hide()

            'Load cmdItemLoc(N)
            'Me.Controls.Item(cmdItemLoc.ToString & N).Show()
            'Me.Controls.Item(cmdItemLoc.ToString & N).Hide()
            'cmdItemLoc(N).Top = T
            'Me.Controls.Item(cmdItemLoc.ToString & N).Top = T
            ctrll = New Button
            ctrll.Name = "cmdItemLoc" & N
            'ctrll.Top = T
            ctrll.Text = "L" & vLoc
            'ctrll.Enabled = AllowItemLocChange
            ctrll.Location = New Point(L7, T)
            ctrll.Size = New Size(30, 20)
            ctrll.Font = New Font("Lucida Console", 8, FontStyle.Regular)
            'Me.Controls.Add(ctrll)

            Me.fraItems.Controls.Add(ctrll)
            DirectCast(Me.fraItems.Controls.Item(ctrll.Name), Button).TextAlign = ContentAlignment.MiddleCenter
            AddHandler ctrll.Click, AddressOf cmdItemLoc_Click
            'ctrll.Hide()

            'Load cmdItemStatus(N)
            'Me.Controls.Item(cmdItemStatus.ToString & N).Show()
            'Me.Controls.Item(cmdItemStatus.ToString & N).Hide()
            'cmdItemStatus(N).Top = T
            'Me.Controls.Item(cmdItemStatus.ToString & N).Top = T
            ctrll = New Button
            ctrll.Name = "cmdItemStatus" & N
            'ctrll.Top = T
            ctrll.Text = status
            ctrll.Enabled = AllowItemStatusChange
            ctrll.Location = New Point(L8, T)
            ctrll.Size = New Size(44, 20)
            ctrll.Font = New Font("Lucida Console", 8, FontStyle.Regular)
            'Me.Controls.Add(ctrll)

            Me.fraItems.Controls.Add(ctrll)
            DirectCast(Me.fraItems.Controls.Item(ctrll.Name), Button).TextAlign = ContentAlignment.MiddleCenter
            AddHandler ctrll.Click, AddressOf cmdItemStatus_Click
            'ctrll.Hide()

        End If

        If N = 1 Then
            lblItemNum.Visible = True
            ToolTip1.SetToolTip(lblItemNum, D)
            lblItem.Visible = True
            lblItem.Text = Style
            ToolTip1.SetToolTip(lblItem, D)

            txtItemQuan.Visible = True
            txtItemQuan.Text = Math.Round(Q, 2)
            txtItemQuan.Tag = SingleQuantity
            txtItemQuan.Enabled = Not AllowAdjustedQuantities   '-> Locked replaced with Enabled. Because Locked propert not available at runtime in vb.net
            If txtItemQuan.Enabled = True Then
                'txtItemQuan.Appearance = 0         Property not available.
                'txtItemQuan.BackColor = &H8000000F   Hexadecimal not accepted in vb.net
                txtItemQuan.BorderStyle = BorderStyle.Fixed3D
            Else
                'txtItemQuan.Appearance = 1
                'txtItemQuan.BackColor = &H80000005
                txtItemQuan.BorderStyle = BorderStyle.Fixed3D

            End If

            If ShowST Then
                lblItemLoc.Visible = True
            Else
                lblItemLocCaption.Visible = False
                lblItemLoc.Visible = False
            End If

            lblItemLoc.Text = A
            lblOnOrd.Visible = True
            lblOnOrd.Text = B
            lblOnOrd.Tag = E
            lblItemAvail.Visible = True
            lblItemAvail.Text = C
            cmdItemLoc.Visible = True
            cmdItemLoc.Text = "L" & vLoc                ' original setup doesn't use property...  don't need update call
            'cmdItemLoc.Enabled = AllowItemLocChange
            cmdItemStatus.Visible = True
            cmdItemStatus.Text = status
            cmdItemStatus.Enabled = AllowItemStatusChange

            'ElseIf N > 1 Then
            '    '    'lblItemNum(N).Visible = True
            '    '    'lblItemNum(N) = N
            '    '    'lblItemNum(N).ToolTipText = D
            '    Me.Controls.Item(lblItemNum.ToString & N).Visible = True
            '    ToolTip1.SetToolTip(Me.Controls.Item(lblItemNum.ToString & N), D)
            '    '    'lblItem(N).Visible = True
            '    '    'lblItem(N) = Style
            '    '    'lblItem(N).ToolTipText = D
            '    Me.Controls.Item(lblItem.ToString & N).Visible = True
            '    Me.Controls.Item(lblItem.ToString & N).Text = Style
            '    ToolTip1.SetToolTip(Me.Controls.Item(lblItem.ToString & N), D)
            '    '    'txtItemQuan(N).Visible = True
            '    '    'txtItemQuan(N) = Round(Q, 2)
            '    '    'txtItemQuan(N).Tag = SingleQuantity
            '    '    'txtItemQuan(N).Locked = Not AllowAdjustedQuantities

            '    '    'If txtItemQuan(N).Locked Then
            '    '        'txtItemQuan(N).Appearance = 0   --> This is only for show some 3-d effects at runtime.
            '    '        'txtItemQuan(N).BackColor = &H8000000F
            '    '        'txtItemQuan(N).BorderStyle = 0   --> This property at runtime is not available. Only available at designtime.
            '    '    Else
            '    '        'txtItemQuan(N).Appearance = 1            --> appearance property not avaible for textbox 
            '    '        'txtItemQuan(N).BackColor = &H80000005    --> backcolor hexadecimal value not accepted in vb.net
            '    '        'txtItemQuan(N).BorderStyle = 1           --> borderstyle not available at runtime.    
            '    '    End If
            '    Me.Controls.Item(txtItemQuan.ToString & N).Visible = True
            '    Me.Controls.Item(txtItemQuan.ToString & N).Text = Math.Round(Q, 2)
            '    Me.Controls.Item(txtItemQuan.ToString & N).Tag = SingleQuantity
            '    Me.Controls.Item(txtItemQuan.ToString & N).Enabled = Not AllowAdjustedQuantities   '-> replaced Locked with Enabled. Locked will not available at runtime in vb.net
            '    If Me.Controls.Item(txtItemQuan.ToString & N).Enabled = True Then
            '        'txtItemQuan(N).Appearance = 0   --> This is only for show some 3-d effects at runtime.
            '        Me.Controls.Item(txtItemQuan.ToString & N).BackColor = Color.White  '->Hexadecimal is not accepted in me.controls.item style.
            '        ''txtItemQuan(N).BorderStyle = 0   --> Border style is available if it is a direct textbox control. For me.controls.item(txtItemQuan) style
            '        '--> Border style is not available. To get the borderstyle, instead of me.controls.item, use the below commented code.
            '        'Dim Tc As New TextBox
            '        'Tc.Name = "txtItemQuan" & N
            '        'Tc.BackColor = Color.White  '-> Hexadecimal is not accepted for backcolor.
            '        'Tc.BorderStyle = 0
            '        'Me.Controls.Add(Tc)
            '    Else
            '        'Appearance property is not available in vb.net
            '        'BorderStyle will not available using me.controls.item style. For this use Dim Tc as New TextBox style.
            '        Me.Controls.Item(txtItemQuan.ToString & N).BackColor = Color.White  '->Hexadecimal is not accepted in me.controls.item style.
            '    End If

            '    If ShowST Then
            '        '        'lblItemLoc(N).Visible = True
            '        Me.Controls.Item(lblItemLoc.ToString & N).Visible = True
            '    Else
            '        lblItemLocCaption.Visible = False
            '        '        'lblItemLoc(N).Visible = False
            '        Me.Controls.Item(lblItemLoc.ToString & N).Visible = False
            '    End If
            '    '    'lblItemLoc(N) = A
            '    Me.Controls.Item(lblItemLoc.ToString & N).Text = A
            '    '    'lblOnOrd(N).Visible = True
            '    '    'lblOnOrd(N) = B
            '    '    'lblOnOrd(N).Tag = E
            '    Me.Controls.Item(lblOnOrd.ToString & N).Visible = True
            '    Me.Controls.Item(lblOnOrd.ToString & N).Text = B
            '    Me.Controls.Item(lblOnOrd.ToString & B).Tag = E
            '    '    'lblItemAvail(N).Visible = True
            '    '    'lblItemAvail(N) = C
            '    Me.Controls.Item(lblItemAvail.ToString & N).Visible = True
            '    Me.Controls.Item(lblItemAvail.ToString & N).Text = C
            '    '    'cmdItemLoc(N).Visible = True
            '    '    'cmdItemLoc(N).Caption = "L" & vLoc                ' original setup doesn't use property...  don't need update call
            '    '    'cmdItemLoc(N).Enabled = AllowItemLocChange
            '    Me.Controls.Item(cmdItemLoc.ToString & N).Visible = True
            '    Me.Controls.Item(cmdItemLoc.ToString & N).Text = "L" & vLoc
            '    Me.Controls.Item(cmdItemLoc.ToString & N).Enabled = AllowItemLocChange
            '    '    'cmdItemStatus(N).Visible = True
            '    '    'cmdItemStatus(N).Caption = status
            '    '    'cmdItemStatus(N).Enabled = AllowItemStatusChange
            '    Me.Controls.Item(cmdItemStatus.ToString & N).Visible = True
            '    Me.Controls.Item(cmdItemStatus.ToString & N).Text = status
            '    Me.Controls.Item(cmdItemStatus.ToString & N).Enabled = AllowItemStatusChange
        End If

        HiLiteKitRow(N)

        If N = 1 Then
            'fraItems.Height = cmdItemStatus(N).Top + cmdItemStatus(N).Height + 60
            fraItems.Height = cmdItemStatus.Top + cmdItemStatus.Height
            fraItems.Visible = True
            fraControls.Top = fraItems.Top + fraItems.Height
            'Height = Height - ScaleHeight + fraControls.Top + fraControls.Height + 120
            Height = Height - Me.ClientSize.Height + fraControls.Top + fraControls.Height
        ElseIf N > 1 Then
            'fraItems.Height = cmdItemStatus(N).Top + cmdItemStatus(N).Height + 60
            'fraItems.Height = Me.Controls.Item(cmdItemStatus.ToString & N).Top + Me.Controls.Item(cmdItemStatus.ToString & N).Height + 60
            Dim Btn As Button
            For Each ctrl As Control In Me.fraItems.Controls
                If ctrl.Name = "cmdItemStatus" & N Then
                    Btn = ctrl
                    fraItems.Height = Btn.Top + Btn.Height
                    Exit For
                End If
            Next
            fraItems.Visible = True
            fraControls.Top = fraItems.Top + fraItems.Height
            'Height = Height - ScaleHeight + fraControls.Top + fraControls.Height + 120
            Height = Height - Me.ClientSize.Height + fraControls.Top + fraControls.Height
        End If
        For Each cc As Control In Me.fraItems.Controls
            Debug.Print(cc.Name)
        Next
    End Sub

    Private Sub GetItem(ByVal vLoc As Integer, ByVal Style As String, ByRef Loc As Double, ByRef OnOrd As Double, ByRef Avl As Double, ByRef Dsc As String, ByRef PreSold As Double)
        Dim cInv As CInvRec
        cInv = New CInvRec
        If cInv.Load(Style, "Style") Then
            Loc = cInv.QueryStock(vLoc)
            OnOrd = cInv.QueryOnOrder(vLoc)
            Avl = cInv.Available
            Dsc = cInv.Desc
            PreSold = cInv.PoSold
        End If
        DisposeDA(cInv)
    End Sub

    Public ReadOnly Property ShowST() As Boolean
        Get
            ShowST = False
        End Get
    End Property

    Private Sub HiLiteKitRow(ByVal Index As Integer)
        'Private Sub HiLiteKitRow(ByVal Currentobj As Object)
        Dim hlkrNormal As Integer, hlkrPink As Integer, hlkrCyan As Integer
        Dim T As TextBox

        hlkrNormal = -2147483633
        hlkrPink = RGB(255, 200, 200)
        hlkrCyan = RGB(128, 255, 255)

        On Error Resume Next
        If ItemStatus(Index) = "PO" Then
            'txtItemQuan(Index).BackColor = IIf(LineOverSold(Index), hlkrCyan, hlkrNormal)
            If Index = 1 Then
                'txtItemQuan.BackColor = IIf(LineOverSold(Index), hlkrCyan, hlkrNormal)
                txtItemQuan.BackColor = IIf(LineOverSold(Index), Color.Cyan, Color.White)
            ElseIf Index > 1 Then
                'Me.Controls.Item(txtItemQuan.ToString & Index).BackColor = IIf(LineOverSold(Index), hlkrCyan, hlkrNormal)

                For Each ctrl As Control In Me.fraItems.Controls
                    If ctrl.Name = "txtItemQuan" & Index Then
                        T = ctrl
                        'T.BackColor = IIf(LineOverSold(Index), hlkrCyan, hlkrNormal)
                        T.BackColor = IIf(LineOverSold(Index), Color.Cyan, Color.White)
                        Exit For
                    End If
                Next
            End If

        Else
            'txtItemQuan(Index).BackColor = IIf(LineOverSold(Index), hlkrPink, hlkrNormal)
            If Index = 1 Then
                'txtItemQuan.BackColor = IIf(LineOverSold(Index), hlkrCyan, hlkrNormal)
                txtItemQuan.BackColor = IIf(LineOverSold(Index), Color.LightPink, Color.White)
            ElseIf Index > 1 Then
                'Me.Controls.Item(txtItemQuan.ToString & Index).BackColor = IIf(LineOverSold(Index), hlkrCyan, hlkrNormal)
                For Each ctrl As Control In Me.fraItems.Controls
                    If ctrl.Name = "txtItemQuan" & Index Then
                        'T = ctrl
                        'T.BackColor = IIf(LineOverSold(Index), hlkrCyan, hlkrNormal)
                        'T.BackColor = IIf(LineOverSold(Index), Color.LightPink, Color.White)
                        'ctrl.BackColor = IIf(LineOverSold(Index), Color.LightPink, Color.White)

                        If LineOverSold(Index) = True Then
                            DirectCast(Me.fraItems.Controls.Item(ctrl.Name), TextBox).BackColor = Color.LightPink
                        Else
                            DirectCast(Me.fraItems.Controls.Item(ctrl.Name), TextBox).BackColor = Color.White
                        End If
                        Exit For
                    End If
                Next
            End If

        End If
        '  lblItemAvail(N).BackColor = txtItemQuan(N).BackColor
    End Sub

    Public Function ItemStatus(ByVal Index As Integer, Optional ByVal Vdata As String = "") As String
        Dim B As Button
        If Vdata = "" Then
            'Get property of vb6.0
            If Index <= 0 Then Exit Function
            If Index > ItemCount Then Exit Function
            'ItemStatus = cmdItemStatus(Index).Caption
            If Index = 1 Then
                ItemStatus = cmdItemStatus.Text
            Else
                'ItemStatus = Me.Controls.Item(cmdItemStatus.ToString & Index).Text
                For Each ctrl As Control In Me.fraItems.Controls
                    If ctrl.Name = "cmdItemStatus" & Index Then
                        B = ctrl
                        ItemStatus = B.Text
                        Exit For
                    End If
                Next
            End If

        Else
            'Let property of vb6.0
            If Index <= 0 Then Exit Function
            'If Index > cmdItemStatus.UBound Then Exit Property
            If Index > ItemCount Then Exit Function
            'cmdItemStatus(Index).Caption = Vdata
            If Index = 1 Then
                cmdItemStatus.Text = Vdata
            Else
                'Me.Controls.Item(cmdItemStatus.ToString & Index).Text = Vdata
                For Each ctrl As Control In Me.fraItems.Controls
                    If ctrl.Name = "cmdItemStatus" & Index Then
                        B = ctrl
                        B.Text = Vdata
                        Exit For
                    End If
                Next
            End If

        End If
    End Function

    Public Function ItemStyle(ByVal Index As Integer, Optional ByVal Vdata As String = "") As String
        Dim L As Label
        If Vdata = "" Then
            'Get property of vb6.0
            If Index <= 0 Then Exit Function
            If Index > ItemCount Then Exit Function
            'ItemStyle = lblItem(Index)
            If Index = 1 Then
                ItemStyle = lblItem.Text
            Else
                'ItemStyle = Me.Controls.Item(lblItem.ToString & Index).Text
                For Each ctrl As Control In Me.fraItems.Controls
                    If ctrl.Name = "lblItem" & Index Then
                        L = ctrl
                        ItemStyle = L.Text
                        Exit For
                    End If
                Next
            End If

        Else
            'Let property of vb6.0
            If Index <= 0 Then Exit Function
            If Index > ItemCount Then Exit Function
            'lblItem(Index).Caption = Vdata
            If Index = 1 Then
                lblItem.Text = Vdata
            Else
                'Me.Controls.Item(lblItem.ToString & Index).Text = Vdata
                For Each ctrl As Control In Me.fraItems.Controls
                    If ctrl.Name = "lblItem" & Index Then
                        L = ctrl
                        L.Text = Vdata
                        Exit For
                    End If
                Next
            End If
        End If
    End Function

    Public Function ItemLoc(ByVal Index As Integer, Optional ByVal vData As Integer = -32767)
        Dim b As Button
        If vData = -32767 Then  'Get property of vb6.0
            If Index <= 0 Then Exit Function
            If Index > ItemCount Then Exit Function
            'ItemLoc = Val(Mid(cmdItemLoc(Index).Caption, 2))
            If Index = 1 Then
                ItemLoc = Val(Mid(cmdItemLoc.Text, 2))
                Exit Function
            End If
            For Each c As Control In Me.fraItems.Controls
                If c.Name = "cmdItemLoc" & Index Then
                    b = c
                    ItemLoc = Val(Mid(b.Text, 2))
                    Exit Function
                End If
            Next
        End If

        'Let property of vb6.0
        If Index <= 0 Then Exit Function
        If Index > ItemCount Then Exit Function
        If ItemLoc(Index) = vData Then Exit Function
        'cmdItemLoc(Index).Text = "L" & vData

        If Index = 1 Then
            cmdItemLoc.Text = "L" & vData
        Else
            For Each c As Control In Me.fraItems.Controls
                If c.Name = "cmdItemLoc" & Index Then
                    b = c
                    b.Text = "L" & vData
                    Exit For
                End If
            Next
        End If
        UpdateKitRow(Index)
    End Function

    Private Sub UpdateKitRow(ByVal Line As Integer)
        Dim A As Double, B As Double, C As Double, D As String, E As Double
        If Line < 1 Or Line > ItemCount Then Exit Sub

        GetItem(ItemLoc(Line), ItemStyle(Line), A, B, C, D, E)

        'lblItemNum(Line).ToolTipText = D
        'lblItem(Line).ToolTipText = D
        ''  txtItemQuan(N) = Round(Q, 2)
        'lblItemLoc(Line) = A
        'lblOnOrd(Line) = B
        'lblOnOrd(Line).Tag = E
        'lblItemAvail(Line) = C

        'Dim LoopItemCount As Integer
        If Line = 1 Then
            ToolTip1.SetToolTip(lblItemNum, D)
            ToolTip1.SetToolTip(lblItem, D)
            lblItemLoc.Text = A
            lblOnOrd.Text = B
            lblOnOrd.Tag = E
            lblItemAvail.Text = C
        Else

            For Each ctrl As Control In Me.fraItems.Controls
                If ctrl.Name = "lblItemNum" & Line Then
                    ToolTip1.SetToolTip(ctrl, D)
                    'LoopItemCount = LoopItemCount + 1
                End If
                If ctrl.Name = "lblItem" & Line Then
                    ToolTip1.SetToolTip(ctrl, D)
                    'LoopItemCount = LoopItemCount + 1
                End If
                If ctrl.Name = "lblItemLoc" & Line Then
                    ctrl.Text = A
                    'LoopItemCount = LoopItemCount + 1
                End If
                If ctrl.Name = "lblOnOrd" & Line Then
                    ctrl.Text = B
                    ctrl.Tag = E
                    'LoopItemCount = LoopItemCount + 1
                End If
                If ctrl.Name = "lblItemAvail" & Line Then
                    ctrl.Text = C
                    'LoopItemCount = LoopItemCount + 1
                End If
                'If LoopItemCount = 5 Then
                '    Exit For
                'End If
            Next
        End If
        HiLiteKitRow(Line)
    End Sub

    Private Sub UpdateAllKitRows()
        Dim I As Integer
        'For I = lblItemNum.LBound To lblItemNum.UBound
        '   UpdateKitRow I
        'Next

        UpdateKitRow(1)  '--For physically placed lblItemNum on a form.
        For Each c As Control In Me.fraItems.Controls   '-> Dynamically added lblItemNum lables on a form.
            If c.Name = "lblItemNum" & I Then
                UpdateKitRow(I)
                I = I + 1
            End If
        Next
    End Sub

    Private Function LineOverSold(ByVal I As Integer) As Boolean
        Dim T As New TextBox
        Dim LItem As New Label, LItemLoc As New Label, LOnOrd As New Label
        'BFH20120726
        'If ItemStatus(I) = "ST" And Val(txtItemQuan(I)) + BillOSale.ItemsSoldOnSale(lblItem(I)) > Val(lblItemLoc(I)) Then LineOverSold = True
        If I = 1 Then
            If ItemStatus(I) = "ST" And Val(txtItemQuan.Text) + BillOSale.ItemsSoldOnSale(lblItem.Text) > Val(lblItemLoc.Text) Then
                LineOverSold = True
            End If
        ElseIf I > 1 Then
            For Each ctrl As Control In Me.fraItems.Controls
                If ctrl.Name = "txtItemQuan" & I Then
                    T = ctrl
                    Exit For
                End If
            Next
            For Each ctrl As Control In Me.fraItems.Controls
                If ctrl.Name = "lblItem" & I Then
                    LItem = ctrl
                    Exit For
                End If
            Next
            For Each ctrl As Control In Me.fraItems.Controls
                If ctrl.Name = "lblItemLoc" & I Then
                    LItemLoc = ctrl
                    Exit For
                End If
            Next
            If ItemStatus(I) = "ST" And Val(T.Text) + BillOSale.ItemsSoldOnSale(LItem.Text) > Val(LItemLoc.Text) Then
                LineOverSold = True
            End If
        End If

        If I = 1 Then
            'If ItemStatus(I) = "PO" And Val(txtItemQuan(I)) > Val(lblOnOrd(I)) - Val(lblOnOrd(I).Tag) Then LineOverSold = True
            If ItemStatus(I) = "PO" And Val(txtItemQuan.Text) > Val(lblOnOrd.Text) - Val(lblOnOrd.Tag) Then
                LineOverSold = True
            End If
        ElseIf I > 1 Then
            For Each ctrl As Control In Me.fraItems.Controls
                If ctrl.Name = "txtItemQuan" & I Then
                    T = ctrl
                    Exit For
                End If
            Next
            For Each ctrl As Control In Me.fraItems.Controls
                If ctrl.Name = "lblOnOrd" & I Then
                    LOnOrd = ctrl
                    Exit For
                End If
            Next
            If ItemStatus(I) = "PO" And Val(T.Text) > Val(LOnOrd.Text) - Val(LOnOrd.Tag) Then
                LineOverSold = True
            End If
        End If
    End Function

    Private Sub frmKitLevels_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage(cmdOK)
        'SetButtonImage(cmdCancel)
        SetButtonImage(cmdOK, 2)
        SetButtonImage(cmdCancel, 3)
        'SetCustomFrame(Me, ncBasicTool)  -> This is not required. This is for changing U.I. (font and color) using modNeoCaption module.

        Cancelled = False

        mLocation = 0
        AllowStatusChange = True
        AllowItemStatusChange = False
        'AllowItemLocChange = True
        AllowItemLocChange = False

        AllowPartialKits = False
        'AllowAdjustedQuantities = True ' IsDevelopment
        AllowAdjustedQuantities = False ' IsDevelopment -> Replaced this line with the above one to change true to false. Because, in vb6 form load is executing
        'before any other code. But in vb.net, this load event is executing after other code.So to work it correctly, change true to false.
        'ClearItems()  '-> Commented it because load Event Is executing In a different time than vb6.0 load Event. Because Of it caling ClearItems Is giving
        'wrong output at runtime.
    End Sub

    Private Sub txtItemQuan_TextChanged(sender As Object, e As EventArgs) Handles txtItemQuan.TextChanged
        HiLiteKitRow(1)
    End Sub

    Private Sub txtItemQuanTextChanged(sender As Object, e As EventArgs)
        Dim t As TextBox

        t = CType(sender, TextBox)
        If t.Name = "txtItemQuan" & N Then
            HiLiteKitRow(N)
        End If
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        If OverSold Then
            If MsgBox("Caution: Over Selling Kit!", vbCritical + vbOKCancel, "Warning") = vbCancel Then Exit Sub
        End If

        Cancelled = False
        If Quantity <= 0 Then Cancelled = True
        Hide()
        IsfrmKitLevelsHide = True
    End Sub

    Private Function OverSold() As Boolean
        Dim I As Integer
        For I = 1 To ItemCount
            If LineOverSold(I) Then OverSold = True : Exit Function
        Next
    End Function

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Cancelled = True
        Hide()
    End Sub

    Public ReadOnly Property Style() As String
        Get
            Style = lblStyle.Text
        End Get
    End Property
    Dim cmdStatusIndex As Integer
    Dim FromcmdStatusClick As Boolean
    Private Sub cmdStatus_Click(sender As Object, e As EventArgs) Handles cmdStatus.Click
        'cmdItemStatus_Click(0)
        'cmdStatusIndex = 0
        FromcmdStatusClick = True
        cmdItemStatus_Click(cmdItemStatus, New EventArgs)
    End Sub

    Private Sub cmdItemStatus_Click(sender As Object, e As EventArgs) Handles cmdItemStatus.Click
        Dim Stat As String, I As Integer, OS As String, Ln As Integer
        Dim b As Button
        Dim cmdItemStatusIndex As Integer = 2

        If FromcmdStatusClick = True Then
            cmdStatusIndex = 0
            FromcmdStatusClick = False
        Else
            b = CType(sender, Button)
            If b.Name = "cmdItemStatus" Then
                cmdStatusIndex = 1
            Else
                'For Each c As Control In Me.fraItems.Controls
                '    If c.Name = "cmdItemStatus" & cmdItemStatusIndex Then
                '        b = c
                '        cmdStatusIndex = cmdItemStatusIndex
                '        Exit For
                '    Else
                '        cmdItemStatusIndex = cmdItemStatusIndex + 1
                '    End If
                'Next
                Ln = Len(b.Name) - 13
                cmdStatusIndex = Microsoft.VisualBasic.Right(b.Name, Ln)
            End If
        End If

        Stat = SelectStatusPopup(IIf(cmdStatusIndex = 0, status, ItemStatus(cmdStatusIndex)))
        If Stat = "" Then Exit Sub
        If cmdStatusIndex = 0 Then
            OS = cmdStatus.Text
            cmdStatus.Text = Stat
        ElseIf cmdStatusIndex = 1 Then
            cmdItemStatus.Text = Stat
        Else
            'cmdItemStatus(Index).Caption = Stat
            b.Text = Stat
        End If

        'If Index = 0 Then
        '    For I = 1 To cmdItemStatus.UBound
        '        If cmdItemStatus(I).Caption = OS Then cmdItemStatus(I).Caption = Stat
        '    Next
        'End If

        If cmdStatusIndex = 0 Then
            If cmdItemStatus.Text = OS Then cmdItemStatus.Text = Stat
            cmdItemStatusIndex = 2
            For Each c As Control In Me.fraItems.Controls
                If c.Name = "cmdItemStatus" & cmdItemStatusIndex Then
                    If c.Text = OS Then
                        c.Text = Stat
                        cmdItemStatusIndex = cmdItemStatusIndex + 1
                    Else
                        cmdItemStatusIndex = cmdItemStatusIndex + 1
                    End If
                End If
            Next
        End If
        HiLiteKitRow(cmdStatusIndex)
    End Sub

    Private Function SelectStatusPopup(Optional ByVal Status As String = "") As String
        Dim cP As clsPopup, Res As Integer
        cP = New clsPopup
        cP.AddItem("ST", , , IIf(Status = "ST", clsPopup.enumMenuItemStates.MFS_CHECKED, 0))
        cP.AddItem("SO", , , IIf(Status = "SO", clsPopup.enumMenuItemStates.MFS_CHECKED, 0))
        cP.AddItem("PO", , , IIf(Status = "PO", clsPopup.enumMenuItemStates.MFS_CHECKED, 0))
        cP.AddItem("LAW", , , IIf(Status = "LAW", clsPopup.enumMenuItemStates.MFS_CHECKED, 0))
        cP.AddItem("DELTW", , , IIf(Status = "DELTW", clsPopup.enumMenuItemStates.MFS_CHECKED, 0))
        'Res = cP.PopupMenu(hWnd)
        Res = cP.PopupMenu(Handle)
        DisposeDA(cP)
        SelectStatusPopup = Choose(Res + 1, "", "ST", "SO", "PO", "LAW", "DELTW")
    End Function

    Private Function SelectLocationPopup(Optional ByVal L As Integer = 0, Optional ByVal Line As Integer = 0) As Integer
        Dim cP As clsPopup, I As Integer, F As String
        Dim R() As Integer, X As Integer
        cP = New clsPopup

        Dim T As CInvRec
        T = New CInvRec
        If Line <> 0 Then
            If Not T.Load(ItemStyle(Line), "Style") Then
                MsgBox("Could not load inventory record.")
                DisposeDA(T)
                Exit Function
            End If
        End If

        X = 0
        For I = 1 To ActiveNoOfLocations
            If Line <> 0 And T.QueryOnOrder(I) <= 0 Then GoTo Skip
            F = "L" & I
            If Line <> 0 Then F = F & " = " & T.QueryOnOrder(I)
            cP.AddItem(F, , , IIf(L = I, clsPopup.enumMenuItemStates.MFS_CHECKED, 0))
            X = X + 1
            'ReDim Preserve R(1 To X)
            ReDim Preserve R(0 To X - 1)
            R(X - 1) = I
Skip:
        Next
        If X = 0 Then
            MsgBox("No locations have On Order.")
            DisposeDA(T)
            Exit Function
        End If

        'SelectLocationPopup = cP.PopupMenu(hWnd)
        SelectLocationPopup = cP.PopupMenu(Handle)
        'If SelectLocationPopup <> 0 Then SelectLocationPopup = R(SelectLocationPopup)
        If SelectLocationPopup <> 0 Then SelectLocationPopup = R(SelectLocationPopup - 1)
        DisposeDA(cP, T)
    End Function

    Public Sub LoadCustomKit(ByVal vLoc As Integer, ByVal vStat As String,
Optional ByRef KI1 As String = "", Optional ByVal KQ1 As Double = 0,
Optional ByRef KI2 As String = "", Optional ByVal KQ2 As Double = 0,
Optional ByRef KI3 As String = "", Optional ByVal KQ3 As Double = 0,
Optional ByRef KI4 As String = "", Optional ByVal KQ4 As Double = 0,
Optional ByRef KI5 As String = "", Optional ByVal KQ5 As Double = 0,
Optional ByRef KI6 As String = "", Optional ByVal KQ6 As Double = 0,
Optional ByRef KI7 As String = "", Optional ByVal KQ7 As Double = 0,
Optional ByRef KI8 As String = "", Optional ByVal KQ8 As Double = 0,
Optional ByRef KI9 As String = "", Optional ByVal KQ9 As Double = 0,
Optional ByRef KI10 As String = "", Optional ByVal KQ10 As Double = 0)

        Locations = vLoc
        status = vStat
        If KI1 <> "" Then AddItem(KI1, KQ1)
        If KI2 <> "" Then AddItem(KI2, KQ2)
        If KI3 <> "" Then AddItem(KI3, KQ3)
        If KI4 <> "" Then AddItem(KI4, KQ4)
        If KI5 <> "" Then AddItem(KI5, KQ5)
        If KI6 <> "" Then AddItem(KI6, KQ6)
        If KI7 <> "" Then AddItem(KI7, KQ7)
        If KI8 <> "" Then AddItem(KI8, KQ8)
        If KI9 <> "" Then AddItem(KI9, KQ9)
        If KI10 <> "" Then AddItem(KI10, KQ10)
    End Sub

    Private Sub cmdItemLoc_Click(sender As Object, e As EventArgs) Handles cmdItemLoc.Click
        Dim Res As Integer
        Dim cmdItemLocIndex As Integer
        Dim b As Button

        b = CType(sender, Button)
        If b.Name = "cmdItemLoc" Then
            cmdItemLocIndex = 1
            Res = SelectLocationPopup(ItemLoc(cmdItemLocIndex), IIf(IsDevelopment, 1, 0))
        Else
            If Len(b.Name) = 11 Then
                cmdItemLocIndex = Microsoft.VisualBasic.Right(b.Name, 1)
            ElseIf Len(b.Name) = 12 Then
                cmdItemLocIndex = Microsoft.VisualBasic.Right(b.Name, 2)
            End If
            Res = SelectLocationPopup(ItemLoc(cmdItemLocIndex), IIf(IsDevelopment, cmdItemLocIndex, 0))
        End If

        'Res = SelectLocationPopup(ItemLoc(1), IIf(IsDevelopment, 1, 0))
        'If Res > 0 Then ItemLoc(Index) = Res
        If Res > 0 Then
            ItemLoc(cmdItemLocIndex, Res)
            'ElseIf Res > 1 Then
            '   ItemLoc(cmdItemLocIndex, Res)
        End If
    End Sub

    'Private Sub cmdItemLocNClick(sender As Object, e As EventArgs)
    '    Dim Res As Integer, cmdItemLocIndex As Integer
    '    Dim b As Button

    '    b = CType(sender, Button)
    '    If b.Name = "cmdItemLoc" Then
    '        Res = SelectLocationPopup(ItemLoc(1), IIf(IsDevelopment, 1, 0))
    '    Else
    '        If Len(b.Name) = 11 Then
    '            cmdItemLocIndex = Microsoft.VisualBasic.Right(b.Name, 1)
    '        ElseIf Len(b.Name) = 12 Then
    '            cmdItemLocIndex = Microsoft.VisualBasic.Right(b.Name, 2)
    '        End If
    '        Res = SelectLocationPopup(ItemLoc(cmdItemLocIndex), IIf(IsDevelopment, cmdItemLocIndex, 0))
    '    End If

    '    'If Res > 0 Then ItemLoc(Index) = Res
    '    If Res = 1 Then
    '        ItemLoc(1, Res)
    '    ElseIf Res > 1 Then
    '        ItemLoc(cmdItemLocIndex, Res)
    '    End If
    'End Sub

    Private Sub txtKitQuantity_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtKitQuantity.Validating
        Dim I As Integer
        If Val(txtKitQuantity.Text) <> Trunc(Val(txtKitQuantity.Text)) And Not AllowPartialKits Then
            MsgBox("No decimals allowed on kit quantities.", vbInformation, "Data Error")
            e.Cancel = True
            Exit Sub
        End If

        If Val(txtKitQuantity.Text) <= 0 Then
            MsgBox("Invalid number.  Please enter a valid number or press cancel.", vbInformation, "Data Error")
            Exit Sub
        End If

        'For I = 1 To ItemCount
        '    txtItemQuan(I) = Val(txtItemQuan(I).Tag) * Val(txtKitQuantity)
        'Next

        txtItemQuan.Text = Val(txtItemQuan.Tag) * Val(txtKitQuantity.Text)  '-> For original txtItemQuan textbox placed on a form.
        For I = 1 To ItemCount
            For Each c As Control In Me.fraItems.Controls      '-> For dynamically added txtItemQuan textboxes.
                If c.Name = "txtItemQuan" & I Then
                    'txtItemQuan(I) = Val(txtItemQuan(I).Tag) * Val(txtKitQuantity)
                    c.Text = Val(c.Tag) * Val(txtKitQuantity.Text)
                    Exit For
                End If
            Next
        Next
        UpdateAllKitRows()
        '  Reload

    End Sub


    'Private Sub tmrReload_Tick(sender As Object, e As EventArgs) Handles tmrReload.Tick
    '    tmrReload.Enabled = False
    '    LoadKit(mLocation, status, Style, Quantity)
    'End Sub

    'Private Sub Reload()
    '    tmrReload.Enabled = False
    '    tmrReload.Interval = 100
    '    tmrReload.Enabled = True
    'End Sub

End Class
