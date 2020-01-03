Imports VBA
Imports Microsoft.VisualBasic.Interaction
Public Class OrdStatus
    Public Dimensions As String
    Public Mode As String

    '###STORECOUNT32
    'Private Const H_FRAONHAND_8 As Integer = 855
    Private Const H_FRAONHAND_8 As Integer = 50
    'Private Const H_FRAONHAND_16 As Integer = 1455
    Private Const H_FRAONHAND_16 As Integer = 95
    'Private Const H_FRAONHAND_24 As Integer = 2052
    Private Const H_FRAONHAND_24 As Integer = 140
    'Private Const H_FRAONHAND_32 As Integer = 2652
    Private Const H_FRAONHAND_32 As Integer = 210

    Private Sub OrdStatus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TotOO As String, I As Integer, X As Integer
        Dim bsStyle As String
        Dim SSCaption As String

        '###STORECOUNT32
        X = LicensedNoOfStores
        fraOnHand.Height = Switch(X > 24, H_FRAONHAND_32, X > 16, H_FRAONHAND_24, X > 8, H_FRAONHAND_16, True, H_FRAONHAND_8)
        'fraItemStatus.Top = fraOnHand.Top + fraOnHand.Height + 120
        fraItemStatus.Top = fraOnHand.Top + fraOnHand.Height + 10
        'cmdApply.Top = fraItemStatus.Top + fraItemStatus.Height + 120
        cmdApply.Top = fraItemStatus.Top + fraItemStatus.Height + 10
        cmdCancel.Top = cmdApply.Top
        'Height = cmdApply.Top + cmdApply.Height + 120 + Height - Me.ClientSize.Height
        Height = cmdApply.Top + cmdApply.Height + 10 + Height - Me.ClientSize.Height

        '  NoPO

        optReduceStock.Checked = True  'take with default
        optReduceStock_Click(optReduceStock, New EventArgs)
        StoreStock = IIf(StoreSettings.bSellFromLoginLocation, StoresSld, 1)

        For I = 1 To cOptionCount
            'StoreStockToolTipText(I) = StoreSettings(I).Address
            'StoreStockLocTip(I) = StoreSettings(I).Address

            'Note: 'In the above two properties(StoreStockToolTipText(I) and StoreStockLocTip(I)) two parameters are there in vb6.0 code.
            'In vb.net, properties will not accept multiple parameters.
            ' So created below two procedures as replacement for the above two Let properties of vb 6.0. 
            StoreStockToolTipText(I)
            StoreStockLocTip(I)
        Next

        If IsUFO() Or CheckStoreName("") Then  'Sets default option
            optLayaway.Checked = True 'LayAWay
        ElseIf IsRockyMountain() Or IsDecoratingOnADime() Then
            optTakeWith.Checked = True 'TW
        ElseIf IsParkPlace Then
            ' bfh20051031
            ' bfh20070405 - change from SpOrd to tw
            optTakeWith.Checked = True
        ElseIf IsLapeer Then
            optSpecOrd.Checked = True
        End If

        GetStore()

        If Mode = "Adj" Then Exit Sub

        ' Style has been set in BillOSale, but not much else.
        '  If True Or BillOSale.QueryStatus(BillOSale.X) <> "" Then
        'NOTE:  Must reduce incoming total by amount reserved for customers

        TotOO = BillOSale.GetTotalOnOrder

        'takes way too long to search!!! removed 6-1-00
        'ProcessTagPO  'finds open orders and reduces quantity
        bsStyle = BillOSale.QueryStyle(BillOSale.X)
        LoadOnOrder(bsStyle)
        lblTotAvail.Text = BillOSale.Rb - BillOSale.ItemsSoldOnSale(bsStyle, 0, 1)
        For I = 1 To cOptionCount
            'NOTE: THE BELOW LINE IS COMMENTED, BECAUSE the property StoreStockCaption(I) in vb6.0 has two parameters.
            'IN VB.NET PROPERTY WILL NOT ACCEPT TWO OR MORE PARAMETERS.
            'SO IT IS COMMENTED AND REPLACED WITH StoreStockCaption(I, SSCaption) PROCEDURE.
            'StoreStockCaption(I) = Left(BillOSale.GetBalance(I) - BillOSale.ItemsSoldOnSale(bsStyle, I, 1), 4)
            SSCaption = Microsoft.VisualBasic.Left(BillOSale.GetBalance(I) - BillOSale.ItemsSoldOnSale(bsStyle, I, 1), 4)
            StoreStockCaption(I, SSCaption)
        Next


        If StoreSettings.bTagIncommingDistinct Then
            Dim S As String, P As Integer, RS As ADODB.Recordset, Amt As Integer
            P = StoresSld
            TotOO = BillOSale.GetOnOrder(P)
            S = "SELECT Sum(Loc" & P & ") AS Amt FROM [Detail] WHERE Style='" & bsStyle & "' AND Trans='PO' AND Loc" & P & "<>0"
            RS = GetRecordsetBySQL(S, , GetDatabaseInventory)
            If Not RS.EOF Then Amt = IfNullThenZero(RS("Amt")) Else Amt = 0
            RS = Nothing

            ' BFH20060420 - removed the minus of 'PR's
            '      S = "SELECT Sum(Loc" & P & ") AS Amt FROM [Detail] WHERE Style='" & bsStyle & "' AND Trans='PR' AND Loc" & P & "<>0"
            '      Set RS = GetRecordsetBySQL(S, , GetDatabaseInventory)
            '      If Not RS.EOF Then Amt = Amt - IfNullThenZero(RS("Amt"))
            '      Set RS = Nothing
            '
            If Amt < 0 Then Amt = 0
            lblTagAmt.Text = TotOO - Amt - BillOSale.ItemsSoldOnSale(bsStyle, , -1)

        Else ' chkTagIncommingDistinct wasn't selected
            lblTagAmt.Text = TotOO - Val(BillOSale.PoSold) - BillOSale.ItemsSoldOnSale(bsStyle, , -1)  'total available for sale (also decrease based on current sale's PO Sold)
        End If
        '  End If

        'bfh20051111
        Quan.Text = "1"
        txtUnitPrice.Text = CurrencyFormat(BillOSale.QueryPrice(BillOSale.X))
        Dimensions = ""
        If OrdSelect.optCarpet.Checked = True Then
            'frmYardage.Show vbModal, OrdStatus
            frmYardage.ShowDialog(Me)
        End If

    End Sub

    Public Property StoreStock() As Integer
        Get
            Dim I As Integer
            'For I = optStock.LBound To optStock.UBound
            '    If optStock(I) Then StoreStock = I : Exit Property
            'Next
            For I = 1 To Setup_MaxStores
                If optStock1.Checked = True Then StoreStock = I : Exit Property
                If optStock2.Checked = True Then StoreStock = I : Exit Property
                If optStock3.Checked = True Then StoreStock = I : Exit Property
                If optStock4.Checked = True Then StoreStock = I : Exit Property
                If optStock5.Checked = True Then StoreStock = I : Exit Property
                If optStock6.Checked = True Then StoreStock = I : Exit Property
                If optStock7.Checked = True Then StoreStock = I : Exit Property
                If optStock8.Checked = True Then StoreStock = I : Exit Property
                If optStock9.Checked = True Then StoreStock = I : Exit Property
                If optStock10.Checked = True Then StoreStock = I : Exit Property
                If optStock11.Checked = True Then StoreStock = I : Exit Property
                If optStock12.Checked = True Then StoreStock = I : Exit Property
                If optStock13.Checked = True Then StoreStock = I : Exit Property
                If optStock14.Checked = True Then StoreStock = I : Exit Property
                If optStock15.Checked = True Then StoreStock = I : Exit Property
                If optStock16.Checked = True Then StoreStock = I : Exit Property
                If optStock17.Checked = True Then StoreStock = I : Exit Property
                If optStock18.Checked = True Then StoreStock = I : Exit Property
                If optStock19.Checked = True Then StoreStock = I : Exit Property
                If optStock20.Checked = True Then StoreStock = I : Exit Property
                If optStock21.Checked = True Then StoreStock = I : Exit Property
                If optStock22.Checked = True Then StoreStock = I : Exit Property
                If optStock23.Checked = True Then StoreStock = I : Exit Property
                If optStock24.Checked = True Then StoreStock = I : Exit Property
                If optStock25.Checked = True Then StoreStock = I : Exit Property
                If optStock26.Checked = True Then StoreStock = I : Exit Property
                If optStock27.Checked = True Then StoreStock = I : Exit Property
                If optStock28.Checked = True Then StoreStock = I : Exit Property
                If optStock29.Checked = True Then StoreStock = I : Exit Property
                If optStock30.Checked = True Then StoreStock = I : Exit Property
                If optStock31.Checked = True Then StoreStock = I : Exit Property
                If optStock32.Checked = True Then StoreStock = I : Exit Property

            Next
        End Get
        Set(value As Integer)
            Select Case value
                Case 1
                    optStock1.Checked = True
                Case 2
                    optStock2.Checked = True
                Case 3
                    optStock3.Checked = True
                Case 4
                    optStock4.Checked = True
                Case 5
                    optStock5.Checked = True
                Case 6
                    optStock6.Checked = True
                Case 7
                    optStock7.Checked = True
                Case 8
                    optStock8.Checked = True
                Case 9
                    optStock9.Checked = True
                Case 10
                    optStock10.Checked = True
                Case 11
                    optStock11.Checked = True
                Case 12
                    optStock12.Checked = True
                Case 13
                    optStock13.Checked = True
                Case 14
                    optStock14.Checked = True
                Case 15
                    optStock15.Checked = True
                Case 16
                    optStock16.Checked = True
                Case 17
                    optStock17.Checked = True
                Case 18
                    optStock18.Checked = True
                Case 19
                    optStock19.Checked = True
                Case 20
                    optStock20.Checked = True
                Case 21
                    optStock21.Checked = True
                Case 22
                    optStock22.Checked = True
                Case 23
                    optStock23.Checked = True
                Case 24
                    optStock24.Checked = True
                Case 25
                    optStock25.Checked = True
                Case 26
                    optStock26.Checked = True
                Case 27
                    optStock27.Checked = True
                Case 28
                    optStock28.Checked = True
                Case 29
                    optStock29.Checked = True
                Case 30
                    optStock30.Checked = True
                Case 31
                    optStock31.Checked = True
                Case 32
                    optStock32.Checked = True

            End Select

        End Set
    End Property

    Private ReadOnly Property cOptionCount() As Integer
        Get
            'cOptionCount = optStock.UBound
            cOptionCount = 32
        End Get
    End Property

    Private Sub StoreStockToolTipText(ByVal I As Integer)
        'For I = 1 To cOptionCount
        Select Case I
            Case 1
                ToolTip1.SetToolTip(optStock1, StoreSettings(I).Address)
            Case 2
                ToolTip1.SetToolTip(optStock2, StoreSettings(I).Address)
            Case 3
                ToolTip1.SetToolTip(optStock3, StoreSettings(I).Address)
            Case 4
                ToolTip1.SetToolTip(optStock4, StoreSettings(I).Address)
            Case 5
                ToolTip1.SetToolTip(optStock5, StoreSettings(I).Address)
            Case 6
                ToolTip1.SetToolTip(optStock6, StoreSettings(I).Address)
            Case 7
                ToolTip1.SetToolTip(optStock7, StoreSettings(I).Address)
            Case 8
                ToolTip1.SetToolTip(optStock8, StoreSettings(I).Address)
            Case 9
                ToolTip1.SetToolTip(optStock9, StoreSettings(I).Address)
            Case 10
                ToolTip1.SetToolTip(optStock10, StoreSettings(I).Address)
            Case 11
                ToolTip1.SetToolTip(optStock11, StoreSettings(I).Address)
            Case 12
                ToolTip1.SetToolTip(optStock12, StoreSettings(I).Address)
            Case 13
                ToolTip1.SetToolTip(optStock13, StoreSettings(I).Address)
            Case 14
                ToolTip1.SetToolTip(optStock14, StoreSettings(I).Address)
            Case 15
                ToolTip1.SetToolTip(optStock15, StoreSettings(I).Address)
            Case 16
                ToolTip1.SetToolTip(optStock16, StoreSettings(I).Address)
            Case 17
                ToolTip1.SetToolTip(optStock17, StoreSettings(I).Address)
            Case 18
                ToolTip1.SetToolTip(optStock18, StoreSettings(I).Address)
            Case 19
                ToolTip1.SetToolTip(optStock19, StoreSettings(I).Address)
            Case 20
                ToolTip1.SetToolTip(optStock20, StoreSettings(I).Address)
            Case 21
                ToolTip1.SetToolTip(optStock21, StoreSettings(I).Address)
            Case 22
                ToolTip1.SetToolTip(optStock22, StoreSettings(I).Address)
            Case 23
                ToolTip1.SetToolTip(optStock23, StoreSettings(I).Address)
            Case 24
                ToolTip1.SetToolTip(optStock24, StoreSettings(I).Address)
            Case 25
                ToolTip1.SetToolTip(optStock25, StoreSettings(I).Address)
            Case 26
                ToolTip1.SetToolTip(optStock26, StoreSettings(I).Address)
            Case 27
                ToolTip1.SetToolTip(optStock27, StoreSettings(I).Address)
            Case 28
                ToolTip1.SetToolTip(optStock28, StoreSettings(I).Address)
            Case 29
                ToolTip1.SetToolTip(optStock29, StoreSettings(I).Address)
            Case 30
                ToolTip1.SetToolTip(optStock30, StoreSettings(I).Address)
            Case 31
                ToolTip1.SetToolTip(optStock31, StoreSettings(I).Address)
            Case 32
                ToolTip1.SetToolTip(optStock32, StoreSettings(I).Address)
        End Select
        'Next
    End Sub

    Private Sub StoreStockLocTip(ByVal I As Integer)
        'For I = 1 To cOptionCount
        Select Case I
            Case 1
                ToolTip1.SetToolTip(lblLoc1, StoreSettings(I).Address)
            Case 2
                ToolTip1.SetToolTip(lblLoc2, StoreSettings(I).Address)
            Case 3
                ToolTip1.SetToolTip(lblLoc3, StoreSettings(I).Address)
            Case 4
                ToolTip1.SetToolTip(lblLoc4, StoreSettings(I).Address)
            Case 5
                ToolTip1.SetToolTip(lblLoc5, StoreSettings(I).Address)
            Case 6
                ToolTip1.SetToolTip(lblLoc6, StoreSettings(I).Address)
            Case 7
                ToolTip1.SetToolTip(lblLoc7, StoreSettings(I).Address)
            Case 8
                ToolTip1.SetToolTip(lblLoc8, StoreSettings(I).Address)
            Case 9
                ToolTip1.SetToolTip(lblLoc9, StoreSettings(I).Address)
            Case 10
                ToolTip1.SetToolTip(lblLoc10, StoreSettings(I).Address)
            Case 11
                ToolTip1.SetToolTip(lblLoc11, StoreSettings(I).Address)
            Case 12
                ToolTip1.SetToolTip(lblLoc12, StoreSettings(I).Address)
            Case 13
                ToolTip1.SetToolTip(lblLoc13, StoreSettings(I).Address)
            Case 14
                ToolTip1.SetToolTip(lblLoc14, StoreSettings(I).Address)
            Case 15
                ToolTip1.SetToolTip(lblLoc15, StoreSettings(I).Address)
            Case 16
                ToolTip1.SetToolTip(lblLoc16, StoreSettings(I).Address)
            Case 17
                ToolTip1.SetToolTip(lblLoc17, StoreSettings(I).Address)
            Case 18
                ToolTip1.SetToolTip(lblLoc18, StoreSettings(I).Address)
            Case 19
                ToolTip1.SetToolTip(lblLoc19, StoreSettings(I).Address)
            Case 20
                ToolTip1.SetToolTip(lblLoc20, StoreSettings(I).Address)
            Case 21
                ToolTip1.SetToolTip(lblLoc21, StoreSettings(I).Address)
            Case 22
                ToolTip1.SetToolTip(lblLoc22, StoreSettings(I).Address)
            Case 23
                ToolTip1.SetToolTip(lblLoc23, StoreSettings(I).Address)
            Case 24
                ToolTip1.SetToolTip(lblLoc24, StoreSettings(I).Address)
            Case 25
                ToolTip1.SetToolTip(lblLoc25, StoreSettings(I).Address)
            Case 26
                ToolTip1.SetToolTip(lblLoc26, StoreSettings(I).Address)
            Case 27
                ToolTip1.SetToolTip(lblLoc27, StoreSettings(I).Address)
            Case 28
                ToolTip1.SetToolTip(lblLoc28, StoreSettings(I).Address)
            Case 29
                ToolTip1.SetToolTip(lblLoc29, StoreSettings(I).Address)
            Case 30
                ToolTip1.SetToolTip(lblLoc30, StoreSettings(I).Address)
            Case 31
                ToolTip1.SetToolTip(lblLoc31, StoreSettings(I).Address)
            Case 32
                ToolTip1.SetToolTip(lblLoc32, StoreSettings(I).Address)
        End Select
        'Next
    End Sub

    Private Function StoreStockCaption(ByVal I As Integer, Optional ByVal Caption As String = "") As String
        'For I = 1 To cOptionCount
        '    StoreStockCaption(I) = Left(BillOSale.GetBalance(I) - BillOSale.ItemsSoldOnSale(bsStyle, I, 1), 4)
        'Next
        ' If Caption ="" Then is for Get property of vb6
        ' Else part is for Let property of vb6

        Select Case I
            Case 1
                If Caption = "" Then '-> For get property of vb6
                    StoreStockCaption = optStock1.Text
                Else                 '-> Forlet property of vb6  
                    optStock1.Text = Caption
                End If
            Case 2
                If Caption = "" Then
                    StoreStockCaption = optStock2.Text
                Else
                    optStock2.Text = Caption
                End If
            Case 3
                If Caption = "" Then
                    StoreStockCaption = optStock3.Text
                Else
                    optStock3.Text = Caption
                End If
            Case 4
                If Caption = "" Then
                    StoreStockCaption = optStock4.Text
                Else
                    optStock4.Text = Caption
                End If
            Case 5
                If Caption = "" Then
                    StoreStockCaption = optStock5.Text
                Else
                    optStock5.Text = Caption
                End If
            Case 6
                If Caption = "" Then
                    StoreStockCaption = optStock6.Text
                Else
                    optStock6.Text = Caption
                End If
            Case 7
                If Caption = "" Then
                    StoreStockCaption = optStock7.Text
                Else
                    optStock7.Text = Caption
                End If
            Case 8
                If Caption = "" Then
                    StoreStockCaption = optStock8.Text
                Else
                    optStock8.Text = Caption
                End If
            Case 9
                If Caption = "" Then
                    StoreStockCaption = optStock9.Text
                Else
                    optStock9.Text = Caption
                End If
            Case 10
                If Caption = "" Then
                    StoreStockCaption = optStock10.Text
                Else
                    optStock10.Text = Caption
                End If
            Case 11
                If Caption = "" Then
                    StoreStockCaption = optStock11.Text
                Else
                    optStock11.Text = Caption
                End If
            Case 12
                If Caption = "" Then
                    StoreStockCaption = optStock12.Text
                Else
                    optStock12.Text = Caption
                End If
            Case 13
                If Caption = "" Then
                    StoreStockCaption = optStock13.Text
                Else
                    optStock13.Text = Caption
                End If
            Case 14
                If Caption = "" Then
                    StoreStockCaption = optStock14.Text
                Else
                    optStock14.Text = Caption
                End If
            Case 15
                If Caption = "" Then
                    StoreStockCaption = optStock15.Text
                Else
                    optStock15.Text = Caption
                End If
            Case 16
                If Caption = "" Then
                    StoreStockCaption = optStock16.Text
                Else
                    optStock16.Text = Caption
                End If
            Case 17
                If Caption = "" Then
                    StoreStockCaption = optStock17.Text
                Else
                    optStock17.Text = Caption
                End If
            Case 18
                If Caption = "" Then
                    StoreStockCaption = optStock18.Text
                Else
                    optStock18.Text = Caption
                End If
            Case 19
                If Caption = "" Then
                    StoreStockCaption = optStock19.Text
                Else
                    optStock19.Text = Caption
                End If
            Case 20
                If Caption = "" Then
                    StoreStockCaption = optStock20.Text
                Else
                    optStock20.Text = Caption
                End If
            Case 21
                If Caption = "" Then
                    StoreStockCaption = optStock21.Text
                Else
                    optStock21.Text = Caption
                End If
            Case 22
                If Caption = "" Then
                    StoreStockCaption = optStock22.Text
                Else
                    optStock22.Text = Caption
                End If
            Case 23
                If Caption = "" Then
                    StoreStockCaption = optStock23.Text
                Else
                    optStock23.Text = Caption
                End If
            Case 24
                If Caption = "" Then
                    StoreStockCaption = optStock24.Text
                Else
                    optStock24.Text = Caption
                End If
            Case 25
                If Caption = "" Then
                    StoreStockCaption = optStock25.Text
                Else
                    optStock25.Text = Caption
                End If
            Case 26
                If Caption = "" Then
                    StoreStockCaption = optStock26.Text
                Else
                    optStock26.Text = Caption
                End If
            Case 27
                If Caption = "" Then
                    StoreStockCaption = optStock27.Text
                Else
                    optStock27.Text = Caption
                End If
            Case 28
                If Caption = "" Then
                    StoreStockCaption = optStock28.Text
                Else
                    optStock28.Text = Caption
                End If
            Case 29
                If Caption = "" Then
                    StoreStockCaption = optStock29.Text
                Else
                    optStock29.Text = Caption
                End If
            Case 30
                If Caption = "" Then
                    StoreStockCaption = optStock30.Text
                Else
                    optStock30.Text = Caption
                End If
            Case 31
                If Caption = "" Then
                    StoreStockCaption = optStock31.Text
                Else
                    optStock31.Text = Caption
                End If
            Case 32
                If Caption = "" Then
                    StoreStockCaption = optStock32.Text
                Else
                    optStock32.Text = Caption
                End If
        End Select
    End Function

    Private Sub GetStore()
        If Not optTagIncoming.Checked = True Then 'tag PO
            StoreStock = IIf(StoreSettings.bSellFromLoginLocation, StoresSld, 1) ' default to store sold
        End If
    End Sub

    Public Sub LoadOnOrder(ByVal ST As String)
        Dim I As Integer, C2 As CInvRec, X As Double
        cmbOnOrd.Items.Clear()
        '    cmbOnord.AddItem "OO Quantities"
        C2 = New CInvRec
        C2.Load(ST, "Style")
        For I = 1 To cOptionCount
            X = C2.QueryOnOrder(I)
            If X > 0 Then
                cmbOnOrd.Items.Add("L" & I & " - " & QuantityFormat(X))
            End If
        Next
        DisposeDA(C2)
    End Sub

    Private Sub optReduceStock_Click(sender As Object, e As EventArgs) Handles optReduceStock.Click
        'reduce stock
        'NoPO  -> Code in this sub procedure is commented in vb6.0.

        If Mode = "Adj" Then
        ElseIf OrderMode("Credit") Then
        Else
            If IsFormLoaded("BillOSale") Then BillOSale.DescEnabled = True 'added 05-21-01 for Sleep Store
        End If
        optTakeWith.TabStop = False
        optSpecOrd.TabStop = False
        optLayaway.TabStop = False

        optTakeWith.Checked = False
        optSpecOrd.Checked = False
        optLayaway.Checked = False

        FocusQuantity
    End Sub

    Private Sub FocusQuantity(Optional ByVal Always As Boolean = False)
        On Error Resume Next
        If Always Or Not SpeechActive() Then Quan.Select()
    End Sub

    Private Sub Quan_TextChanged(sender As Object, e As EventArgs) Handles Quan.TextChanged
        Dimensions = "" ' clear this if they type in a manual quantity... for frmYardage
        If Val(Quan.Text) <> 1 And OrderMode("A") Then
            txtUnitPrice.Visible = True
            lblUnitPrice.Visible = True
        End If
    End Sub

    Private Sub OrdStatus_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        'SetCustomFrame Me, ncBasicTool-------> This is line is not required. It is for U.I design using modNeoCaption module.
        If Mode <> "Adj" And Dimensions = "" Then Quan.Text = "1"

    End Sub

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        Dim I As Integer, LocAvailable As Double
        Dim ScheduledForTransfer As Double
        If Not ValidateQuantity() Then
            MsgBox("Please enter a quantity.", vbCritical, "Error")
            FocusQuantity(True)
            Exit Sub
        End If

        If Mode = "Adj" Then
            If optTagIncoming.Checked = True Then
                LocAvailable = Val(lblTagAmt.Text)
            Else
                LocAvailable = Val(StoreStockCaption(StoreStock))
            End If

            If LocAvailable - Val(Quan.Text) < 0 And Not optSpecOrd.Checked = True Then
                If MsgBox("Caution: Over Selling Item!", vbExclamation + vbOKCancel, "Warning", , , , , , False) = vbCancel Then Exit Sub
            End If
            Hide()
            Exit Sub
        End If

        If IsFormLoaded("BillOSale") Then

            ScheduledForTransfer = GetPendingTransfersFrom(BillOSale.QueryStyle(BillOSale.NewStyleLine), StoreStock)
            If ScheduledForTransfer > Val(StoreStockCaption(StoreStock)) Then
                'BFH20150117
                '"Both stores sell these items into the Red.  One store had -76 items"
                '"Can you modify so that if orig balance is less than 0 not to pop up?"
                If Val(StoreStockCaption(StoreStock)) >= 0 Then
                    If MsgBox("Location " & StoreStock & " has " & FormatQuantity(ScheduledForTransfer) & " item(s) scheduled to be transfered to other location(s)." & vbCrLf & "This may cause inventory to not be available.", vbOKCancel, "Pending Transfers") = vbCancel Then
                        FocusQuantity(True)
                        Exit Sub
                    End If
                End If
            End If



            Dim X As Integer
            X = BillOSale.NewStyleLine '.X
            'X = X + 1
            BillOSale.X = X

            If optReduceStock.Checked = True Then
                If Not CheckQuan(X) Then Exit Sub
                BillOSale.SetStatus(X, "ST")
            ElseIf optTakeWith.Checked = True Then
                If Not CheckQuan(X) Then Exit Sub
                BillOSale.SetStatus(X, "DELTW")  ' Take With Item
            ElseIf optSpecOrd.Checked = True Then
                BillOSale.SetStatus(X, "SO")
            ElseIf optLayaway.Checked = True Then
                BillOSale.SetStatus(X, "LAW")
            ElseIf optTagIncoming.Checked = True Then
                BillOSale.SetStatus(X, "PO")
            End If

            'Locations
            BillOSale.SetLoc(X, StoreStock)
            '        If BillOSale.Status = "PO" Then BillOSale.Loc = StoresSld         'removed 04-05-02

            BillOSale.SetQuan(X, Quan.Text)
            BillOSale.StatusEnabled = False
            'Note: Moved SetPrice code line from here to bottom (after below if condition) to avoid showing description in Price cell of the grid.
            If Dimensions <> "" Then
                BillOSale.SetDesc(X, "(" & Dimensions & ")  " & BillOSale.QueryDesc(X))
            End If
            BillOSale.SetPrice(X, GetPrice(txtUnitPrice.Text) * Val(Quan.Text))

            BillOSale.PriceFocus(X)
            '.CheckAddRow
            BillOSale.StyleAddEnd()
            BillOSale.KitLines = 0
            BillOSale.PriceFocus()
        End If
        'Unload OrdStatus
        Me.Close()

    End Sub

    Private Function ValidateQuantity() As Boolean
        If Not IsNumeric(Quan.Text) Then ValidateQuantity = False : Exit Function
        If GetDouble(Quan.Text) = 0 Then ValidateQuantity = False : Exit Function  ' Allow returns, at least for United.
        '  If Val(Quan) < 0 Then ValidateQuantity = False: Exit Function
        '  If CLng(Quan) <> CDbl(Quan) Then ValidateQuantity = False: Exit Function  ' Allow decimals, for carpet yardage and such.  Maybe make this product-dependent later.
        ValidateQuantity = True
    End Function

    Private Function CheckQuan(ByVal BSX As Integer) As Boolean
        Dim LocAvailable As Double, I As Integer
        CheckQuan = True
        If Mode = "Adj" Then
            Exit Function
        End If

        LocAvailable = Val(StoreStockCaption(StoreStock))

        If LocAvailable - Val(Quan.Text) < 0 Then
            ', , , , , , False
            If MsgBox("Caution: Over Selling Item!", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
                CheckQuan = False
            End If
        ElseIf LocAvailable - Val(Quan.Text) = 0 Then
            If Microsoft.VisualBasic.Left(BillOSale.QueryDesc(BSX), 3) <> "tg " Then
                BillOSale.SetDesc(BSX, "tg " & BillOSale.QueryDesc(BSX))
            Else
                ' Already marked tg, don't re-mark it.
            End If
        End If
    End Function

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        If Mode = "Adj" Then
            Quan.Text = 0
            'Unload Me
            Me.Close()
            Exit Sub
        End If

        Dim X As Integer
        X = BillOSale.X
        BillOSale.RowClear(X)
        'Unload OrdStatus
        Me.Close()
    End Sub

    Private Sub optTagIncoming_Click(sender As Object, e As EventArgs) Handles optTagIncoming.Click
        '  NoPO
        If True Or IsCranes Then
            If Val(Quan.Text) > Val(lblTagAmt.Text) Then
                Dim M As String, R As VbMsgBoxResult
                M = ""
                M = M & "There are not enough items on order for this sale!" & vbCrLf
                M = M & "Select Cancel to Special Order or Continue to oversell Incoming Stock." & vbCrLf
                M = M & vbCrLf
                M = M & "Be sure to notify Management to order item for stock!"
                R = MessageBox.Show(M, "Tag Incoming Stock Oversold", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                If R = vbCancel Then
                    optSpecOrd.Checked = True
                    Exit Sub
                End If
            End If
        End If
        FocusQuantity()
    End Sub

    Private Sub optSpecOrd_Click(sender As Object, e As EventArgs) Handles optSpecOrd.Click
        Dim X As Integer
        If Mode = "Adj" Then
        Else
            X = BillOSale.X
        End If
        'Special Order
        NoPO()

        If Mode = "Adj" Then
        Else
            BillOSale.DescEnabled = True
        End If

        optReduceStock.TabStop = False
        optTakeWith.TabStop = False
        optLayaway.TabStop = False

        optReduceStock.Checked = False
        optTakeWith.Checked = False
        optLayaway.Checked = False

        FocusQuantity()
    End Sub
    Private Sub NoPO()
        '    Dim I as integer
        '    Dim HideAndShow As Boolean, DisableAndEnable As Boolean
        '    HideAndShow = False
        '    DisableAndEnable = False
        '    If optTagIncoming.Value Then
        '        If HideAndShow Then  ' BFH20050120
        '          fraIncoming.Visible = True
        '          fraOnHand.Visible = False
        '        End If
        '        If DisableAndEnable Then
        '          For I = 1 To cOptionCount
        '            optStock(I).Enabled = False
        '            optStock(I).Value = False
        '          Next
        '        End If
        '
        '        For I = 1 To cOptionCount
        '          optOrder(I).Enabled = True
        '        Next
        '    Else
        '        If HideAndShow Then  ' BFH20050120
        '          fraIncoming.Visible = False
        '          fraOnHand.Visible = True
        '        End If
        '        If DisableAndEnable Then
        '          For I = 1 To cOptionCount
        '            optStock(I).Enabled = True
        '            optOrder(I).Enabled = False
        '          Next
        '        End If
        '          For I = 1 To cOptionCount
        '            optOrder(I).Value = False
        '          Next
        '    End If
        '    For I = 1 To cOptionCount
        '      optOrder(I).Value = False
        '    Next
        '
        '    GetStore
    End Sub

    Private Sub optTakeWith_Click(sender As Object, e As EventArgs) Handles optTakeWith.Click
        Dim X As Integer

        'bfh20050711
        ' record and reset store..
        ' easiest way to make checking this option not clear the store above, w/o finding out why..
        X = StoreStock
        'take with
        NoPO()
        StoreStock = X
        StoreStock = StoresSld ' bfh20100823 - Take With is always current store, regardless of bSellFromLoginLocation

        If Mode = "Adj" Then
        Else
            BillOSale.DescEnabled = True 'added 05-21-01 for Sleep Store
        End If
        optReduceStock.TabStop = False
        optSpecOrd.TabStop = False
        optLayaway.TabStop = False

        optReduceStock.Checked = False
        optSpecOrd.Checked = False
        optLayaway.Checked = False

        FocusQuantity()
    End Sub

    Private Sub optLayaway_Click(sender As Object, e As EventArgs) Handles optLayaway.Click
        'Lay A-Way
        NoPO()

        If Mode = "Adj" Then
        Else
            If IsFormLoaded("BillOSale") Then
                BillOSale.DescEnabled = True
            End If
        End If
        optReduceStock.TabStop = False
        optTakeWith.TabStop = False
        optSpecOrd.TabStop = False

        optReduceStock.Checked = False
        optTakeWith.Checked = False
        optSpecOrd.Checked = False

        FocusQuantity()
    End Sub
End Class