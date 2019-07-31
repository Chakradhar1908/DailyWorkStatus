Imports Microsoft.VisualBasic.Interaction
Public Class OrdStatus
    Public Dimensions As String
    Public Mode As String

    '###STORECOUNT32
    'Private Const H_FRAONHAND_8 As Integer = 855
    Private Const H_FRAONHAND_8 As Integer = 60
    'Private Const H_FRAONHAND_16 As Integer = 1455
    Private Const H_FRAONHAND_16 As Integer = 105
    'Private Const H_FRAONHAND_24 As Integer = 2052
    Private Const H_FRAONHAND_24 As Integer = 150
    'Private Const H_FRAONHAND_32 As Integer = 2652
    Private Const H_FRAONHAND_32 As Integer = 220

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
            For I = 1 To 32
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

    Private Sub StoreStockCaption(ByVal I As Integer, ByVal Caption As String)
        'For I = 1 To cOptionCount
        '    StoreStockCaption(I) = Left(BillOSale.GetBalance(I) - BillOSale.ItemsSoldOnSale(bsStyle, I, 1), 4)
        'Next
        Select Case I
            Case 1
                optStock1.Text = Caption
            Case 2
                optStock2.Text = Caption
            Case 3
                optStock3.Text = Caption
            Case 4
                optStock4.Text = Caption
            Case 5
                optStock5.Text = Caption
            Case 6
                optStock6.Text = Caption
            Case 7
                optStock7.Text = Caption
            Case 8
                optStock8.Text = Caption
            Case 9
                optStock9.Text = Caption
            Case 10
                optStock10.Text = Caption
            Case 11
                optStock11.Text = Caption
            Case 12
                optStock12.Text = Caption
            Case 13
                optStock13.Text = Caption
            Case 14
                optStock14.Text = Caption
            Case 15
                optStock15.Text = Caption
            Case 16
                optStock16.Text = Caption
            Case 17
                optStock17.Text = Caption
            Case 18
                optStock18.Text = Caption
            Case 19
                optStock19.Text = Caption
            Case 20
                optStock20.Text = Caption
            Case 21
                optStock21.Text = Caption
            Case 22
                optStock22.Text = Caption
            Case 23
                optStock23.Text = Caption
            Case 24
                optStock24.Text = Caption
            Case 25
                optStock25.Text = Caption
            Case 26
                optStock26.Text = Caption
            Case 27
                optStock27.Text = Caption
            Case 28
                optStock28.Text = Caption
            Case 29
                optStock29.Text = Caption
            Case 30
                optStock30.Text = Caption
            Case 31
                optStock31.Text = Caption
            Case 32
                optStock32.Text = Caption
        End Select
    End Sub

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
End Class