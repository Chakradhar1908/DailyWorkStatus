Imports Microsoft.VisualBasic.Interaction
Public Class OrdStatus
    Public Dimensions As String
    Public Mode As String

    '###STORECOUNT32
    Private Const H_FRAONHAND_8 As Long = 855
    Private Const H_FRAONHAND_16 As Long = 1455
    Private Const H_FRAONHAND_24 As Long = 2052
    Private Const H_FRAONHAND_32 As Long = 2652

    Private Sub OrdStatus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TotOO As String, I As Integer, X As Integer
        Dim bsStyle As String


        '###STORECOUNT32
        X = LicensedNoOfStores
        fraOnHand.Height = Switch(X > 24, H_FRAONHAND_32, X > 16, H_FRAONHAND_24, X > 8, H_FRAONHAND_16, True, H_FRAONHAND_8)
        fraItemStatus.Top = fraOnHand.Top + fraOnHand.Height + 120
        cmdApply.Top = fraItemStatus.Top + fraItemStatus.Height + 120
        cmdCancel.Top = cmdApply.Top
        Height = cmdApply.Top + cmdApply.Height + 120 + Height - Me.ClientSize.Height

        '  NoPO

        optReduceStock.Checked = True  'take with default
        StoreStock = IIf(StoreSettings.bSellFromLoginLocation, StoresSld, 1)

        For I = 1 To cOptionCount
            StoreStockToolTipText(I) = StoreSettings(I).Address
            StoreStockLocTip(I) = StoreSettings(I).Address
        Next

        If IsUFO() Or CheckStoreName("") Then  'Sets default option
            optLayaway = True 'LayAWay
        ElseIf IsRockyMountain() Or IsDecoratingOnADime() Then
            optTakeWith = True 'TW
        ElseIf IsParkPlace Then
            ' bfh20051031
            ' bfh20070405 - change from SpOrd to tw
            optTakeWith = True
        ElseIf IsLapeer Then
            optSpecOrd = True
        End If

        GetStore

        If Mode = "Adj" Then Exit Sub

        ' Style has been set in BillOSale, but not much else.
        '  If True Or BillOSale.QueryStatus(BillOSale.X) <> "" Then
        'NOTE:  Must reduce incoming total by amount reserved for customers

        TotOO = BillOSale.GetTotalOnOrder

        'takes way too long to search!!! removed 6-1-00
        'ProcessTagPO  'finds open orders and reduces quantity
        bsStyle = BillOSale.QueryStyle(BillOSale.X)
        LoadOnOrder bsStyle
  lblTotAvail.Caption = BillOSale.Rb - BillOSale.ItemsSoldOnSale(bsStyle, 0, 1)
        For I = 1 To cOptionCount
            StoreStockCaption(I) = Left(BillOSale.GetBalance(I) - BillOSale.ItemsSoldOnSale(bsStyle, I, 1), 4)
        Next

        If StoreSettings.bTagIncommingDistinct Then
            Dim S As String, P As Integer, RS As Recordset, Amt As Integer
            P = StoresSld
            TotOO = BillOSale.GetOnOrder(P)
            S = "SELECT Sum(Loc" & P & ") AS Amt FROM [Detail] WHERE Style='" & bsStyle & "' AND Trans='PO' AND Loc" & P & "<>0"
    Set RS = GetRecordsetBySQL(S, , GetDatabaseInventory)
    If Not RS.EOF Then Amt = IfNullThenZero(RS("Amt")) Else Amt = 0
    Set RS = Nothing
      
' BFH20060420 - removed the minus of 'PR's
'      S = "SELECT Sum(Loc" & P & ") AS Amt FROM [Detail] WHERE Style='" & bsStyle & "' AND Trans='PR' AND Loc" & P & "<>0"
'      Set RS = GetRecordsetBySQL(S, , GetDatabaseInventory)
'      If Not RS.EOF Then Amt = Amt - IfNullThenZero(RS("Amt"))
'      Set RS = Nothing
'
    If Amt < 0 Then Amt = 0
            lblTagAmt = TotOO - Amt - BillOSale.ItemsSoldOnSale(bsStyle, , -1)

        Else ' chkTagIncommingDistinct wasn't selected
            lblTagAmt = TotOO - Val(BillOSale.PoSold) - BillOSale.ItemsSoldOnSale(bsStyle, , -1)  'total available for sale (also decrease based on current sale's PO Sold)
        End If
        '  End If

        'bfh20051111
        Quan = "1"
        txtUnitPrice = CurrencyFormat(BillOSale.QueryPrice(BillOSale.X))
        Dimensions = ""
        If OrdSelect.optCarpet Then frmYardage.Show vbModal, OrdStatus

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

End Class