Public Class OrdStatus
    Private Sub OrdStatus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TotOO As String, I As Integer, X As Integer
        Dim bsStyle As String


        '###STORECOUNT32
        X = LicensedNoOfStores
        fraOnHand.Height = Switch(X > 24, H_FRAONHAND_32, X > 16, H_FRAONHAND_24, X > 8, H_FRAONHAND_16, True, H_FRAONHAND_8)
        fraItemStatus.Top = fraOnHand.Top + fraOnHand.Height + 120
        cmdApply.Top = fraItemStatus.Top + fraItemStatus.Height + 120
        cmdCancel.Top = cmdApply.Top
        Height = cmdApply.Top + cmdApply.Height + 120 + Height - ScaleHeight

        '  NoPO

        optReduceStock.Value = True  'take with default
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
End Class