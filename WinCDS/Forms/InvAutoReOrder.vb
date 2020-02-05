Public Class InvAutoReOrder
    Private Loading As Boolean
    Private MarginItems() As PotentialOrder
    Private LastStartDate As Date, LastEndDate As Date

    Private Structure PotentialOrder
        Dim StyleIndex as integer   ' Internal pointer
        Dim SaleNo As String
        Dim Style As String
        Dim Vendor As String
        Dim RN as integer
        Dim Name As String
        Dim MarginLine as integer
        Dim Quantity As Double
        Dim LastPayDate As Date
        Dim PickupDate As Date
        Dim Deposit As Decimal
        Dim Balance As Decimal
        Dim OrderMe As Boolean
        Dim Index as integer
        Dim StoreNo as integer '  20050111BFH/MJK
    End Structure

    Public Sub OrderAutomatic(ByVal ShowCost As Boolean)
        If Not ShowCost Then
            'grdOrderItems.GetColumn(3).Visible = False
            lblTotal.Visible = False
            lblTotalOrdered.Visible = False
        End If
        'cmdReset.Value = True
        cmdReset.PerformClick()
    End Sub

    Public Sub OrderByDemand(ByVal StoreNum as integer, ByVal StartDate As Date, ByVal EndDate As Date, ByVal ShowCost As Boolean)
        ' Set these dates in the (to be created) date boxes.
        Loading = True
        MoveGrids(1)
        'grdOrderItems.GetColumn(3).Visible = ShowCost
        lblTotal.Visible = ShowCost
        lblTotalOrdered.Visible = ShowCost
        cboStoreSelect.SelectedIndex = StoreNum
        dteStartDate.Value = StartDate
        dteEndDate.Value = EndDate
        Loading = False
        Application.DoEvents()
        'cmdReset.Value = True
        cmdReset.PerformClick()
    End Sub

    Private Sub MoveGrids(ByVal Pos as integer)
        ' 0: All Items.
        ' 1: Sales on top, items on bottom.
        lblVendor.Visible = True
        cboVendors.Visible = True

        lblDept.Visible = (Pos = 0)
        cboDept.Visible = (Pos = 0)
        fraDemand.Visible = (Pos <> 0)

        If Pos = 0 Then
            'grdSaleItems.Visible = False
            'grdOrderItems.Height = grdSaleItems.Top - grdOrderItems.Top + grdSaleItems.Height
            'grdOrderItems.Move grdOrderItems.Left, grdSaleItems.Top, grdOrderItems.Width, grdOrderItems.GetDBGrid.RowHeight * 23 ' grdOrderItems.Top - grdSaleItems.Top + grdOrderItems.Height
        Else
            'grdSaleItems.Visible = True
            'grdOrderItems.Move grdOrderItems.Left, grdSaleItems.Top + grdSaleItems.Height + 60,
            'grdOrderItems.Width, grdOrderItems.GetDBGrid.RowHeight * 10 ' grdOrderItems.Height - grdSaleItems.Height - 60
            lblVendor.Top = 6840
            cboVendors.Top = 6720
        End If
    End Sub

End Class