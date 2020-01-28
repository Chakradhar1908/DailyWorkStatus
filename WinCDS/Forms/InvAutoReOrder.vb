Public Class InvAutoReOrder
    Public Sub OrderAutomatic(ByVal ShowCost As Boolean)
        If Not ShowCost Then
            grdOrderItems.GetColumn(3).Visible = False
            lblTotal.Visible = False
            lblTotalOrdered.Visible = False
        End If
        cmdReset.Value = True
    End Sub
    Public Sub OrderByDemand(ByVal StoreNum As Long, ByVal StartDate As Date, ByVal EndDate As Date, ByVal ShowCost As Boolean)
        ' Set these dates in the (to be created) date boxes.
        Loading = True
        MoveGrids 1
  grdOrderItems.GetColumn(3).Visible = ShowCost
        lblTotal.Visible = ShowCost
        lblTotalOrdered.Visible = ShowCost
        cboStoreSelect.ListIndex = StoreNum
        dteStartDate.Value = StartDate
        dteEndDate.Value = EndDate
        Loading = False
        DoEvents
        cmdReset.Value = True
    End Sub

End Class