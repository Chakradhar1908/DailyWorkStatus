Public Class frmOptimizeRoute
    Public Network As TSPNetwork

    Public Sub LoadStops()
        Dim R As Object, I As Long, Li As ListViewItem

        R = Network.GetResultSet
        If IsNothing(R) Then Exit Sub
        lvw.View = View.Details
        lvw.Items.Clear()

        For I = LBound(R, 1) To UBound(R, 1)
            'Li = lvw.ListItems.Add(, , R(I, tspRS_ID))
            Li = lvw.Items.Add(R(I, tspRS.tspRS_ID))
            'Li.SubItems(1) = R(I, tspRS.tspRS_Name)
            Li.SubItems.Add(R(I, tspRS.tspRS_Name))
            'Li.SubItems(2) = R(I, tspRS.tspRS_X)
            Li.SubItems.Add(R(I, tspRS.tspRS_X))
            'Li.SubItems(3) = R(I, tspRS.tspRS_Y)
            Li.SubItems.Add(R(I, tspRS.tspRS_Y))
            'Li.SubItems(4) = R(I, tspRS.tspRS_WindowFrom)
            Li.SubItems.Add(R(I, tspRS.tspRS_WindowFrom))
            'Li.SubItems(5) = R(I, tspRS.tspRS_WindowTo)
            Li.SubItems.Add(R(I, tspRS.tspRS_WindowTo))
            'Li.SubItems(6) = R(I, tspRS.tspRS_Distance)
            Li.SubItems.Add(R(I, tspRS.tspRS_Distance))
            'Li.SubItems(7) = R(I, tspRS.tspRS_Delay)
            Li.SubItems.Add(R(I, tspRS.tspRS_Delay))
            'Li.SubItems(8) = R(I, tspRS.tspRS_Arrive)
            Li.SubItems.Add(R(I, tspRS.tspRS_Arrive))
            'Li.SubItems(9) = R(I, tspRS.tspRS_StopTime)
            Li.SubItems.Add(R(I, tspRS.tspRS_StopTime))
            'Li.SubItems(10) = R(I, tspRS.tspRS_Depart)
            Li.SubItems.Add(R(I, tspRS.tspRS_Depart))
            'Li.SubItems(11) = R(I, tspRS.tspRS_Address)
            Li.SubItems.Add(R(I, tspRS.tspRS_Address))
            'Li.SubItems(12) = R(I, tspRS.tspRS_City)
            Li.SubItems.Add(R(I, tspRS.tspRS_City))
            'Li.SubItems(13) = R(I, tspRS.tspRS_State)
            Li.SubItems.Add(R(I, tspRS.tspRS_State))
            '    LI.SubItems(13) = R(I, tspRS_Zip)
            Li = Nothing
        Next
        txtTCost.Text = CurrencyFormat(Network.BestCost)
    End Sub

End Class