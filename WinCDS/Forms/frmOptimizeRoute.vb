Public Class frmOptimizeRoute
    Public Network As TSPNetwork

    Public Sub LoadStops()
        Dim R As Object, I as integer, Li As ListViewItem

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
            'lvw.Items.Add(Li)
            Li = Nothing
        Next
        txtTCost.Text = CurrencyFormat(Network.BestCost)
    End Sub

    Private Sub cmdContinue_Click(sender As Object, e As EventArgs) Handles cmdContinue.Click
        Network.Solve(True)
        LoadStops()
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        'Unload frmOptimize
        'Unload Me
        frmOptimize.Close()
        Me.Close()
    End Sub

    Private Sub cmdUDClick(sender As Object, e As EventArgs) Handles cmdUD0.Click, cmdUD1.Click
        Dim A As Integer, B As Integer, D As Integer, Li As ListViewItem
        Dim Btn As Button

        On Error Resume Next
        'D = IIf(Index = 0, -1, 1)

        Btn = CType(sender, Button)
        If Btn.Name = "cmdUD0" Then
            D = -1
        Else
            D = 1
        End If

        A = -1 : B = -1
        'A = Val(lvw.SelectedItem.Text)
        For i = 0 To lvw.Items.Count - 1
            If lvw.Items(i).Selected = True Then
                A = Val(lvw.Items(i).Text)
                Exit For
            End If
        Next
        'B = Val(lvw.ListItems(lvw.SelectedItem.Index + D).Text)
        Dim PrevItem As Integer
        PrevItem = A + D

        B = Val(lvw.Items(PrevItem - 1).Text)
        If A = -1 Or B = -1 Then Exit Sub
        Network.ForceManualSwap(A, B)
        LoadStops()

        'Set Li = lvw.FindItem("" & A)
        Li = lvw.FindItemWithText("" & A)
        If Not Li Is Nothing Then Li.Selected = True
    End Sub

    Private Sub frmOptimizeRoute_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdOK, 2)
        'SetButtonImage(cmdContinue, "forward")
        SetButtonImage(cmdContinue, 1)
        SetButtonImageSmall(cmdUD0, 5)
        SetButtonImageSmall(cmdUD1, 4)
        SetAlwaysOnTop(Me)
        'HelpContextID = 59650
        'cmdContinue.Image = MainMenu.imlStandardButtons.Images(0)

        'Dim Li As New ListViewItem
        'Li.Text = 1
        'Li.SubItems.Add("one")
        'Li.SubItems.Add("two")
        'Li.SubItems.Add("three")
        'Li.SubItems.Add("four")
        'Li.SubItems.Add("five")
        'Li.SubItems.Add("six")
        'lvw.Items.Add(Li)

        'Li = New ListViewItem
        'Li.Text = 2
        'Li.SubItems.Add("seven")
        'Li.SubItems.Add("eight")
        'Li.SubItems.Add("nine")
        'Li.SubItems.Add("ten")
        'Li.SubItems.Add("eleven")
        'Li.SubItems.Add("twelve")
        'lvw.Items.Add(Li)
    End Sub

    Private Sub frmOptimizeRoute_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        frmDeliveryMap.RouteThisTruck(True)
    End Sub
End Class