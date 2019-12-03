Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class frmOptimizeConfig

    Private Sub LoadFormValues()
        Dim R As String, C As Decimal, N As Integer
        txtStartTime.Text = GetOptimizationSetting("StartTime")
        txtTimePerStop.Text = GetOptimizationSetting("TimePerStop")
        txtCostPerHour.Text = CurrencyFormat(GetOptimizationSetting("CostPerHour"))
        txtCostPerMile.Text = CurrencyFormat(GetOptimizationSetting("CostPerMile"))
        txtTrucks.Text = GetOptimizationSetting("Trucks")
    End Sub

    Private Sub SaveFormValues()
        SetOptimizationSetting("StartTime", txtStartTime.Text)
        SetOptimizationSetting("TimePerStop", txtTimePerStop.Text)
        SetOptimizationSetting("CostPerHour", txtCostPerHour.Text)
        SetOptimizationSetting("CostPerMile", txtCostPerMile.Text)
        SetOptimizationSetting("Trucks", txtTrucks.Text)
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        SaveFormValues()
        'Unload Me
        Me.Close()
    End Sub

    Private Sub frmOptimizeConfig_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage cmdOK
        SetButtonImage(cmdOK, 2)
        LoadFormValues()
        'txtTrucks.Locked = True
        txtTrucks.ReadOnly = True
        'HelpContextID = 59650
    End Sub

    Private Sub txtStartTime_Enter(sender As Object, e As EventArgs) Handles txtStartTime.Enter
        SelectContents(txtStartTime)
    End Sub

    Private Sub txtTimePerStop_TextChanged(sender As Object, e As EventArgs) Handles txtTimePerStop.TextChanged
        txtTimePerStop.Text = "" & CLng(Val(txtTimePerStop.Text))
    End Sub

    Private Sub txtTimePerStop_Enter(sender As Object, e As EventArgs) Handles txtTimePerStop.Enter
        SelectContents(txtTimePerStop)
    End Sub

    Private Sub txtCostPerMile_Enter(sender As Object, e As EventArgs) Handles txtCostPerMile.Enter
        SelectContents(txtCostPerMile)
    End Sub

    Private Sub txtCostPerHour_Enter(sender As Object, e As EventArgs) Handles txtCostPerHour.Enter
        SelectContents(txtCostPerHour)
    End Sub

    Private Sub txtTrucks_Enter(sender As Object, e As EventArgs) Handles txtTrucks.Enter
        SelectContents(txtTrucks)
    End Sub

    Private Sub txtStartTime_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtStartTime.Validating
        If Not IsDate(txtStartTime) Then e.Cancel = True : Exit Sub
        txtStartTime.Text = Format(txtStartTime, "h:mm ampm")
    End Sub

    Private Sub txtCostPerHour_Leave(sender As Object, e As EventArgs) Handles txtCostPerHour.Leave
        txtCostPerHour.Text = CurrencyFormat(GetPrice(txtCostPerHour.Text))
    End Sub

    Private Sub txtCostPerMile_Leave(sender As Object, e As EventArgs) Handles txtCostPerMile.Leave
        txtCostPerMile.Text = CurrencyFormat(GetPrice(txtCostPerMile.Text))
    End Sub
End Class