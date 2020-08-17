Public Class frmAdjustAdd
    Public Style As String

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Hide()
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Dim RN As Integer, CkIt As String, Cancelled As Boolean
        Dim Quan As Double, Status As String, Loc As String
        Dim Vendor As String

        If optEnterStyle.Checked = True Or optCarpet.Checked = True Then
            'InvCkStyle.Show vbModal
            InvCkStyle.ShowDialog()
            RN = InvCkStyle.RN
            CkIt = InvCkStyle.StyleCkIt
            Cancelled = InvCkStyle.Canceled
            Quan = 1
            If Microsoft.VisualBasic.Left(CkIt, 4) = KIT_PFX Then
                Status = InvCkStyle.KitStatus
                Quan = InvCkStyle.KitQuantity
            End If
            'Unload InvCkStyle
            InvCkStyle.Close()

            If Cancelled Then Exit Sub

            If RN <> 0 Then  ' not doing any more special stuff for SS sales
                If optCarpet.Checked = True Then
                    'Load frmYardage
                    frmYardage.Mode = "Adj"
                    'frmYardage.Show vbModal
                    frmYardage.ShowDialog()
                    Quan = frmYardage.Quantity
                    frmYardage.Mode = ""
                    'Unload frmYardage
                    frmYardage.Close()
                End If

                'Load OrdStatus
                OrdStatus.Mode = "Adj"

                If Quan <= 0 Then Quan = 1
                OrdStatus.LoadAdjStyle(RN)
                OrdStatus.Quan.Text = Str(Quan)
                'OrdStatus.Show vbModal
                OrdStatus.ShowDialog()
                Loc = OrdStatus.StoreStock
                Quan = GetPrice(OrdStatus.Quan.Text)
                Status = OrdStatus.QueryStatus
                OrdStatus.Mode = ""
                'Unload OrdStatus
                OrdStatus.Close()
            End If

            If Quan <= 0 Then Exit Sub
            If Status = "" Then Status = "#"

            OnScreenReport.AddInventory(RN, CkIt, Status, Quan, , , , Val(Loc))
        ElseIf optStain.Checked = True Then
            OnScreenReport.AddInventory(0, "STAIN")
        ElseIf optDelivery.Checked = True Then
            OnScreenReport.AddInventory(0, "DEL")
        ElseIf optLabor.Checked = True Then
            OnScreenReport.AddInventory(0, "LAB")
        ElseIf optNotes.Checked = True Then
            OnScreenReport.AddInventory(0, "NOTES")
        End If

        Hide()
    End Sub

    Private Sub frmAdjustAdd_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        On Error Resume Next
        'optEnterStyle.Checked = True
    End Sub

    Private Sub frmAdjustAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Style = ""
        'HelpContextID = 49000
        'SetCustomFrame Me, ncBasicTool
    End Sub

    Private Sub frmAdjustAdd_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'RemoveCustomFrame Me
    End Sub
End Class