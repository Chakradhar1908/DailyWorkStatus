Public Class frmYardage
    Public Mode As String, Cancelled As Boolean

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Cancelled = False
        Select Case Mode
            Case "Adj"
                Hide()
            Case Else
                OrdStatus.Quan.Text = Quantity()
                OrdStatus.Dimensions = Dimensions
                'Unload Me
                Me.Close()
        End Select
    End Sub

    Public Function Quantity() As Double
        Quantity = IIf(optSqFt.Checked = True, SqFt, SqYd)
    End Function

    Public Function SqFt() As Double
        SqFt = GetDouble(txtSqFt.Text.Trim)
    End Function

    Public Function SqYd() As Double
        SqYd = GetDouble(txtSqYd.Text.Trim)
    End Function

    Public Function Units() As String
        Units = IIf(optSqFt.Checked = True, "SqFt", "SqYd")
    End Function

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Cancelled = True
        Select Case Mode
            Case "Adj"
                Hide()
            Case Else
                'Unload Me
                Me.Close()
        End Select
    End Sub

    Private Sub cmdClear_Click(sender As Object, e As EventArgs) Handles cmdClear.Click
        ClearForm
    End Sub

    Private Sub ClearForm()
        Cancelled = False
        OptSqYd.Checked = True
        txtLFt.Text = "1"
        txtLIn.Text = "0"
        txtWFt.Text = "1"
        txtWIn.Text = "0"
        Calculate
    End Sub

    Private Sub Calculate()
        Dim L As Double, W As Double
        L = GetDouble(txtLFt.Text) + GetDouble(txtLIn.Text) / 12.0#
        W = GetDouble(txtWFt.Text) + GetDouble(txtWIn.Text) / 12.0#

        txtSqFt.Text = Format(L * W, "0.0")
        txtSqYd.Text = Format((L * W) / 9.0#, "0.00")
    End Sub

    Private Sub txtLFt_TextChanged(sender As Object, e As EventArgs) Handles txtLFt.TextChanged
        Calculate()
    End Sub

    Private Sub txtLin_TextChanged(sender As Object, e As EventArgs) Handles txtLIn.TextChanged
        Calculate()
    End Sub

    Private Sub txtWFt_TextChanged(sender As Object, e As EventArgs) Handles txtWFt.TextChanged
        Calculate()
    End Sub

    Private Sub txtWIn_TextChanged(sender As Object, e As EventArgs) Handles txtWIn.TextChanged
        Calculate()
    End Sub

    Private Sub updLFt_DownClick(sender As Object, e As EventArgs) Handles updLFt.DownClick
        Adj(txtLFt, False, False)
    End Sub

    Public Function Dimensions() As String
        Dim D As String
        D = ""
        '  D = D & SqFt & " SqFt  "
        D = D & txtWFt.Text & "'"
        If txtWIn.Text <> 0 Then D = D & txtWIn.Text & """"
        D = D & " x "
        D = D & txtLFt.Text & "'"
        If txtLIn.Text <> 0 Then D = D & txtLIn.Text & """"
        Dimensions = D
    End Function

    Private Sub updLFt_UpClick(sender As Object, e As EventArgs) Handles updLFt.UpClick
        Adj(txtLFt, True, False)
    End Sub

    Private Sub updLIn_DownClick(sender As Object, e As EventArgs) Handles updLIn.DownClick
        Adj(txtLIn, False, True)
    End Sub

    Private Sub updLIn_UpClick(sender As Object, e As EventArgs) Handles updLIn.UpClick
        Adj(txtLIn, True, True)
    End Sub

    Private Sub updWFt_DownClick(sender As Object, e As EventArgs) Handles updWFt.DownClick
        Adj(txtWFt, False, False)
    End Sub

    Private Sub updWFt_UpClick(sender As Object, e As EventArgs) Handles updWFt.UpClick
        Adj(txtWFt, True, False)
    End Sub

    Private Sub updWIn_DownClick(sender As Object, e As EventArgs) Handles updWIn.DownClick
        Adj(txtWIn, False, True)
    End Sub

    Private Sub updWIn_UpClick(sender As Object, e As EventArgs) Handles updWIn.UpClick
        Adj(txtWIn, True, True)
    End Sub

    Private Sub txtLFt_Enter(sender As Object, e As EventArgs) Handles txtLFt.Enter
        SelectContents(txtLFt)
    End Sub

    Private Sub txtLIn_Enter(sender As Object, e As EventArgs) Handles txtLIn.Enter
        SelectContents(txtLIn)
    End Sub

    Private Sub txtWFt_Enter(sender As Object, e As EventArgs) Handles txtWFt.Enter
        SelectContents(txtWFt)
    End Sub

    Private Sub txtWIn_Enter(sender As Object, e As EventArgs) Handles txtWIn.Enter
        SelectContents(txtWIn)
    End Sub

    Private Sub frmYardage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage(cmdOK)
        'SetButtonImage(cmdClear)
        'SetButtonImage(cmdCancel)
        SetButtonImage(cmdOK, 2)
        SetButtonImage(cmdClear, 22)
        SetButtonImage(cmdCancel, 3)
        'SetCustomFrame Me, ncBasicTool -> This line not required. It is to change the color and font of the form and the controls.

        ClearForm()
    End Sub

    Private Sub Adj(ByRef txt As TextBox, ByVal Up As Boolean, ByVal Inches As Boolean)
        Dim T As Integer
        T = Val(txt.Text) + IIf(Up, 1, -1)
        If T < 0 Then T = 0
        If Inches And T > 11 Then T = 11
        txt.Text = CStr(T)
    End Sub

End Class