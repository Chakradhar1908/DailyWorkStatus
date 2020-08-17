Public Class InvDefault
    Public X As Integer
    Public Style As String
    Public I As Integer
    Private MeDotClose As Boolean

    'Public mShown As Boolean
    Public Enum ENoStyle
        eNoStyle_NotSet
        eNoStyle_ReEnter
        eNoStyle_EnterItem
        eNoStyle_NotInDBase
        eNoStyle_LayAway
        Unload_BillOfSale
    End Enum

    Private LoadingMfgs As Boolean
    Public mOptionSelected As ENoStyle
    'Private Const FRMW_1 as integer = 2985
    Private Const FRMW_1 As Integer = 190
    Private Const FRMW_2 As Integer = 6100

    Public Function ShowAndTell() As ENoStyle
        If IsChandlers Then optEnterNotInInv.Checked = True
        'Show vbModal, BillOSale
        ShowDialog(BillOSale)
        ShowAndTell = mOptionSelected
        'Unload Me
        Me.Close()
    End Function

    Private Sub InvDefault_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Must be left rem for tab off style to work
        Width = FRMW_1
        lstResults.Items.Clear()
        'SetButtonImage(cmdApply)
        SetButtonImage(cmdApply, 2)
    End Sub

    Private Sub InvDefault_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        On Error Resume Next
        lstResults.Width = Width - lstResults.Left - 30
        'lstResults.Height = ScaleHeight - 120
        lstResults.Height = Me.ClientSize.Height - 10

    End Sub

    Private Sub Form_Activate()
        'No form activate event in vb.net. This is not required.
        'SetCustomFrame Me, ncBasicTool
    End Sub

    Private Sub InvDefault_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'Note: This event is replacement for unload and queryunload events of vb6.0
        'If UnloadMode = vbFormControlMenu Then
        '    optReEnter.Value = True
        '    cmdApply.Value = True
        'End If

        If e.CloseReason = CloseReason.UserClosing And MeDotClose = False Then
            optReEnter.Checked = True
            cmdApply.PerformClick()
        End If
    End Sub

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        mOptionSelected = ENoStyle.eNoStyle_NotSet
        If optReEnter.Checked = True Then mOptionSelected = ENoStyle.eNoStyle_ReEnter
        If optEnterNotInInv.Checked = True Then mOptionSelected = ENoStyle.eNoStyle_EnterItem
        If optSONotCarried.Checked = True Then mOptionSelected = ENoStyle.eNoStyle_NotInDBase
        If optSOLawNotCarried.Checked = True Then mOptionSelected = ENoStyle.eNoStyle_LayAway

        '    If (optSONotCarried Or optSOLawNotCarried) And Trim(BillOSale.QueryMfg(BillOSale.NewStyleLine)) = "" Then
        '      MsgBox "Please select a vendor.", vbExclamation
        '      Exit Sub
        '    End If

        If mOptionSelected = ENoStyle.eNoStyle_NotSet Then
            MsgBox("Option Not Selected!", vbExclamation)
            Exit Sub
        Else
            'BillOSale.Mfg.SetFocus
            If mOptionSelected <> ENoStyle.eNoStyle_ReEnter Then BillOSale.StyleAddEnd()
            MeDotClose = True
            'Unload Me
            Me.Close()
            'Me.Hide()
        End If

    End Sub


End Class