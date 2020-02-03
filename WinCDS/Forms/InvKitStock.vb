Public Class InvKitStock
    Public Quantity As String
    Public Style As String
    Public Desc As String
    Public Landed As String
    Public List As String
    Public PackPrice As String
    Public Comments As String

    Public Sub ShowPackages()
        Show()
        GetKitInfo
    End Sub

    Private Sub GetKitInfo()
        ' If mInvCkStyle is made non-modal, cleanup is required!
        ' If InvCkStyle is changed to not include this form's code, it needs to be defined in the form and withevents.
        Dim mInvCkStyle As InvCkStyle
        mInvCkStyle = New InvCkStyle
        '  mInvCkStyle.ParentForm = Me.Name
        mInvCkStyle.CallingForm = Me.Name 'Added by Robert 5/15/2017
        'mInvCkStyle.Show vbModal, Me
        mInvCkStyle.ShowDialog(Me)
        'Unload mInvCkStyle
        mInvCkStyle.Close()
    End Sub

End Class