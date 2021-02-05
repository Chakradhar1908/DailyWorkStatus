Public Class frmPrintPreviewMain
    Private Sub frmPrintPreviewMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'WindowState = vbMaximized
        'Load frmPrintPreviewDocument
        frmPrintPreviewDocument.Show()
        'Hide()
    End Sub

    Private Sub frmPrintPreviewMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        On Error Resume Next
        frmPrintPreviewDocument.CallingForm.Show
    End Sub
End Class