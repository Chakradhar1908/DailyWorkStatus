Public Class frmDemoNotify
    Public Function RequireLicenseOrQuit() As Boolean
        'Show 1
        ShowDialog()
        If LicenseValid(txtEnterLicense.Text) Then
            License = txtEnterLicense.Text
            RequireLicenseOrQuit = True
            Exit Function
        End If
        End
    End Function

End Class