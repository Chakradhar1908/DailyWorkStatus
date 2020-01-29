Public Class frmLicenseAgreement
    Private Const LicenseVersion As String = "2012.12.27v12b"
    Private Const Key_LAV As String = "License Agreement Version"
    Private Const Key_LAD As String = "License Agreement Date"

    Public Function LicenseAgreement(Optional ByVal ReShow As Boolean = False) As Boolean
        LicenseAgreement = True
        If ReShow Then
            cmd2.Visible = False
            cmd0.Text = "&OK"
        Else
            If TestLicenseAgreed() Then Exit Function
            MessageBox.Show("Our license terms have changed. Please take a moment to review the new terms." & vbCrLf2 & "Click Agree to accept the terms and continue.", "License Terms Update", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

        ShowDialog()
    End Function

    Private Function TestLicenseAgreed() As Boolean
        TestLicenseAgreed = False
        TestLicenseAgreed = TestLicenseAgreed Or IsIDE()
        TestLicenseAgreed = TestLicenseAgreed Or IsCDSComputer()
        '  TestLicenseAgreed = TestLicenseAgreed Or IsDevelopment
        TestLicenseAgreed = TestLicenseAgreed Or LicenseAgreed()
    End Function

    Private Function LicenseAgreed(Optional ByVal doSet As Boolean = False) As Boolean
        If doSet Then
            SetConfigTableValue(Key_LAV, LicenseVersion)
            SetConfigTableValue(Key_LAD, Now)
        End If

        LicenseAgreed = GetConfigTableValue(Key_LAV) = LicenseVersion
    End Function

End Class