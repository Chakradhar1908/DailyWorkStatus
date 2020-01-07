Public Class frmLicenseAgreement
    Public Function LicenseAgreement(Optional ByVal ReShow As Boolean = False) As Boolean
        LicenseAgreement = True
        If ReShow Then
            cmd(2).Visible = False
            cmd(0).Caption = "&OK"
        Else
            If TestLicenseAgreed Then Exit Function
            MsgBox "Our license terms have changed. Please take a moment to review the new terms." & vbCrLf2 & "Click Agree to accept the terms and continue.", vbExclamation, "License Terms Update"
  End If

        Show vbModal
End Function

End Class