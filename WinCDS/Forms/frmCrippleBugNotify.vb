Public Class frmCrippleBugNotify
    Public Sub CrippleBug(Optional ByVal Feature As String = "")
        Dim S As String
        If Feature <> "" Then
            S = ""
            S = S & Feature & " has been disabled." & vbCrLf
            S = S & "The software has expired and currently only allows you to view sales history." & vbCrLf
            S = S & vbCrLf
            S = S & AdminContactString(WebSite:=True, Email:=True, Phone2:=True, Phone:=False, Version:=False, Company:=False)
        Else
            S = ""
            S = S & "If you receive this message, your software has expired. " & vbCrLf2
            S = S & "The software will continue to allow to operate in a minimal operation, "
            S = S & "allowing you to view past sales and reports.  New sales and other changes "
            S = S & "will be disabled."
            S = S & vbCrLf2
            S = S & "To get the current updates or reinitiate your service contract, please contact: " & AdminContactName & " at " & AdminContactPhone2 & "." & vbCrLf
            S = S & vbCrLf
            S = S & AdminContactString(WebSite:=True, Email:=True, Phone2:=True, Phone:=False, Version:=False, Company:=False)
        End If
        lblNotice.Text = S
        'Show 1
        ShowDialog()
    End Sub

End Class