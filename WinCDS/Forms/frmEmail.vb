Public Class frmEmail
    Public Mode As EmailMode
    Private Const EMAIL_SETUP_INST As String = "Goto Store Settings and enter a valid store email address."

    Public Enum EmailMode
        emSimple = 0
        emPO = 1
        emSale = 2
        emPartOrder = 3
        emChargeBack = 4
    End Enum

    Public Sub EmailSale(ByVal SaleNo As String, Optional ByVal StoreNo as integer = 0)
        Dim X As String, E As String, En As String
        Dim Cust As Boolean
        Mode = EmailMode.emSale

        Cust = True
        '  If MsgBox("Customer Copy (No Style Numbers)?", vbYesNo, "Customer Copy") = vbNo Then Cust = False
        If IsBFMyer Then Cust = False

        X = SaleToHTML(SaleNo, StoreNo, E, En, Cust)

        If Trim(E) = "" Then
            MsgBox("No email address in customer information!")
        ElseIf Trim(txtFromAddr.Text) = "" Then
            MsgBox("Store Email Address not specified." & vbCrLf & EMAIL_SETUP_INST, vbExclamation, "No Sender Email Address")
        Else
            E = SendSimpleEmail(txtFromAddr.Text, txtFromName.Text, E, En, "Sale #" & SaleNo & " - " & txtFromName.Text, X)
            MsgBox("Email Sale: " & IIf(E = "", "Success!", "FAILURE - " & E))
        End If
    End Sub



End Class