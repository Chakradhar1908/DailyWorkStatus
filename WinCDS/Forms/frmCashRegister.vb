Public Class frmCashRegister
    Public ReadOnly Property MailZip() As String
        Get
            MailZip = ""
            If MailIndex <> 0 Then
                Dim M As clsMailRec
                M = New clsMailRec
                If M.Load(frmCashRegisterAddress.MailIndex, "#Index") Then
                    MailZip = M.Zip
                End If
                DisposeDA(M)
            End If
        End Get
    End Property

    Public ReadOnly Property MailIndex() As Long
        Get
            MailIndex = Val(lblCust.Tag)
        End Get
    End Property
End Class