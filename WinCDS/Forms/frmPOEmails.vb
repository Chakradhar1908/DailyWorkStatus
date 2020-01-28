Public Class frmPOEmails
    Private mSType As Long

    Public Property sType() As Long
        Get
            sType = mSType
        End Get
        Set(value As Long)
            mSType = nSType
            fraSelect.Caption = Switch(sType = 0, "Not Acknowledged:", True, "Overdue Orders:")
            dtpRunAsDate.Value = IIf(sType = 0, DateAdd("d", -10, Date), Date)
            cmdEditTemplate.Visible = mSType = 0 Or mSType = 1
            RefreshSelect
        End Set
    End Property
End Class