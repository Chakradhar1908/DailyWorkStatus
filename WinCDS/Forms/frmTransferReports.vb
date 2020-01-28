Public Class frmTransferReports

    Public Property Mode() As String
        Get
            Mode = mMode
        End Get
        Set(value As String)
            mMode = vData
            Arrange
        End Set
    End Property
End Class