Public Class frmAutoInv
    Public Property FOwner() As Form
        Get
            FOwner = mFOwner
        End Get
        Set(value As Form)
            mFOwner = vF
        End Set
    End Property
End Class