Public Class frmAutoInv
    Private mFOwner As Form, mShowDetail As Boolean
    Public Property FOwner() As Form
        Get
            FOwner = mFOwner
        End Get
        Set(value As Form)
            mFOwner = value
        End Set
    End Property
End Class