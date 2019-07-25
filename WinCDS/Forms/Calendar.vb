Public Class Calendar
    Private mLoadedByForm As Boolean
    Private mGridMoving As Boolean
    Private AllowMap As Boolean
    Private AllowInstr As Boolean

    Public Property LoadedByForm() As Boolean
        Get
            LoadedByForm = mLoadedByForm
        End Get
        Set(value As Boolean)
            mLoadedByForm = value
            If value Then
                cmdMenu.Text = "Back"
            Else
                cmdMenu.Text = "Menu"
            End If
        End Set
    End Property
End Class