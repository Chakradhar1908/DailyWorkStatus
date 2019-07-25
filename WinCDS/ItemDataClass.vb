Public Class ItemDataClass
    Private Iname As String
    Private Itemid As Object

    Public Sub New()
        Iname = ""
        Itemid = 0
    End Sub
    Public Sub New(ByVal Name As String, ByVal Id As Object)
        Iname = Name
        Itemid = Id
    End Sub
    Public Property Itemname As String
        Get
            Return Iname
        End Get
        Set(value As String)
            Iname = value
        End Set
    End Property
    Public Property ItemData As Object
        Get
            Return Itemid
        End Get
        Set(value As Object)
            Itemid = value
        End Set
    End Property
    Public Overrides Function ToString() As String
        Return Iname
    End Function
End Class
