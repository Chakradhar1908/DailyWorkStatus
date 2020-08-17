Module ImplementsModule
    Private advtype As Boolean
    Public ID as integer  ' Autonumber
    Public AdType As String
    Public OldTypeID As Integer
    Private fromemp As Boolean

    Public Property FromAdvertisingType As Boolean
        Get
            Return advtype
        End Get
        Set(value As Boolean)
            advtype = value
        End Set
    End Property

    Public Property FromEmployees As Boolean
        Get
            Return fromemp
        End Get
        Set(value As Boolean)
            fromemp = value
        End Set
    End Property
End Module
