Public Class TSPNode
    Private Const Radius As Long = 2
    Private Const DIAMETER As Long = 2 * Radius + 1
    Private Const DEFAULT_DELAY_INCREMENT As Long = 15

    Public X As Long, Y As Long

    Public Address As String
    Public City As String, State As String, Zip As String

    Public StopTime As Long
    Private mWindowFrom As Date, mWindowTo As Date

    Public Visited As Boolean
    Public Name As String
    Public IsDepot As Boolean

    Public Property WindowFrom() As Date
        Get
            WindowFrom = mWindowFrom
        End Get
        Set(value As Date)
            mWindowFrom = TimeValue(value)
        End Set
    End Property

    Public Property WindowTo() As Date
        Get
            WindowTo = mWindowTo
        End Get
        Set(value As Date)
            mWindowTo = TimeValue(value)
        End Set
    End Property

    Public Sub Setup(Optional ByVal new_Name As String = "", Optional ByVal vX As Integer = 0, Optional ByVal vY As Integer = 0, Optional ByVal nStopTime As Long, Optional ByVal nWindowFrom As Date = #12:00:00 AM#, Optional ByVal nWindowTo As Date = #11:59:59 PM#, Optional ByVal nAddress As String = "", Optional ByVal nCity As String = "", Optional ByVal nSt As String = "", Optional ByVal nZip As String = "")
        Name = new_Name
        X = vX
        Y = vY
        StopTime = nStopTime
        WindowFrom = nWindowFrom
        WindowTo = nWindowTo
        Address = nAddress
        City = nCity
        State = nSt
        Zip = nZip
    End Sub

End Class
