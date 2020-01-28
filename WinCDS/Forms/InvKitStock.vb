Public Class InvKitStock
    Public Quantity As String
    Public Style As String
    Public Desc As String
    Public Landed As String
    Public List As String
    Public PackPrice As String
    Public Comments As String

    Public Sub ShowPackages()
        Show()
        GetKitInfo
    End Sub

End Class