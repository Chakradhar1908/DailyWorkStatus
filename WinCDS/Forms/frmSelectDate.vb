Public Class frmSelectDate
    Public Result As String

    Public Function SelectDate(Optional ByVal Def As String = NullDateString, Optional ByVal vCaption As String = "Select Date:") As String
        Result = ""
        Text = vCaption
        If IsDate(Def) Then dtp.Value = Def Else dtp.Value = Today
        'Show 1
        ShowDialog()
        Result = dtp.Value
        SelectDate = Result
        'Unload Me
        Me.Close()
    End Function
End Class