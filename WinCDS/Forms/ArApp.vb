Public Class ArApp
    Public Sub GetApp(Optional ByVal mR As Long = 0, Optional ByVal AN As String = "")
        mArNo = "-1"
        If AN = "" Then ArNo = ArCard.ArNo Else ArNo = AN
        If mR = 0 Then
            MailIndex = ArCard.MailRec
            mDBAccessArApp_Init MailIndex, True
  Else
            MailIndex = mR
            ArNo = "#"
            mDBAccessArApp_Init "#" & MailIndex, True
  End If
        mDBAccessArApp.GetRecord    ' this gets the record
        mDBAccessArApp.dbClose
  Set mDBAccessArApp = Nothing
    
  If mArNo = "-1" Then 'not found
            Exit Sub
        End If
    End Sub

End Class