Public Class frmNotes
    Private NoteType as integer, Reference As String
    Public Sub DoNotes(ByVal vNoteType as integer, ByVal vReference As String)
        Dim S As String, RS As ADODB.Recordset
        NoteType = vNoteType
        Reference = vReference

        txtOldNotes.Text = ""
        Select Case vNoteType
            Case 0  ' sale notes
                'HelpContextID = 42000
                S = "SELECT * FROM SaleNotes WHERE BillOSale=""" & ProtectSQL(Reference) & """ ORDER BY NOTEDATE"
                RS = GetRecordsetBySQL(S, , GetDatabaseAtLocation)
                Do While Not RS.EOF
                    AddOldNote(IfNullThenNilString(RS("NoteDate")), IfNullThenNilString(RS("Notes")))
                    RS.MoveNext()
                Loop
            Case 1  ' po notes
                'HelpContextID = 57200
                S = "SELECT * FROM PoNotes WHERE PoNo=""" & ProtectSQL(Reference) & """ ORDER BY NOTEDATE"
                RS = GetRecordsetBySQL(S, , GetDatabaseInventory)
                Do While Not RS.EOF
                    AddOldNote(IfNullThenNilString(RS("NoteDate")), IfNullThenNilString(RS("Notes")))
                    RS.MoveNext()
                Loop
        End Select

        'Show vbModal
        ShowDialog()
    End Sub
    Private Sub AddOldNote(ByVal D As String, ByVal N As String)
        txtOldNotes.Text = "------ " & DateFormat(D) & "  Time: " & Format(D, "h:mm:ss am/pm") & " ------" & vbCrLf & N & vbCrLf2 & txtOldNotes.Text
    End Sub

End Class