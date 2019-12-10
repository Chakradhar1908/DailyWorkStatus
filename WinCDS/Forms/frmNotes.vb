Imports VBA
Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class frmNotes
    Private NoteType As Integer, Reference As String

    Public Sub DoNotes(ByVal vNoteType As Integer, ByVal vReference As String)
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
                    AddOldNote(IfNullThenNilString(RS("NoteDate").Value), IfNullThenNilString(RS("Notes").Value))
                    RS.MoveNext()
                Loop
            Case 1  ' po notes
                'HelpContextID = 57200
                S = "SELECT * FROM PoNotes WHERE PoNo=""" & ProtectSQL(Reference) & """ ORDER BY NOTEDATE"
                RS = GetRecordsetBySQL(S, , GetDatabaseInventory)
                Do While Not RS.EOF
                    AddOldNote(IfNullThenNilString(RS("NoteDate").Value), IfNullThenNilString(RS("Notes").Value))
                    RS.MoveNext()
                Loop
        End Select

        'Show vbModal
        ShowDialog()
    End Sub

    Private Sub frmNotes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetButtonImage cmdSave, "ok"
        'SetButtonImage cmdPrint
        'SetButtonImage cmdCancel
        'SetCustomFrame Me, ncBasicTool
        SetButtonImage(cmdSave, 2)
        SetButtonImage(cmdPrint, 19)
        SetButtonImage(cmdCancel, 3)
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Printer.FontName = "Arial"
        Printer.FontSize = 14
        Printer.FontBold = True
        Select Case NoteType
            Case 0
                Printer.Print("Bill of Sale Notes")
                Printer.FontBold = False
                Printer.Print("Sale No: " & Reference)
            Case 1
                Printer.Print("Purchase Order Notes")
                Printer.FontBold = False
                Printer.Print("PO No: " & Reference)
        End Select

        Printer.Print("Date : " & Now)
        Printer.Print()
        Printer.FontSize = 12
        Printer.FontBold = False
        Printer.Print(WrapLongTextByPrintWidth(Printer, txtOldNotes.Text))
        Printer.EndDoc()
    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        On Error GoTo NoSave
        If txtNewNotes.Text <> "" Then
            Select Case NoteType
                Case 0
                    ExecuteRecordsetBySQL("INSERT INTO SALENOTES (BillOSale,Notes,NoteDate,MailIndex) VALUES (""" & ProtectSQL(Reference) & """, """ & ProtectSQL(txtNewNotes.Text) & """,#" & Now & "#, 0)", , GetDatabaseAtLocation)
                Case 1
                    ExecuteRecordsetBySQL("INSERT INTO PONOTES (PONO,Notes,NoteDate) VALUES (""" & ProtectSQL(Reference) & """, """ & ProtectSQL(txtNewNotes.Text) & """,#" & Now & "#)", , GetDatabaseInventory)
                Case Else
                    MessageBox.Show("Unknown note type: " & NoteType, "Note not saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Select
        End If
        'cmdCancel.Value = True
        cmdCancel.PerformClick()
        Exit Sub
NoSave:
        Dim X As VbMsgBoxResult
        X = MessageBox.Show("Note failed to save" & vbCrLf & Err.Description, "Note not saved", MessageBoxButtons.RetryCancel)
        'If X = vbCancel Then cmdCancel.Value = True
        If X = vbCancel Then cmdCancel.PerformClick()
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        'Unload Me
        Me.Close()
    End Sub

    Private Sub AddOldNote(ByVal D As String, ByVal N As String)
        txtOldNotes.Text = "------ " & DateFormat(D) & "  Time: " & Format(D, "h:mm:ss am/pm") & " ------" & vbCrLf & N & vbCrLf2 & txtOldNotes.Text
    End Sub

End Class