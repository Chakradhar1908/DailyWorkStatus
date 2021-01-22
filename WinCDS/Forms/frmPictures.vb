Imports VBRUN
Public Class frmPictures
    Public Enum dbPicType
        dbpty_Sales = 0
        dbpty_PO = 1
        dbpty_ServiceParts = 2
        dbpty_ServiceCalls = 1
    End Enum

    Private mType As dbPicType, mRef As String, mLoc As Integer
    Private NewIndex As Integer

    Private Function StartAdd() As Integer
        Dim SQL As String, RS As ADODB.Recordset
        SQL = "INSERT INTO [Pictures] (PictureType, PictureRef, Caption) VALUES (" & mType & ", '" & ProtectSQL(mRef) & "', '')"
        ExecuteRecordsetBySQL(SQL, , GetDatabaseAtLocation(mLoc))
        SQL = "SELECT TOP 1 PictureID FROM [Pictures] ORDER BY PictureID DESC"
        RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation(mLoc))
        If Not RS Is Nothing Then
            If Not RS.EOF Then
                StartAdd = RS("PictureID").Value
            End If
        End If
        RS = Nothing
    End Function

    Public Sub LoadPicturesByRef(ByVal pType As dbPicType, ByVal pRef As String, Optional ByVal pLoc As String = "0")
        mType = pType
        mRef = pRef
        If pLoc = 0 Then pLoc = StoresSld
        mLoc = pLoc

        UpdateCaption()
        UpdateData()
        'Show vbModal
        ShowDialog()
    End Sub

    Private Sub frmPictures_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdOK, 2)
        SetButtonImage(cmdDelete, 10)
        SetButtonImage(cmdAdd, 0)
        SetButtonImage(cmdPrint, 19)

        SetButtonImage(cmdMoveFirst, 7)
        SetButtonImage(cmdMovePrevious, 4)
        SetButtonImage(cmdMoveNext, 5)
        SetButtonImage(cmdMoveLast, 6)
    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        'On Error Resume Next
        If cmdAdd.Text = "&Add" Then
            NewIndex = StartAdd()
            UpdateData(ID:=NewIndex)
            'datPictures.Caption = "(Drag image into the box or Dbl Click)"
            lblPictures.Text = "(Drag image into the box or Dbl Click)"
            cmdDelete.Visible = False
            cmdPrint.Visible = False
            cmdAdd.Text = "S&ave"
            cmdOK.Text = "Cance&l"
            SetButtonImage(cmdOK, 2)
        Else
            cmdDelete.Visible = True
            cmdPrint.Visible = True
            cmdAdd.Text = "&Add"
            cmdOK.Text = "&OK"
            SetButtonImage(cmdOK, 2)
            NewIndex = 0
            UpdateData()
        End If

    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        If Val(txtPictureID.Text) <> 0 Then
            If MessageBox.Show("Delete this image?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                DeleteRecord(Val(txtPictureID.Text))
                UpdateData()
            End If
        Else
            MessageBox.Show("Nothing to delete.", "Can't Delete", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub DeleteRecord(ByVal ID As Integer)
        'datPictures.DataBase.Close 'Note: Not using DataControl named datPictures.
        ExecuteRecordsetBySQL("DELETE FROM [Pictures] WHERE PictureID=" & Val(txtPictureID.Text))
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        If NewIndex <> 0 Then
            DeleteRecord(NewIndex)
            NewIndex = 0
        End If
        'Unload Me
        Me.Close()
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Dim X As Integer, Y As Integer
        X = 6000
        Y = 6000
        MaintainPictureRatio(imgPicture, X, Y, False)
        Printer.PaintPicture(imgPicture.Image, (Printer.ScaleWidth - X) / 2, 1000, X, Y)
        Printer.FontName = "Arial"
        Printer.FontSize = 12
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(lblRef.Text)) / 2
        Printer.CurrentY = 700
        Printer.Print(lblRef.Text)
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(txtCaption.Text)) / 2
        Printer.CurrentY = 1000 + Y + 200
        Printer.Print(txtCaption.Text)
        Printer.EndDoc()
    End Sub

    Private Sub UpdateCaption()
        Dim T As String
        Select Case mType
            Case dbPicType.dbpty_Sales : T = "Sale #"
            Case dbPicType.dbpty_PO : T = "PO #"
            Case dbPicType.dbpty_ServiceParts : T = "Service Part Ord #"
            Case dbPicType.dbpty_ServiceCalls : T = "Service Call #"
        End Select
        lblRef.Text = T & mRef
    End Sub

    Private Sub frmPictures_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        'ReceivePictureDrop Data, Effect, Button, Shift, X, Y
        ReceivePictureDrop(e.Data, e.Effect, e.KeyState, e.KeyState, e.X, e.Y)
    End Sub

    Private Sub frmPictures_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

        If Width < 225 Then Width = 225
        If Height < 295 Then Height = 295
        fraPicture.Width = Me.ClientSize.Width - 2 * fraPicture.Left '312,295
        fraPicture.Height = Me.ClientSize.Height - 2 * fraPicture.Top - cmdDelete.Height - 12
        'datPictures.Top = fraPicture.Height - 375
        lblPictures.Top = fraPicture.Height - 30 '266
        txtCaption.Top = fraPicture.Height - 96
        lblCaption.Top = fraPicture.Height - 117
        lblPictureID.Top = lblCaption.Top
        lblPictureType.Top = lblCaption.Top
        lblPictureRef.Top = lblCaption.Top
        'txtPictureID.Top = fraPicture.Height - 139
        'txtPictureType.Top = txtPictureID.Top
        'txtPictureRef.Top = txtPictureID.Top
        txtPictureID.Top = lblCaption.Top
        txtPictureType.Top = lblCaption.Top
        txtPictureRef.Top = lblCaption.Top

        lblRef.Width = fraPicture.Width - 2 * txtCaption.Left
        txtCaption.Width = lblRef.Width
        'datPictures.Width = lblRef.Width
        'lblPictures.Width = lblRef.Width

        imgPicture.Width = lblRef.Width
        imgPicture.Height = fraPicture.Height - 170

        cmdDelete.Top = fraPicture.Top + fraPicture.Height + 3
        cmdAdd.Top = cmdDelete.Top
        cmdPrint.Top = cmdDelete.Top
        cmdOK.Top = cmdDelete.Top

        cmdDelete.Left = Me.ClientSize.Width - (cmdDelete.Width * 4 + 11 * 3)  '329-317
        cmdAdd.Left = cmdDelete.Left + (cmdDelete.Width + 8) '91
        cmdPrint.Left = cmdAdd.Left + (cmdDelete.Width + 8) '170
        cmdOK.Left = cmdPrint.Left + (cmdDelete.Width + 8) '249

        cmdMoveFirst.Top = cmdMoveFirst.Top - 10
        cmdMoveNext.Top = cmdMoveNext.Top - 10
        cmdMovePrevious.Top = cmdMovePrevious.Top - 10
        cmdMoveLast.Top = cmdMoveLast.Top - 10
    End Sub

    Private Sub imgPicture_DoubleClick(sender As Object, e As EventArgs) Handles imgPicture.DoubleClick
        On Error GoTo Canceled
        cdgFile.Filter = "(Image files)|*.bmp;*.jpg;*.png;*.gif|(All Files)|*.*"
        cdgFile.FilterIndex = 1
        cdgFile.CancelError = True
        cdgFile.ShowOpen()
        imgPicture.Image = LoadPictureStd(cdgFile.FileName)
Canceled:
    End Sub

    Private Sub imgPicture_DragDrop(sender As Object, e As DragEventArgs) Handles imgPicture.DragDrop
        'ReceivePictureDrop Data, Effect, Button, Shift, X, Y
        ReceivePictureDrop(e.Data, e.Effect, e.KeyState, e.KeyState, e.X, e.Y)
    End Sub

    Private Sub imgReveal_DoubleClick(sender As Object, e As EventArgs) Handles imgReveal.DoubleClick
        Dim X As Boolean
        X = Not txtPictureID.Visible
        lblPictureID.Visible = X
        txtPictureID.Visible = X
        lblPictureType.Visible = X
        txtPictureType.Visible = X
        lblPictureRef.Visible = X
        txtPictureRef.Visible = X
    End Sub

    Private Sub UpdateData(Optional ByVal Pos As Integer = 0, Optional ByVal ID As Integer = 0)
        On Error Resume Next
        'datPictures.DatabaseName = GetDatabaseAtLocation(mLoc)
        'If ID <> 0 Then
        '    datPictures.RecordSource = "SELECT * FROM [Pictures] WHERE PictureID=" & ID
        'Else
        '    datPictures.RecordSource = "SELECT * FROM [Pictures] WHERE PictureType=" & mType & " AND PictureRef='" & ProtectSQL(mRef) & "' ORDER BY PictureID"
        'End If
        'datPictures.Refresh
        'If Pos <> 0 Then
        '    If Pos > datPictures.Recordset.RecordCount Then Pos = datPictures.Recordset.RecordCount
        '    datPictures.Recordset.AbsolutePosition = Pos
        '    datPictures.Refresh
        'End If

        'On Error Resume Next
        'DoEvents
        'datPictures.Recordset.MoveLast
        'datPictures.Recordset.MoveFirst
        'If datPictures.Recordset.RecordCount = 0 Then
        '    datPictures.Caption = "No Pictures (Click Add)"
        'Else
        '    datPictures.Caption = (datPictures.Recordset.AbsolutePosition + 1) & " of " & datPictures.Recordset.RecordCount
        'End If

        Dim Rs As ADODB.Recordset

        If ID <> 0 Then
            Rs = GetRecordsetBySQL("SELECT * FROM [Pictures] WHERE PictureID=" & ID,, GetDatabaseAtLocation(mLoc))
        Else
            Rs = GetRecordsetBySQL("SELECT * FROM [Pictures] WHERE PictureType=" & mType & " AND PictureRef='" & ProtectSQL(mRef) & "' ORDER BY PictureID",, GetDatabaseAtLocation(mLoc))
        End If

        If Pos <> 0 Then
            'If Pos > datPictures.Recordset.RecordCount Then Pos = datPictures.Recordset.RecordCount
            If Pos > Rs.RecordCount Then Pos = Rs.RecordCount
            'datPictures.Recordset.AbsolutePosition = Pos
            Rs.AbsolutePosition = Pos
            'datPictures.Refresh
        End If
        Rs.MoveLast()
        Rs.MoveFirst()

        'If datPictures.Recordset.RecordCount = 0 Then
        '    datPictures.Caption = "No Pictures (Click Add)"
        'Else
        '    datPictures.Caption = (datPictures.Recordset.AbsolutePosition + 1) & " of " & datPictures.Recordset.RecordCount
        'End If
        If Rs.RecordCount = 0 Then
            lblPictures.Text = "No Pictures (Click Add)"
        Else
            lblPictures.Text = Rs.AbsolutePosition + 1 & " of " & Rs.RecordCount
        End If
    End Sub

    Private Sub ReceivePictureDrop(ByRef Data As DataObject, ByRef Effect As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
        '  DetectDataObjectFormat Data

        If Data.GetFormat(15) Then
            On Error Resume Next
            'If Data.Files.Count >= 1 Then imgPicture.Picture = LoadPictureStd(Data.Files(1))
            If Data.Files.Count >= 1 Then imgPicture.Image = LoadPictureStd(Data.Files(1))
        ElseIf Data.GetFormat(2) Then
            'imgPicture.Picture = Data.GetData(2)
            imgPicture.Image = Data.GetData(2)
        End If
    End Sub

End Class