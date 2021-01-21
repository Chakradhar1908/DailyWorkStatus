Public Class frmPictures
    Public Enum dbPicType
        dbpty_Sales = 0
        dbpty_PO = 1
        dbpty_ServiceParts = 2
        dbpty_ServiceCalls = 1
    End Enum

    Private mType As dbPicType, mRef As String, mLoc As Integer
    Private NewIndex As Integer

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
        SetButtonImage(cmdMoveFirst, 7)
        SetButtonImage(cmdMovePrevious, 4)
        SetButtonImage(cmdMoveNext, 5)
        SetButtonImage(cmdMoveLast, 6)
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

End Class