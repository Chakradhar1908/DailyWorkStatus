Public Class frmItemView
    Private Const MAX_IMG_X As Integer = 5000
    Private Const MAX_IMG_Y As Integer = 5000
    Private Const Sp As Integer = 120

    Dim RN As Integer
    Dim Style As String
    Dim Desc As String
    Dim Comments As String
    Dim SKU As String

    Public Function PreviewStyle(ByVal Style As String, Optional ByRef frm As Form = Nothing) As Boolean
        Dim C As CInvRec
        C = New CInvRec
        If C.Load(Style, "Style") Then PreviewStyle = PreviewInvRec(C, frm)
        DisposeDA(C)
    End Function

    Public Function PreviewInvRec(ByVal Item As CInvRec, Optional ByRef frm As Form = Nothing) As Boolean
        If Item Is Nothing Then
            'Unload Me
            Me.Close()
            Exit Function
        End If
        '  If Not Item.DataAccess.Records_Available Then Exit Function
        If Item.RN = 0 Then
            'Unload Me
            Me.Close()
            Exit Function
        End If
        If ItemPXByRN(Item.RN) = "" Then
            'Unload Me
            Me.Close()
            Exit Function
        End If

        RN = Item.RN
        Style = Item.Style
        lblStyle.Text = Style
        Desc = Item.Desc
        lblDesc.Text = Desc
        Comments = Item.Comments
        SKU = Item.SKU

        'img.ToolTipText = Style & " - " & Desc
        ToolTip1.SetToolTip(img, Style & " - " & Desc)

        'img.Picture = ItemPictureByRN(RN)
        img.Image = ItemPictureByRN(RN)
        'If img.Picture = 0 Then Unload Me: Exit Function
        If img.Image Is Nothing Then Me.Close() : Exit Function
        '  img.Stretch = True
        SetSize
        On Error Resume Next
        Show()

        If Not (frm Is Nothing) Then
            'Move frm.Left + frm.Width, frm.Top
            Location = New Point(frm.Left + frm.Width, frm.Top)
        Else
            'Move Screen.Width - Width, (Screen.Height - Height) / 2
            Location = New Point(Screen.PrimaryScreen.Bounds.Width - Width, (Screen.PrimaryScreen.Bounds.Height - Height) / 2)
        End If

        PreviewInvRec = True
    End Function

    Private Sub SetSize()
        'img.Move Sp, Sp
        img.Location = New Point(Sp, Sp)

        MaintainPictureRatio(img, MAX_IMG_X, MAX_IMG_Y)
        Width = img.Width + Sp * 2
        Height = img.Height + Sp * 2 + IIf(Not ShowStyle, 0, lblStyle.Height) + IIf(Not ShowDesc, 0, lblDesc.Height)
    End Sub

    Private ReadOnly Property ShowStyle() As Boolean
        Get
            ShowStyle = True
        End Get
    End Property

    Private ReadOnly Property ShowDesc() As Boolean
        Get
            ShowDesc = True
        End Get
    End Property

End Class