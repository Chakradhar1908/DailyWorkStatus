Imports stdole

Module modPictures
    Public Function ResizeAndCenterPicture(ByRef pic As PictureBox, ByRef img As IPictureDisp) As Boolean
        ' Assume image box represents maximum image size, and proper location.
        ' Resize image proportionally, and center the image in the space.
        Dim OW As Integer, Oh As Integer, pW As Integer, pH As Integer
        OW = pic.Width
        Oh = pic.Height
        pW = img.Width
        pH = img.Height


        PictureFitDimensions(pW, pH, OW, Oh, True)

        'pic.Picture = img
        pic.Image = img

        pic.Height = pH
        pic.Width = pW

        'pic.Stretch = True

        If pic.Height <> Oh Then pic.Top = pic.Top + (Oh - pic.Height) / 2
        If pic.Width <> OW Then pic.Left = pic.Left + (OW - pic.Width) / 2
    End Function
    Public Function MaintainPictureRatio(ByRef pic As PictureBox, Optional ByRef MaxX As Integer = 0, Optional ByRef MaxY As Integer = 0, Optional ByVal AdjustImage As Boolean = True)
        Dim X As Integer, Y As Integer, dx As Double, dy As Double, pW As Integer, pH As Integer
        On Error GoTo DoExit
        If MaxX = 0 Then MaxX = pic.Width
        If MaxY = 0 Then MaxY = pic.Height

        'X = pic.Picture.Width
        X = pic.Image.Width
        'Y = pic.Picture.Height
        Y = pic.Image.Height
        If X = 0 Or Y = 0 Then Exit Function

        ' e.g.
        '      view       orig      adjustment factor...
        '     ------------------------------
        '       300  /    3000   =  .1
        dx = CDbl(MaxX) / CDbl(X)
        dy = CDbl(MaxY) / CDbl(Y)

        If Y * dx > MaxY Then  ' we are only 'shrinking' this image (mathematically it should be the same).. if its too big one way, the other must fit
            pH = MaxY
            pW = X * dy
        Else
            pW = MaxX
            pH = Y * dx
        End If

        If AdjustImage Then
            pic.Height = pH
            pic.Width = pW
        End If

        MaxX = pW
        MaxY = pH
DoExit:
    End Function

    Public Function PictureFitDimensions(ByRef W As Integer, ByRef H As Integer, ByVal MaxW As Integer, ByVal MaxH As Integer, Optional ByVal Stretch As Boolean = True) As Boolean
        Dim dW As Double, dH As Double, pW As Integer, pH As Integer
        On Error GoTo DoExit

        If W = 0 Or H = 0 Then Exit Function

        ' e.g.
        '      view       orig      adjustment factor...
        '     ------------------------------
        '       300  /    3000   =  .1
        dW = CDbl(MaxW) / CDbl(W)
        dH = CDbl(MaxH) / CDbl(H)

        If H * dW > MaxH Then  ' we are only 'shrinking' this image (mathematically it should be the same).. if its too big one way, the other must fit
            pH = MaxH
            pW = W * dH
        Else
            pW = MaxW
            pH = H * dW
        End If

        If Stretch Or (pW < W And pH < H) Then
            If W <> pW Then PictureFitDimensions = True
            W = pW
            If H <> pH Then PictureFitDimensions = True
            H = pH
        End If
DoExit:
    End Function

End Module
