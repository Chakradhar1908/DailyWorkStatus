Public Class MainMenu4_Images
    Public Function MenuImage(ByVal mNu As String, ByVal Src As String) As Image 'As GDIpImage
        Dim I As Integer

        On Error Resume Next
        MenuImage = imgDefault.Image
        Select Case LCase(Microsoft.VisualBasic.Left(mNu, 4))
            Case "mm"
                'For I = imgMM.LBound To imgMM.UBound
                '    If LCase(imgMM(I).Tag) = LCase(Src) Then MenuImage = imgMM(I).Picture : Exit Function
                'Next
                For Each C As PictureBox In Me.fraMM.Controls
                    If LCase(C.Tag) = LCase(Src) Then MenuImage = C.Image : Exit Function
                Next
            Case "file"
                'For I = imgFile.LBound To imgFile.UBound
                '    If LCase(imgFile(I).Tag) = LCase(Src) Then MenuImage = imgFile(I).Picture : Exit Function
                'Next
                For Each C As PictureBox In Me.fraFile.Controls
                    If LCase(C.Tag) = LCase(Src) Then MenuImage = C.Image : Exit Function
                Next
            Case "acco"
                'For I = imgAccounting.LBound To imgAccounting.UBound
                '    If LCase(imgAccounting(I).Tag) = LCase(Src) Then MenuImage = imgAccounting(I).Picture : Exit Function
                'Next
                For Each C As PictureBox In Me.fraAccounting.Controls
                    If LCase(C.Tag) = LCase(Src) Then MenuImage = C.Image : Exit Function
                Next
            Case "mail"
                'For I = imgMail.LBound To imgMail.UBound
                '    If LCase(imgMail(I).Tag) = LCase(Src) Then MenuImage = imgMail(I).Picture : Exit Function
                'Next
                For Each C As PictureBox In Me.fraMailing.Controls
                    If LCase(C.Tag) = LCase(Src) Then MenuImage = C.Image : Exit Function
                Next
            Case "inve"
                'For I = imgInventory.LBound To imgInventory.UBound
                '    If LCase(imgInventory(I).Tag) = LCase(Src) Then MenuImage = imgInventory(I).Picture : Exit Function
                'Next
                For Each C As PictureBox In Me.fraInventory.Controls
                    If LCase(C.Tag) = LCase(Src) Then MenuImage = C.Image : Exit Function
                Next
            Case "orde", "serv"
                'For I = imgOrder10.LBound To imgOrder10.UBound
                '    If LCase(imgOrder(I).Tag) = LCase(Src) Then MenuImage = imgOrder(I).Picture : Exit Function
                'Next
                For Each C As PictureBox In Me.fraOrder.Controls
                    If LCase(C.Tag) = LCase(Src) Then MenuImage = C.Image : Exit Function
                Next
            Case "inst"
                'For I = imgInstall.LBound To imgInstall.UBound
                '    If LCase(imgInstall(I).Tag) = LCase(Src) Then MenuImage = imgInstall(I).Picture : Exit Function
                'Next
                For Each C As PictureBox In Me.fraInstallment.Controls
                    If LCase(C.Tag) = LCase(Src) Then MenuImage = C.Image : Exit Function
                Next
            Case "neoc"
                'For I = picResource.LBound To picResource.UBound
                '    If LCase(picResource(I).Tag) = LCase(Src) Then MenuImage = picResource(I).Picture : Exit Function
                'Next
                For Each C As PictureBox In Me.framCustomFrames.Controls
                    If LCase(C.Tag) = LCase(Src) Then MenuImage = C.Image : Exit Function
                Next
        End Select

        If IsDevelopment() Then
            DevErr("No Image for " & mNu & "::" & Src, vbExclamation)
            Debug.Print("No Image for " & mNu & "::" & Src, vbExclamation)
        End If
    End Function
End Class