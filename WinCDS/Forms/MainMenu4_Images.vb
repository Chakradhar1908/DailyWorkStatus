Public Class MainMenu4_Images
    Public Function MenuImage(ByVal mNu As String, ByVal Src As String) As Object 'As GDIpImage
        Dim I As Long
        On Error Resume Next
  Set MenuImage = imgDefault.Picture
  Select Case LCase(Left(mNu, 4))
            Case "mm"
                For I = imgMM.LBound To imgMM.UBound
                    If LCase(imgMM(I).Tag) = LCase(Src) Then Set MenuImage = imgMM(I).Picture: Exit Function
                Next
            Case "file"
                For I = imgFile.LBound To imgFile.UBound
                    If LCase(imgFile(I).Tag) = LCase(Src) Then Set MenuImage = imgFile(I).Picture: Exit Function
                Next
            Case "acco"
                For I = imgAccounting.LBound To imgAccounting.UBound
                    If LCase(imgAccounting(I).Tag) = LCase(Src) Then Set MenuImage = imgAccounting(I).Picture: Exit Function
                Next
            Case "mail"
                For I = imgMail.LBound To imgMail.UBound
                    If LCase(imgMail(I).Tag) = LCase(Src) Then Set MenuImage = imgMail(I).Picture: Exit Function
                Next
            Case "inve"
                For I = imgInventory.LBound To imgInventory.UBound
                    If LCase(imgInventory(I).Tag) = LCase(Src) Then Set MenuImage = imgInventory(I).Picture: Exit Function
                Next
            Case "orde", "serv"
                For I = imgOrder.LBound To imgOrder.UBound
                    If LCase(imgOrder(I).Tag) = LCase(Src) Then Set MenuImage = imgOrder(I).Picture: Exit Function
                Next
            Case "inst"
                For I = imgInstall.LBound To imgInstall.UBound
                    If LCase(imgInstall(I).Tag) = LCase(Src) Then Set MenuImage = imgInstall(I).Picture: Exit Function
                Next
            Case "neoc"
                For I = picResource.LBound To picResource.UBound
                    If LCase(picResource(I).Tag) = LCase(Src) Then Set MenuImage = picResource(I).Picture: Exit Function
                Next
        End Select

        If IsDevelopment() Then
            DevErr "No Image for " & mNu & "::" & Src, vbExclamation
    Debug.Print "No Image for " & mNu & "::" & Src, vbExclamation
  End If
    End Function

End Class