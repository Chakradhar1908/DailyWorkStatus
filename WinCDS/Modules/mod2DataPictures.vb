﻿Imports stdole
Module mod2DataPictures
    Public Function FindDatabasePictureID(ByVal mLoc As Long, ByVal mType As Long, Optional ByVal mRef As String = "", Optional ByVal mIdx As Long = 1, Optional ByRef ImgCnt As Long = 0) As Long
        '::::FindDatabasePictureID
        ':::SUMMARY
        ': Used to find PictureID from Database.
        ':::DESCRIPTION
        ': This function is used to find PictureID from Database through sql statement.
        ':::PARAMETERS
        ':::RETURN
        Dim RS As ADODB.Recordset, S As String, I As Long
        S = ""
        S = S & "SELECT PictureID FROM Pictures WHERE 1=1 "
        S = S & "AND PictureType=" & mType & " "
        If mRef <> "" Then S = S & "AND PictureRef='" & ProtectSQL(mRef) & "' "
        S = S & "ORDER BY PictureID ASC"

        RS = GetRecordsetBySQL(S, , GetDatabaseAtLocation(mLoc))
        ImgCnt = RS.RecordCount
        If RS.RecordCount > 0 Then
            For I = 1 To mIdx
                If Not RS.EOF Then
                    FindDatabasePictureID = RS("PictureID").Value
                    RS.MoveNext()
                End If
            Next
        End If
        RS = Nothing
    End Function

    Public Function GetDatabasePictureToTempFile(ByVal PicID As Long) As String
        '::::GetDatabasePictureToTempFile
        ':::SUMMARY
        ': Gets Pictures from Database to temp file.
        ':::DESCRIPTION
        ': This function is used to get Pictures from Database based on PicID.
        ':::PARAMETERS
        ': PicID - Indicates the ID of picture.
        ':::RETURN
        ': String - Returns the result as a string.


        'Dim D As String, P As StdPicture
        Dim D As String, P As Image

        D = TempFile(, "CB_img", ".bmp")
        P = GetDatabasePicture(PicID)
        'SavePicture(P, D)---> Replacement for SavePicture is below line.
        P.Save(D, Imaging.ImageFormat.Bmp)
        GetDatabasePictureToTempFile = D
    End Function

    'Public Function GetDatabasePicture(ByVal PicID As Long) As StdPicture
    Public Function GetDatabasePicture(ByVal PicID As Long) As Image
        '::::GetDatabasePicture
        ':::SUMMARY
        ': Gets Pictures from Database.
        ':::DESCRIPTION
        ': This function is used to get Pictures from picture table through Sql Statement.
        ':::PARAMETERS
        ': - PicID - Indicates the picture ID.
        ':::RETURN
        ': StdPicture - Returns the result as a StdPicture object.
        With MainMenu
            ' BFH20070521 - Don't know any other way to get a picture out of a database than data control...
            '.datPicture.DatabaseName = GetDatabaseAtLocation()
            .datPicture.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetDatabaseAtLocation()
            '.datPicture.RecordSource = "SELECT Picture FROM Pictures WHERE PictureID=" & PicID
            .datPicture.RecordSource = "SELECT Picture FROM Pictures WHERE PictureID=" & PicID
            '.datPicture.Refresh
            .datPicture.Refresh()
            'If .datPicture.Recordset.RecordCount <> 0 Then
            If .datPicture.Recordset.RecordCount <> 0 Then
                'GetDatabasePicture = .imgPicture.Picture
                GetDatabasePicture = .imgPicture.Image
            End If
            '.datPicture.DataBase.Close
            .datPicture.ConnectionString = ""
        End With
    End Function

End Module
