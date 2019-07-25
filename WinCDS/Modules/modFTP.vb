Module modFTP
    Private Const FTP_Port as integer = 21
    Public Function FTP_Put(ByVal vHost As String, ByVal vUser As String, ByVal vPass As String, ByVal vRemoteDir As String, ByVal vLocalFile As String, Optional ByVal CreateFolder As Boolean = False) As Boolean
        '::::FTP_Put
        ':::SUMMARY
        ': Upload a single file
        ':::DESCRIPTION
        ':  Connects via FTP to the Host/Port with user/pass and uploads vLocalFile to the folder vRemoteDir
        ':::PARAMETERS
        ': - vHost - The remote IP address
        ': - vUser - Username to connect with
        ': - vPass - Password for authorization
        ': - vRemoteDir - Remote path.  Can be absolute or relative.
        ': - vLocalFile - The full path to the local file to upload.
        ': - CreateFolder - Optional.  If True, will walk the path and attempt to create each the full directory path before uploading.
        ':::RETURN
        ':  Boolean - Returns True on success.
        ':::SEE ALSO
        ': - FTP_Get, FTP_PutDir
        ': - cFTP
        Dim F As cFTP
        F = New cFTP

        F.SetTransferBinary()
        F.SetModePassive()

        If Not F.OpenConnection(vHost, vUser, vPass, FTP_Port) Then Exit Function
        FTP_NavigateToFolder(F, vRemoteDir, CreateFolder)

        FTP_Put = F.UploadFile(vLocalFile, GetFileName(vLocalFile))
        '  FTP_Put = F.SimplePutFile(vLocalFile, GetFileName(vLocalFile))

        F.CloseConnection()
        DisposeDA(F)
        FTP_Put = True
    End Function

    Private Sub FTP_NavigateToFolder(ByRef F As cFTP, ByVal vDir As String, Optional ByVal CreateFolder As Boolean = False)
        Dim L, P As String, T As String

        On Error Resume Next
        If CreateFolder Then
            P = vDir

            If Left(P, 1) = "/" Then
                F.SetDirectory("/")
                P = Mid(P, 2)
                T = "/"
            End If
            If Right(P, 1) = "/" Then P = Left(P, Len(P) - 1)

            For Each L In Split(P, "/")
                F.CreateDirectory(L)
                F.SetDirectory(L)
            Next
        Else
            F.SetDirectory(vDir)
        End If
    End Sub

End Module