Module modZipUnzip
    Public Structure DCLIST
        Dim ExtractOnlyNewer As Integer      ' 1 to extract only newer
        Dim SpaceToUnderScore As Integer     ' 1 to convert spaces to underscore
        Dim PromptToOverwrite As Integer     ' 1 if overwriting prompts required
        Dim fQuiet As Integer                ' 0 = all messages, 1 = few messages, 2 = no Messages
        Dim ncflag As Integer                ' write to stdout if 1
        Dim ntflag As Integer                ' test zip file
        Dim nvflag As Integer                ' verbose listing
        Dim nUflag As Integer                ' "update" (extract only newer/new files)
        Dim nzflag As Integer                ' display zip file comment
        Dim ndflag As Integer                ' all args are files/dir to be extracted
        Dim noflag As Integer                ' 1 if always overwrite files
        Dim naflag As Integer                ' 1 to do end-of-line translation
        Dim nZIflag As Integer               ' 1 to get zip info
        Dim C_flag As Integer                ' 1 to be case insensitive
        Dim fPrivilege As Integer            ' zip file name
        Dim lpszZipFN As String           ' directory to extract to.
        Dim lpszExtractDir As String
    End Structure

    Private Structure USERFUNCTION
        ' Callbacks:
        Dim lptrPrnt As Object            ' Pointer to application's print routine
        Dim lptrSound As Object           ' Pointer to application's sound routine.  NULL if app doesn't use sound
        Dim lptrReplace As Object         ' Pointer to application's replace routine.
        Dim lptrPassword As Object        ' Pointer to application's password routine.
        Dim lptrMessage As Object         ' Pointer to application's routine for
        ' displaying information about specific files in the Archive
        ' used for listing the contents of the archive.
        Dim lptrService As Object         ' callback function designed to be used for allowing the
        ' app to process Windows messages, or cancelling the Operation
        ' as well as giving option of progress.  If this Function returns
        ' non-zero, it will terminate what it is doing.  It provides the app
        ' with the name of the archive member it has just processed, as well
        ' as the original size.

        ' Values filled in after processing:
        Dim lTotalSizeComp As Integer     ' Value to be filled in for the compressed total Size , excluding
        ' the archive header and central directory list.
        Dim lTotalSize As Integer         ' Total size of all files in the archive
        Dim lCompFactor As Integer        ' Overall archive compression factor
        Dim lNumMembers As Integer        ' Total number of files in the archive
        Dim cchComment As Integer      ' Flag indicating whether comment in archive.
    End Structure

    Private Structure UNZIPnames
        'Dim S(0 To 1023) As String
        Dim S() As String
    End Structure

    Private Structure CBChar
        Dim Ch() As Byte
    End Structure

    Private Structure CBCh
        Dim Ch() As Byte
    End Structure

    Private m_bCancel As Boolean
    Private m_cUnzip As cZipUnzip
    Private Delegate Function DelUnzipPrintCallback(ByRef fName As CBChar, ByVal X As Integer) As Integer
    Private Delegate Function DelUnzipReplaceCallback(ByRef fName As CBChar) As Integer
    Private Delegate Function DelUnzipPasswordCallBack(ByRef Pwd As CBCh, ByVal X As Integer, ByRef S2 As CBCh, ByRef Name As CBCh) As Integer
    Private Delegate Sub DelUnzipMessageCallBack(
      ByVal ucsize As Integer,
      ByVal csiz As Integer,
      ByVal cfactor As Integer,
      ByVal Mo As Integer,
      ByVal dY As Integer,
      ByVal Yr As Integer,
      ByVal HH As Integer,
      ByVal MM As Integer,
      ByVal C As Byte,
      ByRef fName As CBCh,
      ByRef meth As CBCh,
      ByVal crc As Integer,
      ByVal fCrypt As Byte
   )
    Private Delegate Function DelUnZipServiceCallback(ByRef mName As CBChar, ByVal X As Integer) As Integer

    Dim DelUnzip As DelUnzipPrintCallback
    Dim DelUnzipReplace As DelUnzipReplaceCallback
    Dim DelUnzipPassword As DelUnzipPasswordCallBack
    Dim DelUnzipMessage As DelUnzipMessageCallBack
    Dim DelUnzipService As DelUnZipServiceCallback
    Private Declare Function Wiz_SingleEntryUnzip Lib "vbuzip10.dll" (ByVal ifnc As Integer, ByRef ifnv As UNZIPnames, ByVal xfnc As Integer, ByRef xfnv As UNZIPnames, dcll As DCLIST, Userf As USERFUNCTION) As Integer
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (
    lpvDest As Object, lpvSource As Object, ByVal cbCopy As Integer)

    Public Function VBUnzip(ByRef cUnzipObject As cZipUnzip, ByRef tDCL As DCLIST, ByRef iIncCount As Integer, ByRef sInc() As String, ByRef iExCount As Integer, ByRef sExc() As String) As Integer
        Dim tUser As USERFUNCTION
        Dim lR As Integer
        Dim tInc As UNZIPnames
        Dim tExc As UNZIPnames
        Dim I As Integer

        On Error GoTo ErrorHandler

        m_cUnzip = cUnzipObject
        ' Set Callback addresses

        DelUnzip = AddressOf UnzipPrintCallback
        'tUser.lptrPrnt = plAddressOf(AddressOf UnzipPrintCallback)
        tUser.lptrPrnt = plAddressOf1(DelUnzip)

        tUser.lptrSound = 0& ' not supported
        'tUser.lptrReplace = plAddressOf(AddressOf UnzipReplaceCallback)
        DelUnzipReplace = AddressOf UnzipReplaceCallback
        tUser.lptrReplace = plAddressOf2(DelUnzipReplace)

        'tUser.lptrPassword = plAddressOf(AddressOf UnzipPasswordCallBack)
        DelUnzipPassword = AddressOf UnzipPasswordCallBack
        tUser.lptrPassword = plAddressOf3(DelUnzipPassword)

        'tUser.lptrMessage = plAddressOf(AddressOf UnzipMessageCallBack)
        DelUnzipMessage = AddressOf UnzipMessageCallBack
        tUser.lptrMessage = plAddressOf4(DelUnzipMessage)

        'tUser.lptrService = plAddressOf(AddressOf UnZipServiceCallback)
        DelUnzipService = AddressOf UnZipServiceCallback
        tUser.lptrService = plAddressOf5(DelUnzipService)

        ' Set files to include/exclude:
        If (iIncCount > 0) Then
            For I = 1 To iIncCount
                tInc.S(I - 1) = sInc(I)
            Next
            tInc.S(iIncCount) = vbNullChar
        Else
            tInc.S(0) = vbNullChar
        End If
        If (iExCount > 0) Then
            For I = 1 To iExCount
                tExc.S(I - 1) = sExc(I)
            Next
            tExc.S(iExCount) = vbNullChar
        Else
            tExc.S(0) = vbNullChar
        End If
        m_bCancel = False
        VBUnzip = Wiz_SingleEntryUnzip(iIncCount, tInc, iExCount, tExc, tDCL, tUser)

        'Debug.Print "--------------"
        'Debug.Print MYUSER.cchComment
        'Debug.Print MYUSER.TotalSizeComp
        'Debug.Print MYUSER.TotalSize
        'Debug.Print MYUSER.CompFactor
        'Debug.Print MYUSER.NumMembers
        'Debug.Print "--------------"

        Exit Function

ErrorHandler:
        Dim lErr As Integer, sErr As Integer
        lErr = Err.Number : sErr = Err.Description
        VBUnzip = -1
        m_cUnzip = Nothing
        Err.Raise(lErr, WinCDSEXEName() & ".VBUnzip", sErr)
        Exit Function

    End Function

    Private Function plAddressOf1(ByVal lPtr As DelUnzipPrintCallback) As Object
        ' VB Bug workaround fn
        plAddressOf1 = lPtr
    End Function

    Private Function plAddressOf2(ByVal lPtr As DelUnzipReplaceCallback) As Object
        ' VB Bug workaround fn
        plAddressOf2 = lPtr
    End Function

    Private Function plAddressOf3(ByVal lPtr As DelUnzipPasswordCallBack) As Object
        ' VB Bug workaround fn
        plAddressOf3 = lPtr
    End Function

    Private Function plAddressOf4(ByVal lPtr As DelUnzipMessageCallBack) As Object
        ' VB Bug workaround fn
        plAddressOf4 = lPtr
    End Function

    Private Function plAddressOf5(ByVal lPtr As DelUnZipServiceCallback) As Object
        ' VB Bug workaround fn
        plAddressOf5 = lPtr
    End Function

    Private Function UnzipReplaceCallback(ByRef fName As CBChar) As Integer
        Dim eResponse As cZipUnzip.EUZOverWriteResponse
        Dim iPos As Integer
        Dim sFIle As String

        On Error Resume Next
        eResponse = cZipUnzip.EUZOverWriteResponse.euzDoNotOverwrite

        ' Extract the filename:
        sFIle = StrConv(fName.Ch.ToString, VBA.VbStrConv.vbUnicode)
        iPos = InStr(sFIle, vbNullChar)
        If (iPos > 1) Then
            sFIle = Left(sFIle, iPos - 1)
        End If

        ' No backslashes:
        ReplaceSection(sFIle, "/", "\")

        ' Request the overwrite request:
        m_cUnzip.OverwriteRequest(sFIle, eResponse)

        ' Return it to the zipping lib
        UnzipReplaceCallback = eResponse
    End Function

    Private Function ReplaceSection(ByRef sString As String, ByVal sToReplace As String, ByVal sReplaceWith As String) As Integer
        Dim iPos As Integer
        Dim iLastPos As Integer
        iLastPos = 1
        Do
            iPos = InStr(iLastPos, sString, "/")
            If (iPos > 1) Then
                Mid(sString, iPos, 1) = "\"
                iLastPos = iPos + 1
            End If
        Loop While Not (iPos = 0)
        ReplaceSection = iLastPos
    End Function

    Private Function UnzipPrintCallback(ByRef fName As CBChar, ByVal X As Integer) As Integer
        Dim iPos As Integer
        Dim sFIle As String
        Dim B() As Byte

        On Error Resume Next
        ' Check we've got a message:
        If X > 1 And X < 1024 Then
            ' If so, then get the readable portion of it:
            ReDim B(0 To X)
            CopyMemory(B(0), fName, X)
            ' Convert to VB string:
            sFIle = StrConv(B.ToString, VBA.VbStrConv.vbUnicode)

            ' Fix up backslashes:
            ReplaceSection(sFIle, "/", "\")

            ' Tell the caller about it
            m_cUnzip.ProgressReport(sFIle)
        End If
        UnzipPrintCallback = 0
    End Function

    Private Function UnzipPasswordCallBack(ByRef Pwd As CBCh, ByVal X As Integer, ByRef S2 As CBCh, ByRef Name As CBCh) As Integer
        Dim bCancel As Boolean
        Dim sPassword As String
        Dim B() As Byte
        Dim lSize As Integer

        On Error Resume Next

        ' The default:
        UnzipPasswordCallBack = 1

        If m_bCancel Then Exit Function

        ' Ask for password:
        m_cUnzip.PasswordRequest(sPassword, bCancel)

        sPassword = Trim(sPassword)

        ' Cancel out if no useful password:
        If bCancel Or Len(sPassword) = 0 Then
            m_bCancel = True
            Exit Function
        End If

        ' Put password into return parameter:
        lSize = Len(sPassword)
        If lSize > 254 Then
            lSize = 254
        End If
        B(0) = StrConv(sPassword, VBA.VbStrConv.vbFromUnicode)
        CopyMemory(Pwd.Ch(0), B(0), lSize)

        ' Ask UnZip to process it:
        UnzipPasswordCallBack = 0
    End Function

    Private Sub UnzipMessageCallBack(
      ByVal ucsize As Integer,
      ByVal csiz As Integer,
      ByVal cfactor As Integer,
      ByVal Mo As Integer,
      ByVal dY As Integer,
      ByVal Yr As Integer,
      ByVal HH As Integer,
      ByVal MM As Integer,
      ByVal C As Byte,
      ByRef fName As CBCh,
      ByRef meth As CBCh,
      ByVal crc As Integer,
      ByVal fCrypt As Byte
   )
        Dim sFileName As String
        Dim sFolder As String
        Dim DDate As Date
        Dim sMethod As String
        Dim iPos As Integer

        On Error Resume Next

        ' Add to unzip class:
        ' Parse:
        sFileName = StrConv(fName.Ch.ToString, VBA.VbStrConv.vbUnicode)
        ParseFileFolder(sFileName, sFolder)
        DDate = DateSerial(Yr, Mo, HH)
        DDate = DDate + TimeSerial(HH, MM, 0)
        sMethod = StrConv(meth.Ch.ToString, VBA.VbStrConv.vbUnicode)
        iPos = InStr(sMethod, vbNullChar)
        If (iPos > 1) Then
            sMethod = Left(sMethod, iPos - 1)
        End If

        'Debug.Print fCrypt
        m_cUnzip.DirectoryListAddFile(sFileName, sFolder, DDate, csiz, crc, ((fCrypt And 64) = 64), cfactor, sMethod)
    End Sub

    Private Sub ParseFileFolder(ByRef sFileName As String, ByRef sFolder As String)
        Dim iPos As Integer
        Dim iLastPos As Integer

        iPos = InStr(sFileName, vbNullChar)
        If (iPos <> 0) Then
            sFileName = Left(sFileName, iPos - 1)
        End If

        iLastPos = ReplaceSection(sFileName, "/", "\")

        If (iLastPos > 1) Then
            sFolder = Left(sFileName, iLastPos - 2)
            sFileName = Mid(sFileName, iLastPos)
        End If
    End Sub

    Private Function UnZipServiceCallback(ByRef mName As CBChar, ByVal X As Integer) As Integer
        Dim iPos As Integer
        Dim sInfo As String
        Dim bCancel As Boolean
        Dim B() As Byte

        '-- Always Put This In Callback Routines!
        On Error Resume Next

        ' Check we've got a message:
        If X > 1 And X < 1024 Then
            ' If so, then get the readable portion of it:
            ReDim B(0 To X)
            CopyMemory(B(0), mName, X)
            ' Convert to VB string:
            sInfo = StrConv(B.ToString, VBA.VbStrConv.vbUnicode)
            iPos = InStr(sInfo, vbNullChar)
            If iPos > 0 Then
                sInfo = Left(sInfo, iPos - 1)
            End If
            ReplaceSection(sInfo, "\", "/")
            m_cUnzip.Service(sInfo, bCancel)
            If bCancel Then
                UnZipServiceCallback = 1
            Else
                UnZipServiceCallback = 0
            End If
        End If
    End Function
End Module
