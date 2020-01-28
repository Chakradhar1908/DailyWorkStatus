Module modAWS
    Dim mVersionString As String
    Dim LastCmd As String
    Private AmazonAutoBackup As Boolean
    Private Const CONFIG_KEY_AWS_LBD As String = "AWS_LastBackupDate"
    Private Const CONFIG_KEY_AWS_LBD_LC As String = "AWS_LastBackupDate_LastChecked"

    Public Function AWS_AutoBackup() As Boolean
        If Not IsServer() Then Exit Function
        If Not AWS_Check_Install() Then Exit Function
        If Not AWS_Check_Login() Then Exit Function

        If Val(ReadStoreSetting(0, IniSections_StoreSettings.iniSection_Amazon, "NoAutoBackup")) <> 0 Then Exit Function

        AWS_AutoBackup = True
    End Function

    Public Function AWS_Check_Login(Optional ByRef ErrMsg As String = "") As Boolean
        ErrMsg = ""
        If StoreSettings.AmazonKeyID = "" Then ErrMsg = "Key ID Not Set in Store Setup." : Exit Function
        If StoreSettings.AmazonSecretKey = "" Then ErrMsg = "Secret Key Not Set in Store Setup." : Exit Function
        '  If StoreSettings.AmazonUserName = "" Then ErrMsg = "UserName Not Set in Store Setup.": Exit Function
        '  If StoreSettings.AmazonCustomerBucket = "" Then ErrMsg = "Customer Bucket Not Set in Store Setup.": Exit Function
        '  AWS_Command "aws s3api list-objects --bucket " & StoreSettings.AmazonCustomerBucket, ErrMsg
        '  If ErrMsg <> "" Then
        '    AWS_Check_Login = False
        '  Else
        AWS_Check_Login = True
        '  End If
    End Function

    Public Function AWS_Check_Install(Optional ByRef VersionString As String = "") As Boolean
        Dim X As String

        X = "Amazon\AWSCLI\aws.exe"
        If AWS = "aws" Then
            '  If Not FileExists(LocalProgramFilesFolder() & X) And Not FileExists(LocalProgramFilesFolder(True) & X) Then
            ' we will not call the AWS program unless we can find it.
            ' This prevents uninstalled customers from seeing extraneous error messages.
            Exit Function
        End If

        Dim ErrStr As String
        If mVersionString = "" Then
            VersionString = ""
            VersionString = AWS_Command(" --version", ErrStr)
            mVersionString = ErrStr
        Else
            ErrStr = mVersionString
        End If

        If Left(ErrStr, 7) = "aws-cli" Then
            AWS_Check_Install = True
            VersionString = ErrStr
        End If
    End Function

    Private Function AWS_Command(ByVal vCmd As String, Optional ByRef ErrStr As String = "", Optional ByRef P As Object = Nothing) As String
        Dim T As String, tErr As Boolean
        ErrStr = ""
        LastCmd = vCmd

        '  If IsCorvinsETown Then
        '    If Not FolderExists(LocalDesktopFolder & ".aws") Then
        '      AWS_Configure_ViaCommand
        '    End If
        '  End If

        AWS_Command = RunCmdToOutputWithArgs(AWS, vCmd, ErrStr)
        If ErrStr <> "" Then
            If InStr(ErrStr, "Python") = 0 Then
                AWS_Log("AWS COMMAND ERROR: " & ErrStr)
            End If
        Else
            On Error Resume Next
            P = JSON.Parse(AWS_Command)
        End If
    End Function

    Public Function AWS_Log(ByVal vMsg As String) As Boolean
        ' simply keep a running log of events.
        LogFile("AWSLog.txt", vMsg, False)
    End Function

    Private ReadOnly Property AWS() As String
        Get
            '  AWS = "aws"
            '  AWS = "awst" ' can be used to test the install..  by supplying a false EXE name, the program thinks the software is not installed.
            AWS = AWS_CommandPath
        End Get
    End Property

    Public ReadOnly Property AWS_CommandPath() As String
        Get
            Dim X As String
            X = "Amazon\AWSCLI\aws.exe"
            If FileExists(LocalProgramFilesFolder(True) & X) Then AWS_CommandPath = LocalProgramFilesFolder(True) & X : Exit Property
            If FileExists(LocalProgramFilesFolder() & X) Then AWS_CommandPath = LocalProgramFilesFolder() & X : Exit Property
            AWS_CommandPath = "aws"
        End Get
    End Property

    Public Sub DoAmazonAutoBackup()
        Dim X As String, C As clsHashTable, T As BackupType
        Dim B As String
        Dim P As Object, Q As Object
        Dim I As Integer, L As String, M As String, N As String

        ' BFH20150605
        ' If we allow this to run on a CDS computer, and it is configured as a client, it could replace their backup
        If IsCDSComputer() Then Exit Sub
        If IsDemo() Then Exit Sub

        'BFH20151130 - We ran into this issue this year...  A store opened while backup was still going on
        ' We disable auto-backup on Black Friday (day after Thanksgiving) to prevent this collision.
        If DateEqual(Today, BlackFridayDate) Then Exit Sub

        AmazonAutoBackup = True

        C = New clsHashTable
        X = ReadStoreSetting(0, IniSections_StoreSettings.iniSection_Amazon, "AWS Panel Config")
        C.LoadQueryString(X)

        T = BackupType.bkNone
        If Val(C.Item("chkPS")) <> 0 Then T = T + BackupType.bkPS
        If Val(C.Item("chkSS")) <> 0 Then T = T + BackupType.bkSS
        If Val(C.Item("chkPX")) <> 0 Then T = T + BackupType.bkpx
        If Val(C.Item("chkGL")) <> 0 Then T = T + BackupType.bkGL
        If Val(C.Item("chkBK")) <> 0 Then T = T + BackupType.bkBK
        If Val(C.Item("chkAP")) <> 0 Then T = T + BackupType.bkAP
        If Val(C.Item("chkPR")) <> 0 Then T = T + BackupType.bkPR
        If Val(C.Item("chkQB")) <> 0 Then T = T + BackupType.bkQB
        If T = BackupType.bkNone Then T = BackupType.bkAll Else T = T + BackupType.bkLO

        AWS_CheckCredentialFile() ' in case the server runs on a new user account.

        DoAmazonBackup(True, T)

        B = StoreSettings.AmazonCustomerBucket
ResetList:
        P = AWS_ListObjects(B)
        If P Is Nothing Then Exit Sub
        Q = P.Item("Contents")
        For I = 1 To Q.Count
            L = Q(I).Item("Key")
            M = AmazonAutoBackupExpired(L)
            If M <> "" Then
                If N = M Then GoTo FindNextFolder             ' prevent loops
                Debug.Print("Amazon Delete Object: " & M)
                AWS_DeleteObjects(B, M)
                N = M
            End If
FindNextFolder:
        Next

        LastAmazonBackupDate() ' update the value of how recent the backup is
        AmazonAutoBackup = False
    End Sub

    Public Function LastAmazonBackupDate() As Date
        Dim I As Integer
        Dim P As Object, Q As Object
        Dim X As New clsHashTable, N As String, M As String
        Dim V() As Object, L As Object
        Dim Mx As Date

        '  ProgressForm 0, 1, "Loading Restore Points..."
        If StoreSettings.AmazonCustomerBucket = "" Then LastAmazonBackupDate = NullDate : Exit Function

        SuppressMessages(15)
        P = AWS_ListObjects(StoreSettings.AmazonCustomerBucket)
        SuppressMessages()
        If P Is Nothing Then Exit Function
        On Error GoTo NoBuckets
        Q = P.Item("Contents")
        For I = 1 To Q.Count
            N = Q(I).Item("Key")
            M = SplitWord(N, 1, "/")
            If Not X.Exists(M) Then X.Add(M, M)
        Next
NoBuckets:
        Err.Clear()

        Mx = NullDate

        If X.Count > 0 Then
            V = X.Keys(vbFalse)
            On Error GoTo NoItems
            For Each L In V
                If DateAfter(DateStampValue(L), Mx) Then
                    Mx = DateStampValue(L)
                End If
            Next
NoItems:
            Err.Clear()
        End If

        SetConfigTableValue(CONFIG_KEY_AWS_LBD, DateStamp(Mx))
        SetConfigTableValue(CONFIG_KEY_AWS_LBD_LC, DateStamp)

        LastAmazonBackupDate = Mx
    End Function

    Public Function AWS_DeleteObjects(ByVal Bucket As String, Optional ByVal Prefix As String = "") As Boolean
        Dim P As Object, I As Integer, K As String
        P = AWS_ListObjects(Bucket, Prefix)
        If P Is Nothing Then Exit Function
        P = P.Item("Contents")
        If P Is Nothing Then Exit Function
        For I = 1 To P.Count
            K = P(I).Item("Key")
            AWS_DeleteObject(Bucket, K)
        Next

        AWS_DeleteObjects = True
    End Function

    Public Function AWS_DeleteObject(ByVal Bucket As String, ByVal ObjectID As String) As Boolean
        Dim Res As String, ErrMsg As String, P As Object
        Res = AWS_Command(" s3api delete-object --bucket=" & Bucket & " --key=""" & ObjectID & """", ErrMsg, P)
        If ErrMsg <> "" Then AWSErr(ErrMsg) : Exit Function

        AWS_DeleteObject = True
    End Function

    Private Sub AWSErr(ByVal Msg As String, Optional ByVal MsgTitle As String = "AWS Error")
        AWS_Log("**************************************************")
        AWS_Log("*** AWS ERROR - " & Msg)
        AWS_Log("*** AWS LAST CMD - " & LastCmd)
        If Not ACLHasFullAccess(UpdateFolder) Then AWS_Log("*** Update Folder Not Set for Full Access: " & UpdateFolder() & " " & ACL_FA(UpdateFolder))
        If IsInStr(Msg, "credentials") Then AWS_Log("*** CREDENTIALS ERROR, Current User Is [" & GetSystemUserName & "," & GetDirectoryUserName & "]")
        AWS_Log("**************************************************")
        Debug.Print("AWS Error: " & Msg)
        Debug.Print("AWS LastCmd: " & LastCmd)
        MessageBox.Show(Msg, MsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Public Function AmazonAutoBackupExpired(ByVal vKEY As String) As String
        ' Empty String is NOT expired.
        ' Return Datestamp of Key is EXPIRED.
        Dim T As String, D As Date, Ck As Date, isQ As Boolean
        Dim F As String, ckF As Date, isF As Boolean
        Dim ckY As Date, isY As Boolean

        Const PXfile As String = "bupx.zip"
        Const QBfile As String = "buqb.zip"

        T = Left(vKEY, 8)

        If Not (T Like "########") Then Exit Function

        D = DateStampValue(T)
        Ck = DateAdd("m", -6, Today)
        ckY = DateAdd("yyyy", -3, Today)
        isY = DateBefore(D, ckY)

        F = LCase(Mid(vKEY, 10))

        If Month(D) = 12 And IsIn(DateAndTime.Day(D), 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31) Then isQ = True
        If Month(D) = 9 And IsIn(DateAndTime.Day(D), 29, 30) Then isQ = True
        If Month(D) = 6 And IsIn(DateAndTime.Day(D), 29, 30) Then isQ = True
        If Month(D) = 3 And IsIn(DateAndTime.Day(D), 30, 31) Then isQ = True
        If Month(D) = 1 And DateAndTime.Day(D) >= 1 And DateAndTime.Day(D) <= 15 Then isQ = True

        If IsIn(F, PXfile, QBfile) Then
            ckF = DateAdd("d", -7, Today)
            isF = DayOfWeek(D) = "Friday"
            '
            AmazonAutoBackupExpired = IIf((DateBefore(D, Ck) And Not isQ) Or (DateBefore(D, ckF) And Not isF And Not isQ), vKEY, "")
        Else
            AmazonAutoBackupExpired = IIf(DateBefore(D, Ck) And Not isQ Or isY, T, "")
        End If
    End Function

    Public Function AWS_ListObjects(ByVal Bucket As String, Optional ByVal Prefix As String = "") As Object
        Dim Res As String, ErrMsg As String, P As Object, Pfx As String
        If Prefix <> "" Then Pfx = " --prefix """ & Prefix & """"
        Res = AWS_Command(" s3api list-objects --bucket " & Bucket & Pfx, ErrMsg, P)
        'If True Then
        'Debug.Print Res
        'Debug.Print LastCmd
        'End If

        If ErrMsg <> "" Then AWSErr(ErrMsg) : Exit Function
        AWS_ListObjects = P
    End Function

    Public Function AWS_CheckCredentialFile() As Boolean
        If Not FileExists(AWS_CredentialFile) Then AWS_Configure
        AWS_CheckCredentialFile = True
    End Function

    Public Sub AWS_Configure(Optional ByVal Profile As String = "default", Optional ByVal Force As Boolean = False)
        Dim T As String, X As String
        If Not FolderExists(ParentDirectory(GetFilePath(AWS_ConfigFile()))) Then Exit Sub
        If Not FolderExists(GetFilePath(AWS_ConfigFile())) Then MkDir(GetFilePath(AWS_ConfigFile()))


        T = UserFolder() & ".aws\config"
        If Not FileExists(T) Or Force Then
            WriteFile(T, "[" & Profile & "]" & vbCrLf & "aws_access_key_id=" & vbCrLf & "aws_secret_access_key=" & vbCrLf2, True)
            Exit Sub
        End If

        If Profile <> "default" Then Profile = "profile " & Profile

        WriteIniValue(T, Profile, "aws_access_key_id", StoreSettings.AmazonKeyID)
        WriteIniValue(T, Profile, "aws_secret_access_key", StoreSettings.AmazonSecretKey)

        '  WriteIniValue T, Profile, "region", StoreSettings.AmazonRegionName
        '  WriteIniValue T, Profile, "output", StoreSettings.AmazonCustomerBucket
        '  WriteIniValue T, Profile, "profile", StoreSettings.AmazonKeyID
    End Sub


    Public Function AWS_ConfigFile() As String
        Const AWSCONFIG As String = ".aws\config"
        AWS_ConfigFile = UserFolder() & AWSCONFIG
        If Not FileExists(AWS_ConfigFile) And FileExists(LocalDesktopFolder() & AWSCONFIG) Then AWS_ConfigFile = LocalDesktopFolder() & AWSCONFIG
    End Function

    Public ReadOnly Property AWS_CredentialFile() As String
        Get
            AWS_CredentialFile = UserFolder() & ".aws\config"
        End Get
    End Property

    Public Sub DoAmazonBackup(Optional ByVal Suppress As Boolean = False, Optional ByVal Files As BackupType = BackupType.bkAll, Optional ByVal AlternateBucket As String = "", Optional ByVal AlternateDate As String = "")
        Dim X As String, S() As String, L As Object
        Dim F As String, Ex As Date
        Dim Bucket As String
        Dim Folder As String

        If Files = BackupType.bkNone Then Files = BackupType.bkAll

        Bucket = IIf(AlternateBucket <> "", AlternateBucket, StoreSettings.AmazonCustomerBucket)
        Folder = IIf(AlternateDate <> "", AlternateDate, DateStamp)

        If Bucket = "" Then Exit Sub
        If StoreSettings.AmazonKeyID = "" Then Exit Sub
        If StoreSettings.AmazonSecretKey = "" Then Exit Sub
        If StoreSettings.AmazonUserName = "" Then Exit Sub

        AWS_Log("+++++++++++++++++++++++++++++++++++++++++++++++++++")
        AWS_Log("+++++++++  ---  BEGIN AMAZON BACKUP  ---  +++++++++")
        '  AWS_Log "+++++++++++++++++++++++++++++++++++++++++++++++++++"

        If Suppress Then SuppressMessages(60) ' silence all msg boxes for 1 hour

        Select Case DateAndTime.Day(Today)
'    Case 1: Ex = CDate(0)
            Case 2 To 9, 11 To 19, 21 To 29, 31 : Ex = DateAdd("d", 30, Today) ' 1 month
            Case 10, 20, 30 : Ex = DateAdd("d", 120, Today) ' 4 months
        End Select

        X = TempFolder()
        AWS_Log("+++ Program: " & SoftwareVersionForLog())
        AWS_Log("+++ Folder: " & X & " " & FolderExists_E(X) & " " & ACL_FA(X))
        If UseAWSProgressForm Then ProgressForm(0, 1, "Backing Up Databases...")
        frmBackUpGeneric.BackupTo(X, Files)
        If UseAWSProgressForm Then ProgressForm(0, 1, "Uploading to AWS...")

        S = AllFiles(X)
        AWS_Log("+++ Amazon Backup Path File Count: " & UBound(S))
        F = DateStamp()
        AWS_Sync_Folder(StoreSettings.AmazonCustomerBucket, X, F)
        '  For Each l In S
        '    AWS_Log "+++ amazon Backup, Uploading file: " & l
        '    AWS_PutObject StoreSettings.AmazonCustomerBucket, X & l, F & GetFileName(l), Ex
        '  Next

        If UseAWSProgressForm Then ProgressForm(0, 1, "Cleaning Up...")
        AWS_Log("+++ Amazon Backup Cleaning Up")
        ClearFolder(X)
        RmDir(X)
        If UseAWSProgressForm Then ProgressForm()

        LastAmazonBackupDate()

        'AWS_Log "+++++++++++++++++++++++++++++++++++++++++++++++++++"
        AWS_Log("+++++++++   ===  END AMAZON BACKUP  ===   +++++++++")
        AWS_Log("+++++++++++++++++++++++++++++++++++++++++++++++++++")

        SuppressMessages()
    End Sub

    Private ReadOnly Property UseAWSProgressForm() As Boolean
        Get
            UseAWSProgressForm = Not AmazonAutoBackup
        End Get
    End Property

    Public Function AWS_Sync_Folder(ByVal Bucket As String, ByVal Src As String, ByVal Dst As String) As Boolean
        Dim C As String
        Dim Res As String, ErrMsg As String, P As Object

        'aws s3 sync . s3://com.WinCDS.YourFurnitureStore/20141127
        C = " s3 sync " & Src & " s3://" & Bucket & "/" & Dst
        Res = AWS_Command(C, ErrMsg, P)
        If ErrMsg <> "" Then AWSErr(ErrMsg) : Exit Function

        AWS_Sync_Folder = True
    End Function

End Module
