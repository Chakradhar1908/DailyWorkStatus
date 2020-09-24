Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class frmBackUpGeneric
    Dim intLastDrive As Integer
    Private mBackupMode As Integer, mBackupFiles As BackupType
    'Private WithEvents cZ As cZip
    'Private WithEvents cU As cZipUnzip
    Private ZipFileCount As Integer
    Public Expedite As Boolean
    Public IgnoreMissingRestores As Boolean
    Private P As frmProgress
    Private WithEvents cU As cZipUnzip

    Public Function ZipFiles(ByVal CompressPath As String, ByVal ZipDir As String, ByVal ZipFile As String, Optional ByVal Special As Integer = 0, Optional ByVal FileMask As String = "") As Boolean
        'Status = "Zipping " & ZipFile & "..."   ERROR
        'Select Case modBackup.ZipType
        'Case wzt7ZIP : ZipFiles = SevenZipZipFiles(CompressPath, ZipDir, ZipFile, Special, FileMask)
        '    Case wztINFO : ZipFiles = InfoZipZipFiles(CompressPath, ZipDir, ZipFile, Special)
        '    Case wztVJCZ : ZipFiles = VJCZipFiles(CompressPath, ZipDir, ZipFile, Special)
        'Case wztNone : Err.Raise -1, , "No Valid Zip component."   ERROR
        'Case Else : DevErr "frmBackupGeneric.ZipFiles - Unknown Zip Component [" & modBackup.ZipType & "]"  ERROR
        'End Select
    End Function

    Public Property Mode() As BackupMode
        Get
            Mode = mBackupMode
        End Get
        Set(value As BackupMode)
            mBackupMode = value
        End Set
    End Property

    Public Property BackupFiles() As BackupType
        Get
            BackupFiles = mBackupFiles
        End Get
        Set(value As BackupType)
            mBackupFiles = value
        End Set
    End Property

    Public Sub Display(ByVal vMode As Integer, ByVal Files As BackupType, Optional ByRef ParentForm As Form = Nothing, Optional ByVal Modal As Integer = 0)
        ' Banking, GL, Payables, Payroll, POS, All
        Dim ModeName As String

        Mode = vMode
        BackupFiles = Files

        Select Case Mode
            Case BackupMode.bkRestore
                ModeName = "Restore"
                lblDriveSelect.Text = "Select drive to restore from:"
                chkNewFolder.Visible = False
                chkNewFolder.Checked = False
                txtNewFolder.Visible = True
                txtNewFolder.ReadOnly = True
                'Dir1.Visible = True
                lvwFiles.Visible = True
                Width = 8010
                'HelpContextID = 33000

            Case BackupMode.bkBackup
                ModeName = "Back Up"
                lblDriveSelect.Text = "Select drive to backup to:"
                chkNewFolder.Visible = True
                chkNewFolder.Checked = False
                txtNewFolder.ReadOnly = False
                'Dir1.Visible = False
                lvwFiles.Visible = False
                cmdStart.Top = txtNewFolder.Top + txtNewFolder.Height + 60 ' Dir1.Top
                cmdCancel.Top = cmdStart.Top
                Width = 3900
                'HelpContextID = 32000
            Case Else
                MessageBox.Show("Error: Invalid backup mode.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                'Unload Me
                Me.Close()
                Exit Sub
        End Select
        Height = cmdStart.Top + cmdStart.Height + 780

        Select Case Files
            Case BackupType.bkBK : Text = ModeName & " Banking"
            Case BackupType.bkGL : Text = ModeName & " General Ledger"
            Case BackupType.bkAP : Text = ModeName & " Payables"
            Case BackupType.bkPR : Text = ModeName & " Payroll"
            Case BackupType.bkPS : Text = ModeName & " POS"
            Case BackupType.bkpx : Text = ModeName & " PX Folder"
            Case BackupType.bkAll : Text = ModeName & " Everything"
            Case BackupType.bkSS : Text = ModeName & " Store Setup"
            Case Else
                MessageBox.Show("Error: Invalid backup file selection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                'Unload Me
                Me.Close()
                Exit Sub
        End Select

        'Show Modal, ParentForm
        ShowDialog(ParentForm)
    End Sub

    Public Function BackupTo(ByVal Folder As String, Optional ByVal Files As BackupType = BackupType.bkAll) As Boolean
        BackupLog("------------------------------------------------------")
        BackupLog(SoftwareVersionForLog)
        BackupLog("BackupTo(Folder:=" & Folder & ", Files:=" & DescribeBackupType(Files) & ")")
        'drvBackup.Drive = Left(Folder, 2)
        chkNewFolder.Checked = True
        txtNewFolder.Text = Folder
        Expedite = True
        Mode = BackupMode.bkBackup
        BackupFiles = Files 'bkPOS + bkSS + bkGL + bkPayroll + bkPayables + bkBanking
        'cmdStart.Value = True
        cmdStart.PerformClick()
        BackupTo = True
        IgnoreMissingRestores = False ' clear it
        BackupLog("BackupTo: COMPLETE Res=" & TrueFalseString(BackupTo))
        'Unload Me
        Me.Close()
    End Function

    Public Function UnzipFiles(ByVal ZipFile As String, ByVal DestDir As String, Optional ByVal DeleteContents As Boolean = True) As Boolean
        Printer.PSet(Me.ClientSize.Width - 600, 0) : Print(DescribeZipType(ZipType))

        Status = "Unzipping " & ZipFile & "..."
        Select Case modBackup.ZipType
            Case cdsZipType.wzt7ZIP : UnzipFiles = SevenZipUnZipFiles(ZipFile, DestDir, DeleteContents)
            Case cdsZipType.wztINFO : UnzipFiles = InfoZipUnzipFiles(ZipFile, DestDir, DeleteContents)
            Case cdsZipType.wztVJCZ : UnzipFiles = VJCUnzipFiles(ZipFile, DestDir, DeleteContents)
            Case cdsZipType.wztNone : Err.Raise(-1, , "No Valid unZip component.")
            Case Else : DevErr("frmBackupGeneric.UnzipFiles - Unknown Zip Component [" & modBackup.ZipType & "]")
        End Select

        If DestDir = InventFolder() Then CleanInventUnzip
    End Function

    Public Property Status() As String
        Get
            Status = sb.Panels(1).Text
        End Get
        Set(value As String)
            sb.Panels(1).Text = value
            sb.Refresh
            BackupLog("Backup Status: " & value)
            '  DoEvents
        End Set
    End Property

    Public Function InfoZipUnzipFiles(ByVal ZipFile As String, ByVal DestDir As String, Optional ByVal DeleteContents As Boolean = True) As Boolean
        Dim I As Integer, IsNew As Boolean
        On Error Resume Next

        If Microsoft.VisualBasic.Right(DestDir, 1) <> "\" Then DestDir = DestDir & "\"
        If Not FileExists(ZipFile) Then Exit Function
        If Not DirExists(DestDir) Then Exit Function

        If DeleteContents Then Kill(DestDir & "*.*")

        IsNew = False
        cU = New cZipUnzip
        cU.ZipFile = ZipFile
        cU.UnzipFolder = DestDir
        cU.ExtractOnlyNewer = False
        cU.OverwriteExisting = True
        cU.Directory()

        For I = 1 To cU.FileCount
            ' the new zip stores directories... we must handle it if we are looking at an older zip
            If cU.FileDirectory(I) <> "" Then IsNew = True
        Next

        cU.UseFolderNames = IsNew
        ZipFileCount = cU.FileCount
        BackupProgress(0, ZipFileCount, "Unzipping...")
        cU.Unzip()

        DisposeDA(cU)
        BackupProgress()
        InfoZipUnzipFiles = True
    End Function

    Public Function VJCUnzipFiles(ByVal ZipFile As String, ByVal DestDir As String, Optional ByVal DeleteContents As Boolean = True) As Boolean
        On Error GoTo CheckErr
        If Microsoft.VisualBasic.Right(DestDir, 1) <> "\" Then DestDir = DestDir & "\"

        If DeleteContents Then
            Status = "Cleaning directories..."
            Kill(DestDir & "*")  ' Delete old files.
        End If

        Status = "Performing restore..."

        '    NewZipBackup.GetZipFileInfo ZipFile
        '    NewZipBackup.FileUnCompressedSizeArray
        '    NewZipBackup.FileCompressedSizeArray

        NewZipBackup.Notify = False

        '    NewZipBackup.PreserveDirPath = True
        ' PreserveDirPath is added to handle subdirectory restoration..
        ' We're going to cross our fingers and hope nobody has to restore
        ' from a backup made during the "bad" period.
        ' No, we're going to leave it as it is..

        NewZipBackup.PreserveDirPath = True
        NewZipBackup.UnZipAllFiles(ZipFile, DestDir)
        NewZipBackup.ZipClose()

        Status = "Complete!"
        VJCUnzipFiles = True
        Exit Function

CheckErr:
        Select Case Err.Number
            Case 70 ' Permission denied.  This is a fatal error.
                MessageBox.Show("Error during restore!  Please close the program and make sure nothing is using the WinCDS files, then try again.", "Error!")
            Case 53 ' File not found, this is ok.
                Resume Next
            Case Else
                MessageBox.Show("Error during Restore (" & Err.Number & "): " & Err.Description)
                Exit Function
        End Select
    End Function

    Private Sub CleanInventUnzip()
        On Error GoTo Out
        If Dir(PhysicalInvFolder, vbDirectory) = "" Then MkDir(PhysicalInvFolder)
        If Dir(PhysicalInvOldFolder, vbDirectory) = "" Then MkDir(PhysicalInvOldFolder)

        Dim L As Object, A As String
        On Error Resume Next
        Kill(PhysicalInvFolder() & "*.*")
        On Error GoTo Out
        For Each L In AllFiles(InventFolder())
            If IsIn(UCase(Microsoft.VisualBasic.Left(L, 4)), "DISC") Then
                On Error Resume Next
                'Name  InventFolder() & L As PhysicalInvFolder() & L
                My.Computer.FileSystem.MoveFile(InventFolder() & L, PhysicalInvFolder() & L)
                On Error GoTo Out
            End If
        Next

        On Error Resume Next
        Kill(PhysicalInvOldFolder() & "*.*")
        On Error GoTo Out
        For Each L In AllFiles(InventFolder())
            If IsIn(UCase(Microsoft.VisualBasic.Left(L, 4)), "SIMP", "GROU") Then
                On Error Resume Next
                'Name(InventFolder() & L As PhysicalInvOldFolder & L)
                My.Computer.FileSystem.MoveFile(InventFolder() & L, PhysicalInvOldFolder() & L)
                On Error GoTo Out
            End If
        Next

Out:
    End Sub

    Public Sub BackupProgress(Optional ByVal N As Integer = -2, Optional ByVal Max As Integer = -1, Optional ByVal Str As String = "#", Optional ByVal DoShow As Boolean = False)
        ' we use this extra prg on this form because it's true modal (so our progress form wont show up on top of it)
        ' we use the frmProgress handling because it's already got the suitable algorithms for making all this work well reinventing everything

        If N = -2 Then
            fraExtra.Visible = False
            On Error Resume Next
            'Unload P
            P.Close()
            P = Nothing
            Exit Sub
        End If

        If P Is Nothing Then P = New frmProgress

        If N = -1 Then N = prgExtra.Value + 1
        'P.AltPrg = prgExtra
        fraExtra.Visible = True
        fraExtra.Text = Str

        P.Progress(N, Max, Str, True)
    End Sub
End Class