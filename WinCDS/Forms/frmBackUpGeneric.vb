Public Class frmBackUpGeneric
    Dim intLastDrive As Integer
    Private mBackupMode As Integer, mBackupFiles As BackupType
    'Private WithEvents cZ As cZip
    'Private WithEvents cU As cZipUnzip
    Private ZipFileCount As Integer
    Public Expedite As Boolean
    Public IgnoreMissingRestores As Boolean
    Private P As frmProgress

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
End Class