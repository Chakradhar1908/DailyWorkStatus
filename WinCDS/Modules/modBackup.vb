Imports Microsoft.VisualBasic.Interaction
Public Module modBackup
    Public Enum BackupMode
        bkBackup = 0
        bkRestore = 1
    End Enum
    Public Enum BackupType
        ' In case we want to specify this.
        bkNone = 0

        ' The 4 accounting modules (Bank, GenLedgr, Payroll, Accounts Payable)
        bkBK = 1
        bkGL = 2
        bkPR = 4
        bkAP = 8

        ' The 3 WinCDS modules (POS - GM and invent databases, Store Setups, InventPX Folder)
        bkPS = 16           ' POS
        bkSS = 32           ' Store Setup
        bkpx = 64
        bkfx = 128

        ' Other folders
        bkQB = 256          ' Quickbooks
        bkmc = 512          ' Misc

        bkLO = 1024         ' Logs


        ' Future Expansion

        bkXb12 = 2048
        bkXb13 = 4096
        bkXb14 = 8192
        bkXb15 = 16384
        bkXb16 = 32768

        ' This will currently catch all of the above (15 bits, 10 used + 5 unused)
        bkAll = 32767
        '  bkAll = 65535         ' VB6 uses signed numbers..  we will just hold off unless we need it.
    End Enum
    Public Enum cdsZipType
        wztNone = 0
        wztVJCZ = 1
        wztINFO = 2
        wzt7ZIP = 3
    End Enum

    Public Function ZipFiles(ByVal CompressPath As String, ByVal ZipDir As String, ByVal ZipFile As String, Optional ByVal Special As Integer = 0) As Boolean
        Dim UnloadAfter As Boolean
        UnloadAfter = Not IsFormLoaded("frmBackupGeneric")
        ZipFiles = frmBackUpGeneric.ZipFiles(CompressPath, ZipDir, ZipFile, Special)
        If UnloadAfter Then
            'Unload frmBackUpGeneric
            frmBackUpGeneric.Close()
        End If
    End Function

    Public Sub BackupLog(ByVal vMsg As String)
        Dim T As String
        If IsFormLoaded("frmBackupGeneric") Then
            T = Switch(frmBackUpGeneric.Mode = BackupMode.bkBackup, "B", frmBackUpGeneric.Mode = BackupMode.bkRestore, "R", True, "!")
        Else
            T = "?"
        End If
        LogFile("Backup.txt", T & " " & vMsg, False)
    End Sub

    Public Function DescribeBackupType(ByVal Files As BackupType) As String
        Dim X As String

        If Files = BackupType.bkAll Then
            X = "-- ALL --"
        ElseIf Files = BackupType.bkNone Then
            X = "-- NONE --"
        Else
            If (Files And BackupType.bkBK) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkBK"
            If (Files And BackupType.bkGL) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkGL"
            If (Files And BackupType.bkPR) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkPR"
            If (Files And BackupType.bkAP) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkAP"
            If (Files And BackupType.bkPS) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkPS"
            If (Files And BackupType.bkSS) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkSS"
            If (Files And BackupType.bkpx) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkPX"
            If (Files And BackupType.bkfx) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkFX"
            If (Files And BackupType.bkQB) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkQB"
            If (Files And BackupType.bkmc) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkMC"
            If (Files And BackupType.bkLO) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkLO"

            If (Files And BackupType.bkXb12) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkXb12"
            If (Files And BackupType.bkXb13) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkXb13"
            If (Files And BackupType.bkXb14) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkXb14"
            If (Files And BackupType.bkXb15) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkXb15"
            If (Files And BackupType.bkXb16) <> 0 Then X = X & IIf(X = "", "", " & ") & "bkXb16"
        End If

        DescribeBackupType = "[ " & X & " ]"
    End Function

    Public Function BackupTo(ByVal Folder As String, Optional ByVal Files As BackupType = BackupType.bkAll) As Boolean
        Dim UnloadAfter As Boolean
        UnloadAfter = Not IsFormLoaded("frmBackupGeneric")
        BackupTo = frmBackUpGeneric.BackupTo(Folder, Files)
        'If UnloadAfter Then Unload frmBackUpGeneric
        If UnloadAfter Then frmBackUpGeneric.Close()
    End Function

    Public ReadOnly Property ZipType() As cdsZipType
        Get
            If USE_SEVENZIP_ZIP Then ZipType = cdsZipType.wzt7ZIP : Exit Property
            If USE_INFO_ZIP Then ZipType = cdsZipType.wztINFO : Exit Property
            If USE_VJC_ZIP Then ZipType = cdsZipType.wztVJCZ : Exit Property
            ZipType = cdsZipType.wztNone
        End Get
    End Property

    Public ReadOnly Property USE_VJC_ZIP() As Boolean
        Get
            If USE_INFO_ZIP Then Exit Property
            USE_VJC_ZIP = FileExists(System32Folder(True) & "vjczip.ocx")
        End Get
    End Property

    Public ReadOnly Property USE_INFO_ZIP() As Boolean
        Get
            USE_INFO_ZIP = FileExists(AppFolder() & "vbzip11.dll") And FileExists(AppFolder() & "vbuzip10.dll")
            'BFH20150128 - Apparently, Infozip causes many errors..  we'll leave it on for dev
            If Not IsDevelopment() Then USE_INFO_ZIP = False
            '  USE_INFO_ZIP = False
        End Get
    End Property

    Public ReadOnly Property USE_SEVENZIP_ZIP() As Boolean
        Get
            USE_SEVENZIP_ZIP = FileExists(AppFolder() & "7za.exe")
            '  If Not IsDevelopment Then USE_SEVENZIP_ZIP = False
        End Get
    End Property

    Public Function DescribeZipType(ByVal ZT As cdsZipType) As String
        Select Case ZT
            Case cdsZipType.wztVJCZ : DescribeZipType = "VJCZip" : Exit Function
            Case cdsZipType.wztINFO : DescribeZipType = "InfoZp" : Exit Function
            Case cdsZipType.wzt7ZIP : DescribeZipType = "SevenZ" : Exit Function
            Case cdsZipType.wztNone : DescribeZipType = "[None]" : Exit Function
            Case Else : DevErr("modBackup.DescribeZipType() - Invalid Zip Type [" & ZT & "]")
        End Select
    End Function

    Public Function SevenZipUnZipFiles(ByVal ZipFile As String, ByVal DestDir As String, Optional ByVal DeleteContents As Boolean = True) As Boolean
        On Error Resume Next

        If Right(DestDir, 1) <> "\" Then DestDir = DestDir & "\"
        If Not FileExists(ZipFile) Then Exit Function
        If Not DirExists(DestDir) Then Exit Function

        If DeleteContents Then Kill(DestDir & "*.*")

        Dim R As String, C As String, E As String, A As String
        ' command line options used:
        '   x  (command, specifies extract with paths)
        '   zipfile == archive to extract from
        '   destdir == where to unzip them to
        '
        '   -y force "Yes" to all prompts

        '  C = "7za x """ & ZipFile & """ -o" & DestDir & " -y"
        A = SevenZip7ZA
        C = " x """ & ZipFile & """ -o" & DestDir & " -y"
        'Debug.Print "modBackup.SZ Restore Command: " & C
        R = RunCmdToOutputWithArgs(A, C, E)
        If E <> "" Then BackupLog("Restore Error String (SevenZipUnZipFiles): " & E)

        SevenZipUnZipFiles = True

        If Not SevenZipUnZipFiles Then BackupLog("Restore Failed (SevenZipUnZipFiles): " & R)
    End Function

    Public Function SevenZip7ZA() As String
        Dim A As String

        A = AppFolder() & "7za.exe"
        If FileExists(A) Then SevenZip7ZA = A : Exit Function

        A = LocalProgramFilesFolder() & "WinCDS\7za.exe"
        If FileExists(A) Then SevenZip7ZA = A : Exit Function

        A = LocalProgramFilesFolder(True) & "WinCDS\7za.exe"
        If FileExists(A) Then SevenZip7ZA = A : Exit Function

        SevenZip7ZA = "7za"
        If FileExists(A) Then SevenZip7ZA = A : Exit Function
    End Function

End Module