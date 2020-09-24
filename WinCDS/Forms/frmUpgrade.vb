Public Class frmUpgrade
    Private mCurrentVersion As String
    Public Function DoSilentUpdate(Optional ByRef Msg As String = "Your software has been updated.", Optional ByRef FileList As String = "") As Boolean
        UpdateLog("************************************")
        UpdateLog(SoftwareVersionForLog)
        UpdateLog("DoSilentUpdate")
        UpdateMsg = ""
        UpdateFileList = ""
        UpgradeNoMessages = True

        If ValidateUpdateFolder() = "" Then Exit Function
        UpdateLog("Getting Update List...")
        If Not GetUpdateList(CurrentVersionURL) Then Exit Function
        UpdateLog("Processing Updates...")
        DoSilentUpdate = CheckUpdateStatus(True) > 0
        UpdateLog("Processing Complete (SILENT)")
        UpgradeNoMessages = False
        Msg = "Your " & UpdateMsg & " has been updated."
        FileList = UpdateFileList
    End Function

    Public Function ValidateUpdateFolder(Optional ByVal TestFileWrite As Boolean = True) As String
        On Error Resume Next
        Dim TF As String

        ValidateUpdateFolder = UpdateFolder()
        MkDir(ValidateUpdateFolder)
        Kill(CurrentVersionLCL)
        'Kill ValidateUpdateFolder & "*"

        On Error Resume Next
        SetAttr(UpdateFolder, 0)
        SetAttr(AppFolder, 0)
        SetAttr(WinCDSEXEFile(True), 0)
        SetAttr(AppFolder() & "wait.exe", 0)

        If Not DirExists(ValidateUpdateFolder) Then ValidateUpdateFolder = "" : Exit Function

        If TestFileWrite Then
            If Not CanWriteToFolder(UpdateFolder) Then ValidateUpdateFolder = "" : Exit Function
            '    If Not CanWriteToFolder(AppFolder) Then ValidateUpdateFolder = "": Exit Function
        End If
    End Function

    Public Function GetUpdateList(ByVal nURL As String) As Boolean
        GetUpdateList = DownloadURLToFile(nURL, CurrentVersionLCL)
        mCurrentVersion = ReadFile(CurrentVersionLCL)
        GetUpdateList = mCurrentVersion <> ""
    End Function

    Private Function CheckUpdateStatus(ByVal DoUpdate As Boolean, Optional ByRef Msg As String = "") As Integer
        Dim objDoc As Object 'As MSXML2.DOMDocument40
        Dim objNodeList As Object 'As MSXML2.IXMLDOMNodeList
        Dim objNode As Object 'As MSXML2.IXMLDOMNode

        Dim FCount As Integer, UCount As Integer, FUrgent As Boolean
        Dim FProgram As String, FSpecVer As String, FMsg As String

        Dim Nm As String, DS As String, dT As String, Sz As String
        Dim LC As String, nT As String, VR As String, UR As String
        Dim RQ As String

        Dim Need As Boolean, Want As Boolean, Anyways As Boolean

        Dim DownloadedBytesSoFar As Integer

        On Error Resume Next
        '  Set objDoc = New MSXML2.DOMDocument30
        objDoc = CreateObject("MSXML2.DOMDocument")
        If objDoc Is Nothing Then
            CheckUpdateStatus = -2
            Exit Function
        End If

        If Not objDoc.LoadXml(mCurrentVersion) Then
            UMsgBox("Invalid WebUpdate Format", vbCritical, "Unable to update")
        End If
        objNode = objDoc.SelectSingleNode("WebUpdate")
        For Each objNode In objNode.Attributes
            Select Case objNode.BaseName
                Case "count" : FCount = Val(objNode.Text)
                Case "urgent" : FUrgent = (LCase(objNode.Text) = "yes")
                Case "program" : FProgram = objNode.Text
                Case "upgrade-xml-spec" : FSpecVer = objNode.Text
                Case "msg" : FMsg = objNode.Text
            End Select
        Next
        UCount = 0

        If FCount < 0 Then
            CheckUpdateStatus = -1
            Msg = FMsg
            Exit Function
        End If

        objNodeList = objDoc.SelectNodes("WebUpdate/file")
        'prgComplete.Min = 0   --> Code commented, because prgComplete is "ucPBar" custom active control. Most of this custom control is not supporting in vb.net. Find an alternative control to use it here.
        'If Not DoUpdate Then prgComplete.Max = 0
        'prgComplete.Value = 0

        For Each objNode In objNodeList
            Nm = "" : Nm = objNode.SelectSingleNode("name").Text
            DS = "" : DS = objNode.SelectSingleNode("desc").Text
            dT = "" : dT = objNode.SelectSingleNode("date").Text
            Sz = "" : Sz = objNode.SelectSingleNode("size").Text
            LC = "" : LC = objNode.SelectSingleNode("location").Text
            nT = "" : nT = objNode.SelectSingleNode("install").Text
            VR = "" : VR = objNode.SelectSingleNode("version").Text
            UR = "" : UR = objNode.SelectSingleNode("url").Text
            RQ = "" : RQ = objNode.SelectSingleNode("require").Text

            'If IsDevelopment And LCase(Nm) = "wincds.exe" Then Stop

            Need = FileNeedsUpdate(FUrgent, Nm, DS, dT, Sz, LC, nT, VR, UR, RQ)
            Want = FileIsSelected(Nm) Or UpgradeNoMessages
            ActiveLog("frmUpgrade::CheckUpdateStatus: F=" & Nm & ", NEED=" & Need & " WANT=" & Want)

            If DoUpdate Then
                If Want Then
                    Anyways = False ' Effectively blocks one path.  Possibility for furhter expansion.
                    If Not Need And Not Anyways Then
                        '          UMsgBox "You do not need to update " & Nm & vbCrLf & "It is already up to date!", vbExclamation, "Not Upgraded!"
                    Else
                        ActiveLog("frmUpgrade::CheckUpdateStatus: F=" & Nm & " -->ProcessUpdateFile")
                        ProcessUpdateFile(Nm, DS, dT, Sz, LC, nT, VR, UR)
                        DownloadedBytesSoFar = DownloadedBytesSoFar + Val(Sz)
                        UCount = UCount + 1
                    End If
                End If
            Else
                If Need Then
                    ActiveLog("frmUpgrade::CheckUpdateStatus: F=" & Nm & " ADDING TO lstFiles")
                    'lstFiles.AddItem AlignString(Nm, 20, vbAlignLeft, True) & " - " & IIf(Need, "Needs Update" & " - " & GetInstallDir(LC), "*Up To Date!")
                    lstFiles.Items.Add(AlignString(Nm, 20, VBRUN.AlignConstants.vbAlignLeft, True) & " - " & IIf(Need, "Needs Update" & " - " & GetInstallDir(LC), "*Up To Date!"))
                    'lstFiles.Selected(lstFiles.NewIndex) = True
                    lstFiles.SetSelected(lstFiles.SelectedIndex, True)
                    UCount = UCount + 1
                    'prgComplete.Max = prgComplete.Max + Val(Sz)  -->Comented, because "prgComplete" is ucPBar custom activex control which most of its code is not supporting in vb.net
                End If
            End If

            If DoUpdate Then
                'prgComplete.Value = DownloadedBytesSoFar
            End If
        Next

        objNode = Nothing
        objNodeList = Nothing
        objDoc = Nothing

        CheckUpdateStatus = UCount
    End Function

    Private Function FileNeedsUpdate(ByVal Urgent As Boolean, ByVal fName As String, ByVal FDesc As String, ByVal FDate As String, ByVal FSize As String, ByVal FLoca As String, ByVal FInst As String, ByVal FVers As String, ByVal FURL As String, ByVal Require As String) As Boolean
        Dim InstallDir As String, P As String, X As Integer
        Dim IsDifferent As Boolean, NeedInstall As Boolean
        InstallDir = GetInstallDir(FLoca)
        If InstallDir = "" Then Exit Function 'nothing to do!
        If Not DirExists(InstallDir) Then Exit Function
        If Require <> "" Then FileNeedsUpdate = True : Exit Function

        ' in Silent mode, we reduce load on the server..  usually will only get
        ' upgrades once a week.
        'If False And IsCDSComputer("LAPTOP") And LCase(fname) = "wincds.exe" Then
        '  MsgBox "WinCDS Check (WinCDS Laptop Only):" & vbCrLf2 & "UpgradeNoMessages: " & UpgradeNoMessages & vbCrLf & "Urgent: " & Urgent & vbCrLf & "Scheduled?" & ScheduledUpdateToday & vbCrLf2 & "Proceed? " & Not (UpgradeNoMessages And Not Urgent And Not ScheduledUpdateToday)
        '  If IsIDE Then Stop
        'End If

        If UpgradeNoMessages And Not Urgent And Not ScheduledUpdateToday() Then Exit Function

        P = InstallDir & fName

        NeedInstall = False
        IsDifferent = False
        If Not FileExists(P) Then
            IsDifferent = True
            NeedInstall = True
        Else
            If IsDate(FDate) Then
                X = DateDiff("d", DateValue(FileDateTime(P)), DateValue(FDate))
                If X <> 0 Then
                    IsDifferent = True
                    If X > 0 Then NeedInstall = True
                End If
            End If
            If Val(FSize) > 0 Then
                If FileLen(P) <> Val(FSize) Then IsDifferent = True
                '      If FileLen(P) = Val(FSize) Then IsDifferent = False: NeedInstall = False
            End If
            If Len(FVers) > 0 Then
                If FVers <> FileVersion(P) Then IsDifferent = True : NeedInstall = True
            End If
        End If

        FileNeedsUpdate = NeedInstall

        ' if we've already downloaded the most current version but it's simply not installing, why not ignore it and not download it again...
        If FileNeedsUpdate And FileExists(UpdateFolder() & fName) And Not IsIn(LCase(fName), "wincds.exe", "wait.exe") Then
            Dim CV As Date ' currently downloaded version
            CV = DateValue(FileDateTime(UpdateFolder() & fName))
            If DateAfter(CV, DateValue(FDate)) And DateBefore(Today, DateAdd("d", 7, CV)) Then FileNeedsUpdate = False
        End If
    End Function

    Private Function FileIsSelected(ByVal FileName As String) As Boolean
        Dim I As Integer
        For I = 0 To lstFiles.Items.Count - 1
            If Trim(Microsoft.VisualBasic.Left(lstFiles.GetItemText(lstFiles.SelectedIndex), 20)) = Microsoft.VisualBasic.Left(FileName, 20) Then
                'If lstFiles.Selected(I) Then
                If lstFiles.GetSelected(I) Then
                    FileIsSelected = True
                    Exit Function
                End If
            End If
        Next
    End Function

    Private Function ProcessUpdateFile(ByVal fName As String, ByVal FDesc As String, ByVal FDate As String, ByVal FSize As String, ByVal FLoca As String, ByVal FInst As String, ByVal FVers As String, ByVal FURL As String) As Boolean
        Dim InstallDir As String, Rsn As String
        InstallDir = GetInstallDir(FLoca)
        If InstallDir = "" Then Exit Function    'nothing to do!

        On Error Resume Next
        If Not UpgradeNoMessages Then ProgressForm(-2, 1, "Downloading && Installing File: " & fName & "...")

        '  Set iB = New clsMyIBindCallback
        '  If Not iB.DownloadFileProgress(FURL, UpdateFolder & FName) Then
        '    UMsgBox "Could not download " & FURL & vbCrLf2 & "Msg: " & Rsn & vbCrLf & "Please feel free to try this upgrade again.", vbExclamation, "Upgrade failed!"
        '  Else
        '    InstallUpgrade FName, UpdateFolder, InstallDir, FInst
        '  End If
        '  Set iB = Nothing
        If Not DownloadURLToFile(URLEncode(FURL), UpdateFolder() & fName, Val(FSize), Rsn, prgCurrentFile, prgComplete) Then  'Note: prgCurrentFile and prgComplete are ucPBar controls. Use alternative controls.
            UMsgBox("Could not download " & FURL & vbCrLf2 & "Msg: " & Rsn & vbCrLf & "Please feel free to try this upgrade again.", vbExclamation, "Upgrade failed!")
        Else
            InstallUpgrade(fName, UpdateFolder, InstallDir, FInst)
        End If

        If Not UpgradeNoMessages Then ProgressForm()
        ProcessUpdateFile = True
    End Function

    Public Sub NotifyUpgrade(Optional ByVal Notify As Boolean = False)
        On Error Resume Next
        If Notify Then frmUpgradeNotify.Notify("Your program has been updated.")
        DownloadURLToFile(NotifyUpgradeURL, CurrentVersionLCL)
        'Unload Me
        Me.Close()
    End Sub

    Public Sub DoCommandLineUpdate()
        CommandLineUpdate = True
        UpgradeNoMessages = True
        DoSilentUpdate()
    End Sub
End Class