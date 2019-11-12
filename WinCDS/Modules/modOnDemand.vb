Module modOnDemand
    Private Enum OnDemandType
        odtExists
        odtCreateObject
        odtControl
        odtFont
    End Enum

    Private Structure OnDemandDef
        Dim vType As OnDemandType
        Dim Reference As String
        Dim FileName As String
        Dim Location As String
        Dim Install As String
    End Structure
    Private ODList() As OnDemandDef, ODInit As Boolean
    Public OnDemandUpdate As Boolean
    Public OfflineUpdate As Boolean

    Public Function Gif89Installed() As Boolean
        If Val(GetCDSSetting("HasGif89")) <> 0 Then Gif89Installed = True : Exit Function

        If OnDemand_Test(OnDemandEntry("Gif89.Gif89.1")) Then
            SaveCDSSetting("HasGif89", "1")
            Gif89Installed = True
            Exit Function
        End If

        Gif89Installed = False
    End Function

    Private Function OnDemandEntry(ByVal ControlName As String) As OnDemandDef
        Dim I As Integer
        initOnDemand()

        For I = LBound(ODList) To UBound(ODList)
            If ODList(I).Reference = ControlName Or ODList(I).FileName = ControlName Then OnDemandEntry = ODList(I) : Exit Function
        Next
        DevErr("modOnDemand.OnDemandEntry: Unknown Ref: " & ControlName)
    End Function

    Private Function OnDemand_Test(D As OnDemandDef) As Boolean
        On Error GoTo NoControl
        Dim C As Object

        Select Case D.vType
            Case OnDemandType.odtControl
                'Load PracticeOnDemandControl
                PracticeOnDemandControl.Show()
                'C = PracticeOnDemandControl.Controls.Add(D.Reference, "testcontrol_" & DateTimeStamp() & "_" & Random(1000))
                Dim L1, L2 As New Label
                L1.Name = D.Reference
                PracticeOnDemandControl.Controls.Add(L1)
                L2.Name = "testcontrol_" & DateTimeStamp() & "_" & Random(1000)
                PracticeOnDemandControl.Controls.Add(L2)

                'If C Is Nothing Then GoTo NoControl
                If PracticeOnDemandControl.Controls.Count = 0 Then GoTo NoControl
                'Unload PracticeOnDemandControl
                PracticeOnDemandControl.Close()
                OnDemand_Test = True
            Case OnDemandType.odtCreateObject
                C = CreateObject(D.Reference)
                If Not (C Is Nothing) Then OnDemand_Test = True
            Case OnDemandType.odtFont
                OnDemand_Test = FontExists(D.Reference)
            Case OnDemandType.odtExists
                OnDemand_Test = FileExists(GetInstallDir(D.Location) & D.FileName)
            Case Else
                DevErr("modOnDemand.OnDemand_Test: Unknown OnDemand Type")
        End Select

        Exit Function

NoControl:
        'Unload PracticeOnDemandControl
        PracticeOnDemandControl.Close()
    End Function

    Private Sub initOnDemand()
        Dim I As Integer
        If ODInit Then Exit Sub
        ODInit = True
        For I = BarCodeFonts.bcfFirst To BarCodeFonts.bcfLast
            AddOnDemand(OnDemandType.odtFont, BarCodeFontName(I), BarCodeFontFile(I), "$FontDir", "$fontregister")
        Next

        AddOnDemand(OnDemandType.odtControl, "LaVolpeAlphaImg.AlphaImgCtl", "LaVolpeAlphaImg2.ocx", "$WinSysDir", "$dllselfregister")
        AddOnDemand(OnDemandType.odtControl, "MSChart20Lib.MSChart", "MSChrt20.ocx", "$WinSysDir", "$dllselfregister")
        AddOnDemand(OnDemandType.odtControl, "Gif89.Gif89.1", "gif89.dll", "$WinSysDir", "$dllselfregister")

        '  AddOnDemand odtControl, "InetCtls.Inet.1", "MSINET.ocx", "$WinSysDir", "$dllselfregister"
        '  AddOnDemand odtCreateObject, "SSubTimer6.CTimer", "ssubtmr6.dll", "$WinSysDir", "$dllselfregister"

        AddOnDemand(OnDemandType.odtExists, "", "ChilkatAx-9.5.0-win32.DLL", "$WinSysDir", "$DLLSelfRegister")
        AddOnDemand(OnDemandType.odtExists, "", "VBCCR14.OCX", "$WinSysDir", "$DLLSelfRegister")
        AddOnDemand(OnDemandType.odtExists, "", "FreeImage.dll", "$WinSysDir", "")
    End Sub

    Private Sub AddOnDemand(ByVal vTestType As OnDemandType, ByVal vReference As String, ByVal vFileName As String, ByVal vLocation As String, ByVal vInstall As String)
        Dim X As Integer
        On Error Resume Next
        X = UBound(ODList)
        X = X + 1
        'ReDim Preserve ODList(1 To X)
        ReDim Preserve ODList(0 To X - 1)
        On Error GoTo 0

        'ODList(X).vType = vTestType
        'ODList(X).Reference = vReference
        'ODList(X).FileName = vFileName
        'ODList(X).Location = vLocation
        'ODList(X).Install = vInstall

        ODList(X - 1).vType = vTestType
        ODList(X - 1).Reference = vReference
        ODList(X - 1).FileName = vFileName
        ODList(X - 1).Location = vLocation
        ODList(X - 1).Install = vInstall
    End Sub

End Module