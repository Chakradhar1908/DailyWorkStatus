Module modUpgrade
    Public Function GetInstallDir(ByVal vKEY As String) As String
        Dim R As String
        If Left(vKEY, 1) = "$" Then
            Select Case LCase(vKEY)
                Case "$apppath", "$appfolder" : GetInstallDir = AppFolder()
                Case "$winsysdir", "$winsyspath"
                    R = GetWindowsDir & "\SysWOW64"
                    If DirExists(R) Then
                        GetInstallDir = R
                    Else
                        GetInstallDir = GetWindowsSystemDir
                    End If
                Case "$windir", "$winpath" : GetInstallDir = GetWindowsDir()
                Case "$fontdir", "$winfontdir" : GetInstallDir = SpecialFolder(FolderEnum.feWindowsFonts)
                Case "$px" : GetInstallDir = PXFolder()
                Case "$fx" : GetInstallDir = FXFolder()
                Case "$fxcontrol" : GetInstallDir = FXControlFolder()
                Case "$fxwallpaper" : GetInstallDir = FXWallpaperFolder
                Case "$fxtaglayout" : GetInstallDir = TagLayoutFolder
                Case "$invent" : GetInstallDir = InventFolder()
                Case "$store1" : GetInstallDir = StoreFolder(1)
                Case "$store2" : GetInstallDir = StoreFolder(2)
                Case "$store3" : GetInstallDir = StoreFolder(3)
                Case "$store4" : GetInstallDir = StoreFolder(4)
                Case "$store5" : GetInstallDir = StoreFolder(5)
                Case "$store6" : GetInstallDir = StoreFolder(6)
                Case "$store7" : GetInstallDir = StoreFolder(7)
                Case "$store8" : GetInstallDir = StoreFolder(8)
                Case "$store9" : GetInstallDir = StoreFolder(9)
                Case "$store10" : GetInstallDir = StoreFolder(10)
                Case "$store11" : GetInstallDir = StoreFolder(11)
                Case "$store12" : GetInstallDir = StoreFolder(12)
                Case "$store13" : GetInstallDir = StoreFolder(13)
                Case "$store14" : GetInstallDir = StoreFolder(14)
                Case "$store15" : GetInstallDir = StoreFolder(15)
                Case "$store16" : GetInstallDir = StoreFolder(16)
                Case "$store17" : GetInstallDir = StoreFolder(17)
                Case "$store18" : GetInstallDir = StoreFolder(18)
                Case "$store19" : GetInstallDir = StoreFolder(19)
                Case "$store20" : GetInstallDir = StoreFolder(20)
                Case "$store21" : GetInstallDir = StoreFolder(21)
                Case "$store22" : GetInstallDir = StoreFolder(22)
                Case "$store23" : GetInstallDir = StoreFolder(23)
                Case "$store24" : GetInstallDir = StoreFolder(24)
                Case "$store25" : GetInstallDir = StoreFolder(25)
                Case "$store26" : GetInstallDir = StoreFolder(26)
                Case "$store27" : GetInstallDir = StoreFolder(27)
                Case "$store28" : GetInstallDir = StoreFolder(28)
                Case "$store29" : GetInstallDir = StoreFolder(29)
                Case "$store30" : GetInstallDir = StoreFolder(30)
                Case "$store31" : GetInstallDir = StoreFolder(31)
                Case "$store32" : GetInstallDir = StoreFolder(32)
                Case "$cdsdata" : GetInstallDir = CDSDataFolder()
                Case "$qbcommon" : GetInstallDir = LocalProgramFilesFolder() & "Common Files\Intuit\QuickBooks\"
                Case "$pr" : GetInstallDir = PRFolder(, True)
                Case "$gl" : GetInstallDir = GLFolder(, True)
                Case "$ap" : GetInstallDir = APFolder(, True)
                Case "$bk" : GetInstallDir = BKFolder(, True)
                Case Else : GetInstallDir = ""
            End Select
        Else
            If vKEY = "" Then vKEY = AppFolder()   ' default to home dir
            GetInstallDir = vKEY
        End If
        GetInstallDir = CleanDir(GetInstallDir)
    End Function

End Module
