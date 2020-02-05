Public Class frmAutoWeb
    Private BuildDir As String
    Public FOwner As Form

    Public Function BuildCSV() As String
        BuildCSV = BuildDir & DefaultCSVFile
        Generate2DataCSV(BuildCSV, txtSiteAddr.Text)
        BuildLog("Created " & DefaultCSVFile & " (" & BuildCSV & ")")
    End Function

    Public ReadOnly Property DefaultCSVFile()
        Get
            DefaultCSVFile = "2data.csv"
        End Get
    End Property

    Private Sub BuildLog(ByVal Msg As String)
        Dim FN As Integer

        FN = FreeFile()
        'Open(BuildDir & LogFileName For Append As #FN)
        FileOpen(FN, BuildDir & LogFileName, OpenMode.Append)
        'Print(#FN, "" & Now & ": " & Msg)
        Print(FN, "" & Now & ": " & Msg)
        'Close(#FN)
        FileClose(FN)
    End Sub

    Public ReadOnly Property LogFileName()
        Get
            LogFileName = "SiteBuild.log"
        End Get
    End Property

    Public Function SiteDepartmentURL(ByVal DeptName As String, Optional ByVal PageNum as integer = 1) As String
        SiteDepartmentURL = "/dept/" & ProtectFileName(DeptName) & PageNum & ".html"
    End Function

    Public Function ProtectFileName(ByVal FN As String) As String
        FN = Replace(FN, ".", "_")
        FN = Replace(FN, "*", "_")
        FN = Replace(FN, "/", "_")

        FN = Replace(FN, " ", "")
        FN = Replace(FN, "&", "")
        FN = Replace(FN, "\", "")
        FN = Replace(FN, """", "")
        FN = Replace(FN, "'", "")
        ProtectFileName = LCase(FN)
    End Function

End Class