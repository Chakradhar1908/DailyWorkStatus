Module modCodeDistribution
    Public Const WEB_WEBUPDATE_USER As String = "webupdate@wincdspro.com"
    Public Const WEB_WEBUPDATE_PASS As String = "4XUkggTcRqRt"
    Public Const WEB_UPLOAD_USER As String = "upload@wincdspro.com"
    Public Const WEB_UPLOAD_PASS As String = "3Uv94BBAzS5p"
    Public Const WEB_AUTODOC_USER As String = "autodoc@wincdspro.com"
    Public Const WEB_AUTODOC_PASS As String = "NeFNXP$w0RAT"

    Public Function DistributionCSV(Optional ByVal vYear As String = "") As String
        DistributionCSV = AppFolder() & "distribution-" & DistributionYEAR(vYear) & ".csv"
    End Function

    Private Function DistributionYEAR(Optional ByVal vYear As String = "") As Integer
        DistributionYEAR = FitRange(2015, IIf(Val(vYear) = 0, Year(Now), Val(vYear)), Year(Now))
    End Function
End Module
