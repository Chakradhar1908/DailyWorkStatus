Public Class frmAutoWeb
    Public FOwner As Form
    Public Function BuildCSV() As String
        BuildCSV = BuildDir & DefaultCSVFile
        Generate2DataCSV BuildCSV, txtSiteAddr
  BuildLog "Created " & DefaultCSVFile & " (" & BuildCSV & ")"
End Function

End Class