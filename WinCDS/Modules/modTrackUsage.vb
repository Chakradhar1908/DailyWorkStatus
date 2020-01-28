Module modTrackUsage
    Public Function TrackUsage(ByVal Title As String, Optional ByVal Desc As String) As Boolean
        Dim V As String
        V = Title & IIf(Desc = "", "", " - " & Desc)
        SetConfigLastRun Title
  PostUsage Title, Desc

  LogFile "Usage", Title, False

  TrackUsage = True
    End Function

End Module
