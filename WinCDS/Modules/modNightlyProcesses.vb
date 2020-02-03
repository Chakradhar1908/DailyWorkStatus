Module modNightlyProcesses
    Public Sub NightlyProcesses()
        If AllowRaycom And DayOfWeek(Of Date)() = "Sunday" Then RaycomNightlyReport
    End Sub

End Module
