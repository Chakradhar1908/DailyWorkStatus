Module modNightlyProcesses
    Public Sub NightlyProcesses()
        If AllowRaycom() And DayOfWeek(Today) = "Sunday" Then RaycomNightlyReport
    End Sub

End Module
