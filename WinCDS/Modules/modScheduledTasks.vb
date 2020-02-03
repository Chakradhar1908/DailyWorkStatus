Module modScheduledTasks
    Public Function CheckScheduledTasks(Optional ByVal ForceReset As Boolean = False, Optional ByVal Remove As Boolean = False) As Boolean
        ' Only check for the install
        If IsDemo() Then Exit Function              ' Demo software should not install scheduled tasks.  Once they have a valid license, the next authorized start should install them.
        If IsWin5() Then Exit Function              ' Task Scheduler 1.0 isn't compatible with this code...  We also don't need it, because there's no UAC.
        If Not IsElevated() Then Exit Function      ' If it is a Task Scheduler 2.0 machine, and they're not elevated, it will fail because of UAC.

        If Not ServiceMode Then
            If DateBefore(Of Date, "6/30/2016")() Then ForceReset = True
        End If

        If UseScheduledTask Then CheckScheduledTask TASK_SERVICE_NAME, ForceReset, Remove

'  CheckScheduledTask TASK2_SERVICE_NAME, ForceReset, Remove
        '  CheckScheduledTask TASK4_SERVICE_NAME, ForceReset, Remove

        ' BFH20160624 - We don't need an AutoVNC launch at all now
        '  If Not IsDevelopment Then Exit Function
        '  CheckScheduledTask TASK3_SERVICE_NAME, ForceReset, Remove
    End Function

End Module
