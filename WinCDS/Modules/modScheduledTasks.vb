Module modScheduledTasks
    Private Const TASK_SERVICE_NAME As String = "WinCDS Nightly Service"
    Private Const TASK2_SERVICE_NAME As String = "WinCDSUpgrade"
    Private Const TASK4_SERVICE_NAME As String = "WinCDSLaunch"
    Private Const TASK4_SERVICE_DESC As String = "Allows admin launch of WinCDS without prompt."
    Private Const cWake As Boolean = True
    Private Const TASK_ACTION_EXEC As Integer = 0
    Private Const TASK_CREATE_OR_UPDATE As Integer = 6
    Private Const TASK_LOGON_3 As Integer = 3
    Private Const TASK2_SERVICE_DESC As String = "Allows upgrade of WinCDS without elevation."
    Private Const TASK2_SERVICE_ARGS As String = "-x -y -q -mService"
    Private Const TASK_SERVICE_DESC As String = "Perform nightly maintenance including updates and cloud backup (if available)."
    Private Const TASK_TRIGGER_DAILY As Integer = 2
    Private Const TASK_SERVICE_TIME As String = "1:00a"
    Private Const TASK_SERVICE_ARGS As String = "/Service"

    Public Function CheckScheduledTasks(Optional ByVal ForceReset As Boolean = False, Optional ByVal Remove As Boolean = False) As Boolean
        ' Only check for the install
        If IsDemo() Then Exit Function              ' Demo software should not install scheduled tasks.  Once they have a valid license, the next authorized start should install them.
        If IsWin5() Then Exit Function              ' Task Scheduler 1.0 isn't compatible with this code...  We also don't need it, because there's no UAC.
        If Not IsElevated() Then Exit Function      ' If it is a Task Scheduler 2.0 machine, and they're not elevated, it will fail because of UAC.

        If Not ServiceMode Then
            If DateBefore(Today, "6/30/2016") Then ForceReset = True
        End If

        If UseScheduledTask Then CheckScheduledTask(TASK_SERVICE_NAME, ForceReset, Remove)

        '  CheckScheduledTask TASK2_SERVICE_NAME, ForceReset, Remove
        '  CheckScheduledTask TASK4_SERVICE_NAME, ForceReset, Remove

        ' BFH20160624 - We don't need an AutoVNC launch at all now
        '  If Not IsDevelopment Then Exit Function
        '  CheckScheduledTask TASK3_SERVICE_NAME, ForceReset, Remove
    End Function

    Public Function CheckScheduledTask(ByVal TaskName As String, Optional ByVal ForceReset As Boolean = False, Optional ByVal Remove As Boolean = False) As Boolean
        On Error GoTo QuickExit
        If ForceReset Or Remove Then ScheduledTaskDeleteTask(TaskName)
        If Remove Then Exit Function
        If Not VerifyScheduledTask(TaskName) Then CreateWinCDSTask(TaskName)
        CheckScheduledTask = True
QuickExit:
        Err.Clear()
    End Function

    Public Function CreateWinCDSTask(ByVal TaskName As String)
        Select Case TaskName
            Case TASK_SERVICE_NAME : CreateWinCDSServiceScheduledTask()
            Case TASK2_SERVICE_NAME : CreateWaitScheduledTask()
'    Case TASK3_SERVICE_NAME: CreateAutoVNCScheduledTask
            Case TASK4_SERVICE_NAME : CreateWinCDSLaunchScheduledTask()
            Case Else : DevErr("Cannot Create Scheduled Task: " & TaskName)
        End Select
    End Function

    Private ReadOnly Property LOGFILEID() As String
        Get
            LOGFILEID = DateTimeStamp() & Random(999999)
        End Get
    End Property

    Private Function CreateWinCDSLaunchScheduledTask() As Boolean
        Dim strId As String
        Dim TaskDef As Object 'TaskScheduler.ITaskDefinition

        strId = LOGFILEID ' Replace$(txtLogFile.Text, ".", "-") 'Create an Id value, no periods here!
        With CreateObject("Schedule.Service") 'New TaskScheduler.TaskScheduler
            .Connect

            ' This method call forces us to use late binding here, because .NewTask()
            ' takes an Unsupported Variant type argument "flags" (UInt?).  It has to
            ' be 0 anyway, so this works fine:
            TaskDef = .NewTask(0)

            With TaskDef
                With .RegistrationInfo
                    .Description = TASK4_SERVICE_DESC
                    .Author = SoftwareVersion(True, False, True)
                    .URI = CompanyURL
                    .Source = ProgramShort
                    .Version = "1.0"
                End With
                With .Settings
                    .AllowDemandStart = True
                    .AllowHardTerminate = True

                    .Enabled = True
                    .StartWhenAvailable = True
                    .WakeToRun = cWake
                    .Hidden = False
                End With
                With .Actions.Create(TASK_ACTION_EXEC)
                    .Path = "%windir%\System32\cmd.exe"
                    .WorkingDirectory = GetFilePath(TASK4_SERVICE_PATH)
                    .Arguments = "/c start ""WinCDS POS Software"" """ & WinCDSEXEFile(True, True, True) & """"
                End With
                With .Principal
                    '        Dim ID
                    .RunLevel = 1
                    '        .LogonType = 2 ' Run without being logged on.
                End With
            End With

            ' this one creates the task
            With .GetFolder("\")
                .RegisterTaskDefinition(TASK4_SERVICE_NAME, TaskDef, TASK_CREATE_OR_UPDATE, , , TASK_LOGON_3)
            End With
        End With

        CreateWinCDSLaunchScheduledTask = True
    End Function

    Private ReadOnly Property TASK4_SERVICE_PATH()
        Get
            TASK4_SERVICE_PATH = WinCDSEXEFile(True, True, True)
        End Get
    End Property

    Private ReadOnly Property TASK2_SERVICE_PATH()
        Get
            TASK2_SERVICE_PATH = WaitEXEFile(True)
        End Get
    End Property

    Private Function CreateWaitScheduledTask() As Boolean
        '  If Not IsServer Then Exit Function
        'https://msdn.microsoft.com/en-us/library/windows/desktop/aa446862(v=vs.85).aspx
        ' also
        'http://www.vbforums.com/showthread.php?636846-VB6-Use-Vista-Task-Scheduler-2-0-API

        Dim strId As String
        Dim TaskDef As Object 'TaskScheduler.ITaskDefinition

        strId = LOGFILEID ' Replace$(txtLogFile.Text, ".", "-") 'Create an Id value, no periods here!
        With CreateObject("Schedule.Service") 'New TaskScheduler.TaskScheduler
            .Connect

            ' This method call forces us to use late binding here, because .NewTask()
            ' takes an Unsupported Variant type argument "flags" (UInt?).  It has to
            ' be 0 anyway, so this works fine:
            TaskDef = .NewTask(0)

            With TaskDef
                With .RegistrationInfo
                    .Description = TASK2_SERVICE_DESC
                    .Author = SoftwareVersion(True, False, True)
                    .URI = CompanyURL
                    .Source = ProgramShort
                    .Version = "1.0"
                End With
                With .Settings
                    .AllowDemandStart = True      ' This is the entire point of the command...
                    .AllowHardTerminate = True

                    .Enabled = True
                    .StartWhenAvailable = True
                    .WakeToRun = cWake
                    .MultipleInstances = False
                    .RestartCount = 1

                    .Hidden = False
                End With
                With .Actions.Create(TASK_ACTION_EXEC)
                    .Path = TASK2_SERVICE_PATH ' full path
                    .WorkingDirectory = GetFilePath(TASK2_SERVICE_PATH)
                    .Arguments = TASK2_SERVICE_ARGS & " " & strId
                End With
                With .Principal
                    Dim ID
                    '        .UserID = "SYSTEM"
                    .RunLevel = 1
                End With
            End With

            ' this one creates the task
            With .GetFolder("\")
                .RegisterTaskDefinition(TASK2_SERVICE_NAME, TaskDef, TASK_CREATE_OR_UPDATE, , , TASK_LOGON_3)
            End With
        End With

        CreateWaitScheduledTask = True
        WriteStoreSetting(-1, IniSections_StoreSettings.iniSection_Program, "SchTsk_Wait", 1)
    End Function

    Private Function CreateWinCDSServiceScheduledTask() As Boolean
        '  If Not IsServer Then Exit Function
        'https://msdn.microsoft.com/en-us/library/windows/desktop/aa446862(v=vs.85).aspx
        ' also
        'http://www.vbforums.com/showthread.php?636846-VB6-Use-Vista-Task-Scheduler-2-0-API

        Dim strId As String
        Dim TaskDef As Object 'TaskScheduler.ITaskDefinition
        Const THREAD_PRIORITY_BELOW_NORMAL As Integer = 7

        strId = LOGFILEID ' Replace$(txtLogFile.Text, ".", "-") 'Create an Id value, no periods here!
        With CreateObject("Schedule.Service") 'New TaskScheduler.TaskScheduler
            .Connect

            ' This method call forces us to use late binding here, because .NewTask()
            ' takes an Unsupported Variant type argument "flags" (UInt?).  It has to
            ' be 0 anyway, so this works fine:
            TaskDef = .NewTask(0)

            With TaskDef
                With .RegistrationInfo
                    .Description = TASK_SERVICE_DESC
                    .Author = SoftwareVersion(True, False, True)
                    .URI = CompanyURL
                    .Source = ProgramShort
                    .Version = "1.0"
                End With
                With .Settings
                    .AllowDemandStart = True
                    .AllowHardTerminate = True

                    .StartWhenAvailable = False
                    .WakeToRun = False
                    .RunOnlyIfNetworkAvailable = True

                    .Enabled = True
                    .ExecutionTimeLimit = "PT45M"  ' 45 minute execution time limit

                    .StartWhenAvailable = False
                    .AllowHardTerminate = True

                    .Hidden = False
                End With
                With .Triggers.Create(TASK_TRIGGER_DAILY)
                    .StartBoundary = XmlTime(TASK_SERVICE_TIME)
                    '        .EndBoundary = XmlTime(DateAdd("n", CutoffMinutes, StartTime)) '30 minutes from dtpStart.
                    .DaysInterval = 1
                    .ExecutionTimeLimit = XmlDuration(Minutes:=45) 'Not sure why we have two of these ^^
                    .ID = "T" & strId
                    .Enabled = True
                End With
                With .Actions.Create(TASK_ACTION_EXEC)
                    .Path = TASK_SERVICE_PATH ' full path
                    .WorkingDirectory = GetFilePath(TASK_SERVICE_PATH)
                    .Arguments = TASK_SERVICE_ARGS & " " & strId
                    '        .ExecutionTimeLimit = "PT45M"  ' 45 minute execution time limit
                End With
                With .Principal
                    Dim ID
                    '        .DisplayName = "SYSTEM"
                    .UserId = "SYSTEM"
                    '        .Id = "NTAuthority\SYSTEM"
                    .RunLevel = 1
                End With
            End With

            ' this one creates the task
            With .GetFolder("\")
                .RegisterTaskDefinition(TASK_SERVICE_NAME, TaskDef, TASK_CREATE_OR_UPDATE, , , TASK_LOGON_3)
            End With
        End With

        CreateWinCDSServiceScheduledTask = True
    End Function

    Private ReadOnly Property TASK_SERVICE_PATH()
        Get
            TASK_SERVICE_PATH = WinCDSEXEFile(True, True, True)
        End Get
    End Property

    Private Function XmlDuration(Optional ByVal Years As Integer = 0, Optional ByVal Months As Integer = 0, Optional ByVal Days As Integer = 0, Optional ByVal Hours As Integer = 0, Optional ByVal Minutes As Integer = 0, Optional ByVal Seconds As Integer = 0) As String
        'In theory values like "P0YT20M" are valid, but Task Scheduler seems to reject them.  Use this strategy to suppres zeros.
        Dim strDate As String
        Dim strTime As String

        strDate = "P"
        If Years > 0 Then strDate = strDate & CStr(Years) & "Y"
        If Months > 0 Then strDate = strDate & CStr(Months) & "M"
        If Days > 0 Then strDate = strDate & CStr(Days) & "D"

        strTime = "T"
        If Hours > 0 Then strTime = strTime & CStr(Hours) & "H"
        If Minutes > 0 Then strTime = strTime & CStr(Minutes) & "M"
        If Seconds > 0 Then strTime = strTime & CStr(Seconds) & "S"

        If Len(strTime) = 1 Then strTime = ""
        XmlDuration = strDate & strTime
    End Function

    Private Function XmlTime(ByVal Timestamp As Date) As String
        'If CLng(DateValue(Timestamp)) = 0 Then Timestamp = Date & " " & Timestamp
        If IsNothing(DateValue(Timestamp)) Then Timestamp = Today & " " & Timestamp
        XmlTime = Format(Timestamp, "yyyy-mm-dd\THh:Nn:Ss")
    End Function

    Public Function VerifyScheduledTask(ByVal TaskName As String) As Boolean
        VerifyScheduledTask = InStr(ScheduledTaskListTaskNames(), TaskName)
    End Function

    Public Function ScheduledTaskListTaskNames() As String
        Dim objTaskService, objTaskFolder, colTasks, objTask

        objTaskService = CreateObject("Schedule.Service")
        objTaskService.Connect


        ' Get the task folder that contains the tasks.
        objTaskFolder = objTaskService.GetFolder("\")

        ' Get all of the tasks (Enumeration of 0 Shows all including Hidden.  1 will not show hidden)
        colTasks = objTaskFolder.GetTasks(0)

        If colTasks.Count = 0 Then
            ScheduledTaskListTaskNames = "No tasks are registered."
            Exit Function
        Else
            For Each objTask In colTasks
                ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa382079%28v=vs.85%29.aspx
                ScheduledTaskListTaskNames = ScheduledTaskListTaskNames & IIf(Len(ScheduledTaskListTaskNames) = 0, "", ",") & objTask.Name
            Next
        End If
    End Function

    Public Function ScheduledTaskDeleteTask(ByVal TaskName As String) As Boolean
        On Error Resume Next
        Dim objTaskService, objTaskFolder, colTasks, objTask
        If Not VerifyScheduledTask(TaskName) Then Exit Function

        objTaskService = CreateObject("Schedule.Service")
        objTaskService.Connect

        ' Get the task folder that contains the tasks.
        objTaskFolder = objTaskService.GetFolder("\")
        objTaskFolder.DeleteTask(TaskName, 0)
        ScheduledTaskDeleteTask = True
    End Function

End Module
