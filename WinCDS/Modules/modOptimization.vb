Imports Microsoft.VisualBasic.Compatibility.VB6
Module modOptimization
    Public Const DEFAULT_STOP_TIME As Integer = 10
    Public Const tspRS_MAX As Integer = 14
    Public Const MULTI_DAY_PENALTY As Integer = 1000000
    Private Const MySubSection As String = "Optimization"
    Public Const MISSED_WINDOW_PENALTY As Integer = 100000

    Public Enum tspRS
        tspRS_ID = 0
        tspRS_Name = 1
        tspRS_X = 2
        tspRS_Y = 3
        tspRS_WindowFrom = 4
        tspRS_WindowTo = 5
        tspRS_Distance = 6
        tspRS_Delay = 7
        tspRS_Arrive = 8
        tspRS_StopTime = 9
        tspRS_Depart = 10
        tspRS_Address = 11
        tspRS_City = 12
        tspRS_State = 13
        tspRS_Zip = 14
    End Enum

    Public Function GetOptimizationSetting(ByVal K As String) As String
        Dim A As String, B As Decimal
        A = GetCDSSetting(K, , MySubSection)
        Select Case K
            Case "StartTime"
                If Not IsDate(A) Then A = "7:00 AM"
                GetOptimizationSetting = Format(TimeValue(A), "h:mm ampm")
            Case "TimePerStop"
                GetOptimizationSetting = Val(A)
                If Val(GetOptimizationSetting) < 0 Then GetOptimizationSetting = CLng(10)
            Case "CostPerHour", "CostPerMile"
                GetOptimizationSetting = GetPrice(A)
                If Val(GetOptimizationSetting) <= 0 Then GetOptimizationSetting = IIf(K = "CostPerHour", 11, 0.45)
            Case "Trucks"
                GetOptimizationSetting = Val(A)
                If Val(GetOptimizationSetting) <= 0 Then GetOptimizationSetting = 1
            Case Else : Err.Raise(-1, , "GetOptimizationSetting -- Invalid Optimzation Setting: " & K)
        End Select
    End Function

End Module
