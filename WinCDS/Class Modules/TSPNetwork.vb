Imports MapPoint
Imports MapPoint.GeoCountry
Imports MapPoint.GeoFindResultsQuality

Public Class TSPNetwork
    Private mStartTime As Date
    Public CostPerMile As Decimal
    Public CostPerHour As Decimal
    Public DefaultStopTime As Long
    Public mMap As Map
    Private m_BestSolution() As Object
    Private m_Distances(100, 100) As Single '@NO-LINT
    Private m_Times(100, 100) As Single '@NO-LINT

    Private m_BestCost As Single
    Private m_BestSolutionDelays() As Object
    Private m_BestSolutionEndTime As Date

    Private m_TestSolution() As Object
    Private m_TestSolutionDelays() As Object

    Private NewDay As Boolean, DupDepot As Boolean
    Private D_Best() As Object, C_Best As Single, c_BestEndTime As Date    ' for recursive delay calculation
    Public Nodes As Collection
    Private Const MAX_SINGLE As Single = 3.402823E+38
    Public N As Long
    Private mTrucks As Long
    Private mTrialCount As Long
    Private mRunning As Boolean
    Private mDoCancel As Boolean
    Private mMaxFails As Long
    Private Const MAX_LONG As Long = 2147483647

    Public Sub Setup(Optional ByVal nStartTime As Date = #9:00:00 AM#, Optional ByVal nCostPerMile As Double = 0.45, Optional ByVal nCostPerHour As Double = 10.0#, Optional ByRef nDefaultStopTime As Long = DEFAULT_STOP_TIME, Optional ByRef nMap As Object = Nothing)
        StartTime = nStartTime
        CostPerMile = nCostPerMile
        CostPerHour = nCostPerHour
        DefaultStopTime = nDefaultStopTime
        mMap = nMap
    End Sub

    Public Function GetResultSet() As Object
        Dim Z As Long, I As Long, J As Long, N As TSPNode
        Dim IX As Long, Prev As Long, Dly As Long, Trn As Long, Dst As Long
        Dim T As Date
        Dim R As Object
        ReDim R(Count + Trucks, tspRS_MAX)
        If Count < 2 Then GetResultSet = R : Exit Function
        On Error GoTo NONENONE


        N = Node(0)
        R(0, tspRS.tspRS_ID) = 0
        R(0, tspRS.tspRS_Name) = N.Name
        R(0, tspRS.tspRS_X) = N.X
        R(0, tspRS.tspRS_Y) = N.Y
        R(0, tspRS.tspRS_WindowFrom) = ""
        R(0, tspRS.tspRS_WindowTo) = ""
        R(0, tspRS.tspRS_Distance) = 0
        R(0, tspRS.tspRS_Delay) = 0
        R(0, tspRS.tspRS_Arrive) = Format(StartTime, "h:mmampm")
        R(0, tspRS.tspRS_StopTime) = 0
        R(0, tspRS.tspRS_Depart) = Format(StartTime, "h:mmampm")
        R(0, tspRS.tspRS_Address) = N.Address
        R(0, tspRS.tspRS_City) = N.City
        R(J, tspRS.tspRS_State) = N.State
        R(0, tspRS.tspRS_Zip) = N.Zip

        Z = GetStartingIndex(m_BestSolution)
        J = 1
        For I = 1 To Count
            IX = (Z + I) Mod Count
            Prev = (Z + I - 1) Mod Count
            Dst = m_Distances(m_BestSolution(Prev), m_BestSolution(IX))
            Trn = m_Times(m_BestSolution(Prev), m_BestSolution(IX))
            Dly = m_BestSolutionDelays(IX)
            N = Node(m_BestSolution(IX))

            T = AddMinutes(Trn + Dly, R(J - 1, 10))

            If N.IsDepot And I <> Count Then    ' new truck, finish and start over
                R(J, tspRS.tspRS_ID) = 0
                R(J, tspRS.tspRS_Name) = N.Name
                R(J, tspRS.tspRS_X) = N.X
                R(J, tspRS.tspRS_Y) = N.Y
                R(J, tspRS.tspRS_WindowFrom) = ""
                R(J, tspRS.tspRS_WindowTo) = ""
                R(J, tspRS.tspRS_Distance) = Dst
                R(J, tspRS.tspRS_Delay) = Dly
                R(J, tspRS.tspRS_Arrive) = Format(T, "h:mmampm")
                R(J, tspRS.tspRS_StopTime) = 0
                R(J, tspRS.tspRS_Depart) = Format(T, "h:mmampm")
                R(J, tspRS.tspRS_Address) = N.Address
                R(J, tspRS.tspRS_Address) = N.City
                R(J, tspRS.tspRS_State) = N.State
                R(J, tspRS.tspRS_Zip) = N.Zip

                J = J + 1
                T = StartTime
                R(J, tspRS.tspRS_ID) = 0
                R(J, tspRS.tspRS_Name) = N.Name
                R(J, tspRS.tspRS_X) = N.X
                R(J, tspRS.tspRS_Y) = N.Y
                R(J, tspRS.tspRS_WindowFrom) = ""
                R(J, tspRS.tspRS_WindowTo) = ""
                R(J, tspRS.tspRS_Distance) = 0
                R(J, tspRS.tspRS_Delay) = 0
                R(J, tspRS.tspRS_Arrive) = Format(T, "h:mmampm")
                R(J, tspRS.tspRS_StopTime) = 0
                R(J, tspRS.tspRS_Depart) = Format(T, "h:mmampm")
                R(J, tspRS.tspRS_Address) = N.Address
                R(J, tspRS.tspRS_City) = N.City
                R(J, tspRS.tspRS_State) = N.State
                R(J, tspRS.tspRS_Zip) = N.Zip
            Else
                R(J, tspRS.tspRS_ID) = m_BestSolution(IX)
                R(J, tspRS.tspRS_Name) = N.Name
                R(J, tspRS.tspRS_X) = N.X
                R(J, tspRS.tspRS_Y) = N.Y
                R(J, tspRS.tspRS_WindowFrom) = IIf(I = Count, "", IIf(N.WindowFrom = #12:00:00 AM#, "", Format(N.WindowFrom, "h:mmampm")))
                R(J, tspRS.tspRS_WindowTo) = IIf(I = Count, "", IIf(N.WindowTo = #11:59:59 PM#, "", Format(N.WindowTo, "h:mm ampm")))
                R(J, tspRS.tspRS_Distance) = Dst
                R(J, tspRS.tspRS_Delay) = Dly
                R(J, tspRS.tspRS_Arrive) = Format(T, "h:mmampm")
                R(J, tspRS.tspRS_StopTime) = N.StopTime
                R(J, tspRS.tspRS_Depart) = Format(AddMinutes(N.StopTime, T), "h:mmampm")
                R(J, tspRS.tspRS_Address) = N.Address
                R(J, tspRS.tspRS_City) = N.City
                R(J, tspRS.tspRS_State) = N.State
                R(J, tspRS.tspRS_Zip) = N.Zip
            End If
            J = J + 1
        Next
        GetResultSet = R
NONENONE:
    End Function

    Public ReadOnly Property Count() As Long
        Get
            Count = Nodes.Count
        End Get
    End Property

    Public Property StartTime() As Date
        Get
            StartTime = mStartTime
        End Get
        Set(value As Date)
            mStartTime = TimeValue(value)
        End Set
    End Property

    ' Add a new node at this point.
    Public Sub AddLocation(Optional ByVal Name As String="", Optional ByVal Address As String="", Optional ByRef City As String="", Optional ByRef State As String="", Optional ByRef Zip As String="", Optional ByVal StopTime As Long = -1, Optional ByVal WindowFrom As Date = #12:00:00 AM#, Optional ByVal WindowTo As Date = #11:59:59 PM#)
        Dim N As New TSPNode
        If Name = "" Then Name = "Node " & Count + 1
        If DateAfter(WindowFrom, WindowTo, , "n") Then
            Err.Raise(-1, , "The delivery window is invalid.  From must be before to!")
        End If

        If Not TestAddressWithMapPoint(Address, City, State, Zip) Then
            MessageBox.Show("The following address could not be added to the route:" & vbCrLf2 & Name & vbCrLf & Address & vbCrLf & City & ", " & State & " " & Zip, "Invalid Address", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If StopTime < 0 Then StopTime = DefaultStopTime

        N.IsDepot = (Count = 0) Or DupDepot
        If N.IsDepot Then
            N.Setup(Name, 0, 0, 0, , , Address, City, State, Zip)
        Else
            N.Setup(Name, 0, 0, StopTime, WindowFrom, WindowTo, Address, City, State, Zip)
        End If
        Nodes.Add(N)
    End Sub

    ' Find a solution by using random improvement and two-opt.
    Public Function Solve(Optional ByVal DoContinue As Boolean = False) As Single
        Dim I As Long, J As Long
        Dim Times As Long, Num_fails As Long
        ' Make sure we're ready.
        If CheckReady(DoContinue) Then Solve = MAX_SINGLE : Exit Function

        Times = 10
        ProgressForm(0, Times, "Optimizing...", vbCancel)
        For I = 1 To Times
            ProgressForm(I)
            N = I
            ReportSolutionCostProgress

            Num_fails = 0
            Do While Num_fails < MaxFails
                If Not CheckCancel Then Exit Function
                If TryRandomImprovement() Then Num_fails = 0 Else Num_fails = Num_fails + 1
                'DoEvents
                My.Application.DoEvents()
            Loop

            Num_fails = 0
            Do While Num_fails < MaxFails   ' Use Two-Opt.
                If Not CheckCancel Then Exit Function
                If TryTwoOpt() Then Num_fails = 0 Else Num_fails = Num_fails + 1
                My.Application.DoEvents()
            Loop

            Num_fails = 0
            Do While Num_fails < MaxFails   ' Use four opt.
                If Not CheckCancel Then Exit Function
                If TryFourOpt() Then Num_fails = 0 Else Num_fails = Num_fails + 1
                My.Application.DoEvents()
            Loop

            Num_fails = 0
            Do While Num_fails < MaxFails   ' Use Std Dev Clustering Technique
                If Not CheckCancel Then Exit Function
                If TryStdDev() Then Num_fails = 0 Else Num_fails = Num_fails + 1
                My.Application.DoEvents()
            Loop

            '    num_fails = 0
            '    Do While num_fails < 3   ' Use Line Straightener
            '      If Not CheckCancel Then Exit Function
            '      If TryLine() Then num_fails = 0 Else num_fails = num_fails + 1
            '      DoEvents
            '    Loop
        Next
        ProgressForm()

        Solve = m_BestCost    ' Return the best cost.
        CheckDone
    End Function

    Public Property Trucks() As Long
        Get
            Trucks = mTrucks
        End Get
        Set(value As Long)
            If value < 1 Then value = 1
            If value > 10 Then value = 10
            mTrucks = value
        End Set
    End Property

    Public ReadOnly Property Node(ByVal IX As Long) As TSPNode
        Get
            On Error Resume Next
            Node = Nodes(IX + 1)
        End Get
    End Property

    Private Function GetStartingIndex(ByRef Solution() As Object) As Long
        GetStartingIndex = GetCurrentIndex(Solution, 0)
NONENONE:
    End Function

    Private Function AddMinutes(ByVal N As Long, ByVal T As Date) As Date
        Dim L As Long
        AddMinutes = TimeValue(DateAdd("n", N, T))
        L = DateDiff("n", T, AddMinutes)
        If L <> N Then
            NewDay = True
        Else
            NewDay = False
        End If
        '  NewDay = (L <> N)
    End Function

    Private Function TestAddressWithMapPoint(ByRef A As String, ByRef C As String, ByRef S As String, ByRef Z As String) As Boolean
        Dim FR As FindResults
        On Error GoTo NoGood
        If mMap Is Nothing Then TestAddressWithMapPoint = True : Exit Function
        FR = mMap.FindAddressResults(A, C, , , Z, geoCountryUnitedStates)
        TestAddressWithMapPoint = FR.ResultsQuality = geoFirstResultGood Or FR.ResultsQuality = geoAmbiguousResults Or FR.ResultsQuality = geoNoGoodResult
NoGood:
    End Function

    Private Function CheckReady(Optional ByVal DoContinue As Boolean = False) As Boolean
        Dim I As Long
        If (Count < 4) Then
            MessageBox.Show("The network must contain at least 3 Deliveries.", "No Nodes", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            CheckReady = True
            Exit Function
        End If

        If TrialCount <= 0 Then DoContinue = False
        Running = True
        mDoCancel = False

        If Not DoContinue Then
            PrepareTrucks

            ' Initialize variables.
            m_BestCost = MAX_SINGLE
            ReDim m_BestSolution(Count - 1)
            ReDim m_BestSolutionDelays(Count - 1)
            ReDim m_TestSolution(Count - 1)
            ReDim m_TestSolutionDelays(Count - 1)


            ' Fill in a (non-random) test solution.
            For I = 0 To Count - 1
                m_TestSolution(I) = I
            Next

            TrialCount = -1
            CalculateDistances ' Calculate distances between nodes.
            If Not Running Then CheckReady = True : Exit Function ' might have cancelled in calculating distances
            TrialCount = 0

            TryRandomSolution     'puts test solution into best
        End If
    End Function

    Private Sub ReportSolutionCostProgress()
        If Not IsFormLoaded("frmOptimize") Then frmOptimize.Show()
        frmOptimize.Network = Me
        If IsFormLoaded("frmOptimize") Then DrawNetwork(frmOptimize.picNetwork)
    End Sub

    Public Property MaxFails() As Long
        Get
            MaxFails = IIf(mMaxFails <= 10, 10, mMaxFails)
        End Get
        Set(value As Long)
            mMaxFails = value
        End Set
    End Property

    ' return true if we keep going, false if we stop
    Private Function CheckCancel() As Boolean
        If Not DoCancel Then CheckCancel = True : Exit Function
        Running = False
    End Function

    ' Try swapping two nodes to see if there is an improvement.
    ' Return True if there is an improvement.
    Private Function TryRandomImprovement() As Boolean
        Dim Node1 As Long, Node2 As Long
        Node1 = RandomIndexRange(Count - 1, 0)    ' Find two distinct nodes.
        Node2 = (Node1 + RandomIndexRange(1, Count)) Mod Count
        If Node2 >= Node1 Then Node2 = (Node2 + 1) Mod Count
        TryRandomImprovement = TrySwapImprovement(Node1, Node2)
    End Function

    ' Use 2-opt to try to improve the solution.
    ' Return True if we have made an improvement.
    Private Function TryTwoOpt() As Boolean
        ' Pick two different random links to swap.
        Dim FR1 As Long, FR2 As Long, To1 As Long, To2 As Long
        FR1 = RandomIndexRange(0, Count)
        To1 = (FR1 + 1) Mod Count
        FR2 = (To1 + RandomIndexRange(1, Count - 2)) Mod Count
        To2 = (FR2 + 1) Mod Count

        ' Make sure the links are separate.
        'If fr1 = fr2 Or fr1 = To2 Or to1 = fr2 Or to1 = To2 Then Stop

        ' Replace links fr1 -> to1 and fr2 -> to2
        ' with links fr1 -> fr2 and to1 -> to2.
        m_TestSolution = CopyArr(m_BestSolution)
        ' Note that swapping the links is equivalent to reversing
        ' the order of the nodes between to1 and fr2 or those
        ' between to2 and fr1, inclusive.

        ' Note also that reversing the nodes between to1 and fr2
        ' is almost the same as reversing the nodes between fr2 and to1.
        ' It it reverses the route but has the same total cost.
        If To1 < FR2 Then
            m_TestSolution = RevArrSub(m_BestSolution, To1, FR2 - To1 + 1)
        Else
            m_TestSolution = RevArrSub(m_BestSolution, To2, FR1 - To2 + 1)
        End If
        TryTwoOpt = CheckForImprovement
    End Function

    ' Try swapping two nodes to see if there is an improvement.
    ' Return True if there is an improvement.
    Private Function TryFourOpt() As Boolean
        Dim Before1 As Long, Before2 As Long, Node1 As Long, Node2 As Long, After1 As Long, After2 As Long
        Dim Temp As Long
        Before1 = RandomNodeIndex 'm_Random.Next(0, Nodes.Count - 1)
        Before1 = RandomIndexRange(0, Count - 1) 'm_Random.Next(0, Nodes.Count - 1)
        Before2 = (Before1 + RandomIndexRange(2, Count - 1)) Mod Count
        Node1 = (Before1 + 1) Mod Count
        Node2 = (Before2 + 1) Mod Count
        After1 = (Node1 + 1) Mod Count
        After2 = (Node2 + 1) Mod Count

        ' For our test solution
        m_TestSolution = CopyArr(m_BestSolution)

        ' Make the swap.
        Temp = m_TestSolution(Node1)
        m_TestSolution(Node1) = m_TestSolution(Node2)
        m_TestSolution(Node2) = Temp

        TryFourOpt = CheckForImprovement
    End Function

    Private Function TryStdDev() As Boolean
        Dim N As Long, X() As Object, I As Long, D As Double, V() As Object, vC As Long, uX As Long, Uy As Long
        Dim L() As Object
        N = RandomNodeIndex

        ReDim X(0 To Count - 1)
        For I = 0 To Count - 1
            X(I) = m_Distances(N, I)
        Next
        D = ArrayStdDev(X)

        ReDim V(0 To Count)
        vC = 0

        For I = 0 To Count - 1
            If m_Distances(N, I) <= D Then
                V(vC) = I
                vC = vC + 1
            End If
        Next
        ReDim Preserve V(0 To vC - 1)

        If vC <= 1 Then Exit Function

        ReDim L(vC)
        On Error GoTo DoneIt
        For I = 0 To vC
            L(I) = GetCurrentIndex(m_BestSolution, X(I))
        Next
DoneIt:

        uX = RandomIndexRange(0, vC)
        Uy = RandomIndexRange(0, vC)
        If uX = Uy Then Uy = (Uy + 1) Mod vC


        TryStdDev = TrySwapImprovement(uX, Uy)
    End Function

    Private Function CheckDone() As Boolean
        Running = False
    End Function

    Private Function GetCurrentIndex(ByRef Solution() As Object, ByVal NodeID As Long) As Long
        Dim I As Long, Z As Long
        On Error GoTo NONENONE
        Z = 0
        For I = LBound(Solution) To UBound(Solution)
            If Solution(I) = NodeID Then Z = I : Exit For
        Next
        GetCurrentIndex = Z
NONENONE:
    End Function

    Public Property TrialCount() As Long
        Get
            TrialCount = mTrialCount
        End Get
        Set(value As Long)
            mTrialCount = value
        End Set
    End Property

    Public Property Running() As Boolean
        Get
            Running = mRunning
        End Get
        Set(value As Boolean)
            mRunning = value
        End Set
    End Property

    Private Sub PrepareTrucks()
        Dim I As Long
        For I = 1 To Count - 1
            If Node(I).IsDepot Then RemoveNode(I)
        Next

        If Trucks > 1 Then
            For I = 2 To Trucks
                DupDepot = True
                AddNode(Node(0).Name, Node(0).X, Node(0).Y)
                DupDepot = False
            Next
        End If
    End Sub

    ' Calculate the distances between every pair of nodes.
    Private Sub CalculateDistances()
        Dim I As Long, J As Long, dX As Single, dY As Single, Dst As Long, Tme As Long
        Dim X As Long, Y As Long
        '  ReDim m_Distances(Nodes.Count - 1, Nodes.Count - 1)

        ProgressForm(0, Count * Count, "Preparing...")
        For I = 0 To Count - 1
            If Not mMap Is Nothing Then
                CalculatePositionWithMapPoint(I, X, Y)
                Node(I).X = X
                Node(I).Y = Y
            End If
            For J = I To Count - 1
                If Not CheckCancel() Then
                    ProgressForm()
                    Exit Sub
                End If
                ProgressForm(I * Count + J)
                If I = J Then
                    m_Distances(I, J) = 0
                    m_Times(I, J) = 0
                Else
                    If mMap Is Nothing Then
                        dX = Node(I).X - Node(J).X
                        dY = Node(I).Y - Node(J).Y
                        m_Distances(I, J) = CLng(Math.Sqrt(dX * dX + dY * dY) / 50)
                        m_Times(I, J) = CLng((Math.Abs(dX) + Math.Abs(dY)) / 120)
                    Else
                        CalculateLegWithMapPoint(I, J, Dst, Tme)
                        m_Distances(I, J) = Dst
                        m_Times(I, J) = Tme
                    End If

                    m_Distances(J, I) = m_Distances(I, J)
                    m_Times(J, I) = m_Times(I, J)
                End If
            Next
        Next
        ProgressForm()
    End Sub

    ' Try a random solution.
    Private Sub TryRandomSolution()
        Dim I As Long, J As Long, Temp As Long
        ' Pick random nodes for all positions.
        For I = 0 To Count - 2
            ' Pick a node for position i.
            '    J = RandomNodeIndex 'm_Random.Next(I, Nodes.Count)
            J = RandomIndexRange(I, Count)

            ' Swap nodes i and j.
            Temp = m_TestSolution(I)
            m_TestSolution(I) = m_TestSolution(J)
            m_TestSolution(J) = Temp
        Next

        ' See if this solution is an improvement.
        CheckForImprovement
    End Sub

    Public Sub DrawNetwork(ByVal pic As PictureBox)
        Dim N As TSPNode, I As Long, L As Long, R As Long
        Dim XOff As Long, YOff As Long, XScl As Double, YScl As Double
        Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
        R = (MaxY - MinY) * 1.1
        If R <> 0 Then
            YScl = pic.ScaleHeight / R
            YOff = MinY - 0.05 * R
            R = (MaxX - MinX) * 1.1
            XScl = pic.ScaleWidth / R
            XOff = MinX - 0.05 * R
        End If

        If YScl > XScl Then YScl = XScl
        If XScl > YScl Then XScl = YScl

        pic.Cls

        pic.ForeColor = vbBlack
        pic.DrawWidth = 2
        pic.ForeColor = ChangeNetworkLineColor(pic.ForeColor)
        For I = 0 To Count - 2
            If Node(m_BestSolution(I)).IsDepot Then pic.ForeColor = ChangeNetworkLineColor(pic.ForeColor)
            X1 = XScl * (Node(m_BestSolution(I)).X - XOff)
            Y1 = YScl * (Node(m_BestSolution(I)).Y - YOff)
            X2 = XScl * (Node(m_BestSolution(I + 1)).X - XOff)
            Y2 = YScl * (Node(m_BestSolution(I + 1)).Y - YOff)
            pic.Line(X1, Y1)-(X2, Y2)
  Next
        If Node(m_BestSolution(Count - 1)).IsDepot Then pic.ForeColor = ChangeNetworkLineColor(pic.ForeColor)
        X1 = XScl * (Node(m_BestSolution(Count - 1)).X - XOff)
        Y1 = YScl * (Node(m_BestSolution(Count - 1)).Y - YOff)
        X2 = XScl * (Node(m_BestSolution(0)).X - XOff)
        Y2 = YScl * (Node(m_BestSolution(0)).Y - YOff)
        pic.Line(X1, Y1)-(X2, Y2)

  For Each N In Nodes ' Draw the nodes.
            N.DrawNode pic, XOff, YOff, XScl, YScl
  Next

        pic.ForeColor = vbBlack
        pic.FontName = "Arial"
        pic.FontSize = 8

        X1 = pic.ScaleWidth - 1200
        pic.CurrentX = X1
        pic.CurrentY = 100
        pic.Print "Nodes = " & Count
  pic.CurrentX = X1
        pic.Print "Cost = " & CurrencyFormat(m_BestCost)
  pic.CurrentX = X1
        pic.Print "Trials=" & TrialCount
  pic.CurrentX = X1
        pic.Print "N   =   " & N
End Sub

    Private ReadOnly Property DoCancel() As Boolean
        Get
            DoCancel = mDoCancel
        End Get
    End Property

    Private Function RandomIndexRange(ByVal X As Long, ByVal Y As Long) As Long
        Dim Ub As Long, LB As Long
        Ub = Y - 1
        LB = X
        RandomIndexRange = Int((Ub - LB + 1) * Rnd() + LB)
    End Function

    Private Function TrySwapImprovement(ByRef N1 As Long, ByRef N2 As Long) As Boolean
        Dim Temp As Long
        m_TestSolution = CopyArr(m_BestSolution)

        Temp = m_TestSolution(N1)
        m_TestSolution(N1) = m_TestSolution(N2)
        m_TestSolution(N2) = Temp

        TrySwapImprovement = CheckForImprovement
    End Function

    Private Function CheckForImprovement() As Boolean
        Dim Test_Cost As Single
        TrialCount = TrialCount + 1
        Test_Cost = CalculateSolutionCost(m_TestSolution)

        If Test_Cost < m_BestCost Then      ' Save this solution.
            m_BestCost = Test_Cost
            m_BestSolution = CopyArr(m_TestSolution)
            m_BestSolutionDelays = CopyArr(m_TestSolutionDelays)
            m_BestSolutionEndTime = c_BestEndTime
            ReportSolutionCostProgress()
            CheckForImprovement = True
        End If
    End Function

    Private Function RandomNodeIndex() As Long
        Dim Ub As Long, LB As Long
        Ub = Count - 1
        LB = 0
        RandomNodeIndex = Int((Ub - LB + 1) * Rnd() + LB)
    End Function

    Private Function ArrayStdDev(ByRef Arr As Variant, Optional ByRef SampleStdDev As Boolean = True, Optional ByRef IgnoreEmpty As Boolean = True) As Double
        Dim Sum As Double
        Dim sumSquare As Double
        Dim Value As Double
        Dim Count As Long
        Dim Index As Long

        ' evaluate sum of values
        ' if arr isn't an array, the following statement raises an error
        For Index = LBound(Arr) To UBound(Arr)
            Value = Arr(Index)
            ' skip over non-numeric values
            If IsNumeric(Value) Then
                ' skip over empty values, if requested
                If Not (IgnoreEmpty And IsEmpty(Value)) Then
                    ' add to the running total
                    Count = Count + 1
                    Sum = Sum + Value
                    sumSquare = sumSquare + Value * Value
                End If
            End If
        Next

        ' evaluate the result
        ' use (Count-1) if evaluating the standard deviation of a sample
        If SampleStdDev Then
            ArrayStdDev = Sqr((sumSquare - (Sum * Sum / Count)) / (Count - 1))
        Else
            ArrayStdDev = Sqr((sumSquare - (Sum * Sum / Count)) / Count)
        End If
    End Function

    Public Sub RemoveNode(ByVal N As Long)
        Nodes.Remove N + 1
End Sub

    Public Sub AddNode(Optional ByVal Name As String, Optional ByVal X As Long, Optional ByVal Y As Long, Optional ByVal StopTime As Long = -1, Optional ByVal WindowFrom As Date = #12:00:00 AM#, Optional ByVal WindowTo As Date = #11:59:59 PM#)
        Dim N As New TSPNode
        If Name = "" Then Name = "Node " & Count + 1
        If DateAfter(WindowFrom, WindowTo, , "n") Then
            Err.Raise -1, , "The delivery window is invalid.  From must be before to!"
  End If

        If StopTime < 0 Then StopTime = DefaultStopTime

        N.IsDepot = (Count = 0) Or DupDepot
        If N.IsDepot Then
            N.Setup Name, X, Y, 0
  Else
            N.Setup Name, X, Y, StopTime, WindowFrom, WindowTo
  End If
        Nodes.Add N
End Sub

    Private Sub CalculatePositionWithMapPoint(ByVal IX As Long, ByRef X As Long, ByRef Y As Long)
        Dim Nx As TSPNode, FR As FindResults
  Set Nx = Node(IX)
  mMap.ActiveRoute.Clear()
          Set FR = mMap.FindAddressResults(Nx.Address, Nx.City, , , Nx.Zip, geoCountryUnitedStates)
  If FR.Count >= 1 Then
            X = (FR(1).Longitude * 10000)
            Y = -(FR(1).Latitude * 10000)
        End If
  Set FR = Nothing
End Sub

    Private Sub CalculateLegWithMapPoint(ByVal IX As Long, ByVal IY As Long, ByRef Dist As Long, ByRef Time As Long)
        Dim Nx As TSPNode, nY As TSPNode
        Dim FR As FindResults, Pin As Pushpin
  Set Nx = Node(IX)
  Set nY = Node(IY)
  
  mMap.ActiveRoute.Clear()
          Set FR = mMap.FindAddressResults(Nx.Address, Nx.City, , , Nx.Zip, geoCountryUnitedStates)
  If FR.Count >= 1 Then
    Set Pin = mMap.AddPushpin(FR(1), Nx.Name)
    mMap.ActiveRoute.Waypoints.Add Pin, Pin.Name
  End If
  Set FR = mMap.FindAddressResults(nY.Address, nY.City, , , nY.Zip, geoCountryUnitedStates)
  If FR.Count >= 1 Then
    Set Pin = mMap.AddPushpin(FR(1), nY.Name)
    mMap.ActiveRoute.Waypoints.Add Pin, Pin.Name
  End If
        mMap.ActiveRoute.Calculate()
        Dist = mMap.ActiveRoute.Distance
        Time = mMap.ActiveRoute.DrivingTime / geoOneHour * 60
    End Sub

    Public ReadOnly Property MaxX() As Long
        Get
            Dim I As Long
            MaxX = -MAX_LONG
            For I = 0 To Count - 1
                If Node(I).X > MaxX Then MaxX = Node(I).X
            Next
        End Get
    End Property

    Public ReadOnly Property MaxY() As Long
        Get
            Dim I As Long
            MaxY = -MAX_LONG
            For I = 0 To Count - 1
                If Node(I).Y > MaxY Then MaxY = Node(I).Y
            Next
        End Get
    End Property

    Public ReadOnly Property MinX() As Long
        Get
            Dim I As Long
            MinX = MAX_LONG
            For I = 0 To Count - 1
                If Node(I).X < MinX Then MinX = Node(I).X
            Next
        End Get
    End Property

    Public ReadOnly Property MinY() As Long
        Get
            Dim I As Long
            MinY = MAX_LONG
            For I = 0 To Count - 1
                If Node(I).Y < MinY Then MinY = Node(I).Y
            Next
        End Get
    End Property

End Class
