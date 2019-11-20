Imports MapPoint
Imports MapPoint.GeoCountry
Imports MapPoint.GeoTimeConstants
Imports Microsoft.VisualBasic.Compatibility.VB6

Public Class frmDeliveryMap
    Private Printed As Boolean
    Public HasOptimizer As Boolean
    Dim Network As TSPNetwork
    Dim WithEvents Opt As OptimRoute.RouteOptimizer
    Dim IOpt As OptimRoute.IRouteOptimizer, IOptCont As OptimRoute.IRouteContainer
    Dim HiddenMapControl As Application, HiddenMap As Map, XYMap As Map

    Public Function CreateRoute(ByVal StoreNum As Integer, ByVal DeliveryDate As String) As Route
        ' StoreNum gets us the store address, for the start and end points.
        ' Customers gets us mailing/delivery addresses.  We auto-generate the route and the user change it.

        Text = "Delivery Route for " & DeliveryDate

        Dim WP As Waypoint, wpStore As Waypoint
        Dim FirstStore As Integer, LastStore As Integer, GMDB As String, SQL As String
        Dim I As Integer, K As String
        Dim StoreCity As String, StoreState As String, StoreZip As String
        Dim Sales As CGrossMargin, LastSale As String
        '  Dim Cust As clsMailRec, Shipping As MailNew2
        Dim Service As clsServiceOrder, ListCancel As Boolean

        Dim OrderList() As Object, OrderCount As Integer, UseOrderList As Boolean


        ' Start point
        StoreCity = StoreSettings.City
        CitySTZip(StoreCity, StoreState, StoreZip)

        ' If city/state/zip isn't available, alert the user?
        wpStore = AddWaypoint(mapDelivery.ActiveMap, StoreSettings.Address, StoreCity, , StoreState, StoreZip, geoCountryUnitedStates, StoreSettings.Name)
        '    If wpStore Is Nothing Then Exit Function

        'StoreNum = 0
        If StoreNum = 0 Then
            FirstStore = 1 'StoreNum
            LastStore = LicensedNoOfStores() 'StoreNum
            UseOrderList = True
        Else
            FirstStore = StoreNum
            LastStore = StoreNum
            UseOrderList = False
        End If

        LoadAllStops(DeliveryDate, FirstStore, LastStore)
        SelectAllStops()

        If UseOrderList Then
            'cmdSplit.Value = True
            cmdSplit.PerformClick()
            Show() 'vbModal
            Exit Function
        End If

        RouteThisTruck()
        Show()
    End Function

    Private Function AddWaypoint(ByRef M As Map, Optional ByVal Street As String = "", Optional ByVal City As String = "", Optional ByVal OtherCity As String = "", Optional ByVal State As String = "", Optional ByVal Zip As String = "", Optional ByVal Country As Integer = 0, Optional ByVal PointName As String = "", Optional ByVal DoSelect As Boolean = False) As Waypoint
        Dim FR As FindResults, Pin As Pushpin, I As Integer, L As Location
        On Error GoTo BadWpt
        FR = M.FindAddressResults(Street, City, OtherCity, State, Zip, Country)

        On Error GoTo 0
        If FR.Count > 0 Then
            If FR.Count > 1 Then
                If FR.ResultsQuality > 1 Then
                    ' Choose an option, or cancel?
                    Dim AddrArray() As Object
                    'ReDim AddrArray(1 To FR.Count)
                    ReDim AddrArray(0 To FR.Count - 1)
                    For I = 1 To FR.Count
                        'AddrArray(I) = FR.Item(I).Name
                        AddrArray(I - 1) = FR.Item(I).Name
                    Next
                    I = SelectOptionArray("Clarify " & PointName & " Address", 0, AddrArray)
                    If I < 0 Then I = 0
                End If
            Else
                I = 1
            End If
            If I > 0 Then
                On Error Resume Next
                L = FR(I)
                Pin = M.AddPushpin(L, PointName)
                If DoSelect Then Pin.Select()
                AddWaypoint = M.ActiveRoute.Waypoints.Add(Pin)
            End If
        End If
        Exit Function
BadWpt:
        MessageBox.Show("Couldn't route the following address:" & vbCrLf2 & Street & vbCrLf & City & ", " & State & " " & Zip & vbCrLf2 & "NOTE: You cannot route to a PO box.", "Bad Address", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Function

    Private Sub LoadAllStops(ByVal DeliveryDate As String, ByVal FirstStore As Integer, ByVal LastStore As Integer)
        Dim WP As Waypoint
        Dim GMDB As String, SQL As String
        Dim I As Integer, K As String
        Dim Sales As CGrossMargin, LastSale As String
        '  Dim Cust As clsMailRec, Shipping As MailNew2
        Dim Service As clsServiceOrder
        'Dim Li As ListItem
        Dim Li As ListViewItem
        Dim Cubes As Double

        'lvwAllStops.ListItems.Clear ' BFH20080608 - Moved before the for loop so previous stores still show
        lvwAllStops.Items.Clear()
        For I = FirstStore To LastStore
            GMDB = GetDatabaseAtLocation(I)

            Sales = New CGrossMargin
            Sales.DataAccess.DataBase = GMDB
            SQL = ""
            SQL = SQL & "SELECT * FROM GrossMargin "
            SQL = SQL & " WHERE [DelDate]=#" & DeliveryDate & "#"
            SQL = SQL & " AND Trim(Style) Not In (" & NonItemStyleString() & ")"
            SQL = SQL & " AND [PorD]<>'P'"
            SQL = SQL & " AND (Left([Status],3)<>'DEL')" 'BFH20080608 - 'DEL%' to Left(..,3)<>'DEL'
            SQL = SQL & " ORDER BY [SaleNo], [MarginLine]"
            Sales.DataAccess.Records_OpenSQL(SQL)

            Do While Sales.DataAccess.Records_Available
                K = "LOC " & I & " - " & Sales.SaleNo

                If LastSale <> Sales.SaleNo Then
                    LastSale = Sales.SaleNo
                    Cubes = GetCubesOnSale(Sales.SaleNo, DeliveryDate)
                    'Li = lvwAllStops.ListItems.Add(, K, "Sale " & Sales.SaleNo & ", " & GetMailLastNameByIndex(Sales.Index, I, True) & " (L" & I & ", " & Format(Cubes, "0.00") & ")", , "stop")
                    Li = lvwAllStops.Items.Add(K, "Sale " & Sales.SaleNo & ", " & GetMailLastNameByIndex(Sales.Index, I, True) & " (L" & I & ", " & Support.Format(Cubes, "0.00"), 2)
                    SetStopInfoByLI(Li, I, "Sale", Sales.SaleNo, GetMailLastNameByIndex(Sales.Index, I, True), Sales.Index, Sales.StopStart, Sales.StopEnd, Cubes)
                End If
            Loop
            DisposeDA(Sales)

            Service = New clsServiceOrder
            Service.DataAccess.DataBase = GMDB
            SQL = ""
            SQL = SQL & "SELECT * FROM Service "
            SQL = SQL & "WHERE [Status]='Open' "
            SQL = SQL & "AND [ServiceOnDate]=#" & DeliveryDate & "# "
            SQL = SQL & "ORDER BY [ServiceOrderNo]"
            Service.DataAccess.Records_OpenSQL(SQL)

            Do While Service.DataAccess.Records_Available
                K = "LOC " & I & " - SC#" & Service.ServiceOrderNo
                'Li = lvwAllStops.ListItems.Add(, K, "Serv " & Service.ServiceOrderNo & ", " & Service.LastName & " (L" & I & ")", , "service")
                Li = lvwAllStops.Items.Add(K, "Serv " & Service.ServiceOrderNo & ", " & Service.LastName & " (L" & I & ")", 1)
                SetStopInfoByLI(Li, I, "Service", Service.ServiceOrderNo, Service.LastName, Service.MailIndex, Service.StopStart, Service.StopEnd, 0)
            Loop
            DisposeDA(Service)
        Next

        DisposeDA(Li)
    End Sub

    Private Sub SelectAllStops(Optional ByVal Remove As Boolean = False)
        Dim I As Integer, Li As ListViewItem
        If Remove Then
            'For I = lvwThisTruck.ListItems.Count To 1 Step -1
            For I = lvwThisTruck.Items.Count To 1 Step -1
                'SelectStop lvwThisTruck.ListItems(I).key, True
                SelectStop(lvwThisTruck.Items(I).ImageKey, True)
            Next
        Else
            'For I = 1 To lvwAllStops.ListItems.Count
            For I = 1 To lvwAllStops.Items.Count
                'Set Li = lvwAllStops.ListItems(I)
                Li = lvwAllStops.Items(I)
                'If Not Li.Ghosted Then SelectStop Li.key ----COMMENTED THIS LINE. BECAUSE GHOSTED PROPERTY IS NOT IN VB.NET. NEED TO FIND AN ALTERNATIVE.
            Next
        End If
    End Sub

    Public Sub RouteThisTruck(Optional ByVal DontRoute As Boolean = False)
        Dim I As Integer, R As Object, M As Map
        Dim FR As FindResults, Pin As Pushpin, WP As Waypoint

        DoControls(False)

        If Not DontRoute Then
            Network = New TSPNetwork
            Network.Setup(GetOptimizationSetting("StartTime"), GetOptimizationSetting("CostPerMile"), GetOptimizationSetting("CostPerHour"), GetOptimizationSetting("TimePerStop"), HiddenMap)
        End If

        DoControls(False, True)
        If DontRoute Then
            R = Network.GetResultSet        ' they might have already built this!
        Else
            R = OptimizeStops()               ' this could take a while...
        End If
        'If IsEmpty(R) Then
        If IsNothing(R) Then
            DoControls(True)
            Exit Sub
        End If
        DoControls(False)

        If Network.Count <= 1 Then
            MessageBox.Show("No stops selected.", "Nothing To Do", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            DoControls(True)
            Exit Sub
        End If

        If mapDelivery.ActiveMap Is Nothing Then
            DoControls(True)
            Exit Sub
        End If

        mapDelivery.Visible = False
        ProgressForm(0, 1, "Drawing Map...")
        M = mapDelivery.ActiveMap

        M.ActiveRoute.Clear()

        For I = LBound(R) To UBound(R)
            FR = M.FindAddressResults(R(I, tspRS.tspRS_Address), R(I, tspRS.tspRS_City), , R(I, tspRS.tspRS_State), R(I, tspRS.tspRS_Zip), geoCountryUnitedStates)
            If FR.Count < 1 Then
                Err.Raise(-1, , "Bad Address after optimization: " & R(I, tspRS.tspRS_Address) & " | " & R(I, tspRS.tspRS_City) & ", " & R(I, tspRS.tspRS_State) & " | " & R(I, 13))
            End If
            Pin = M.AddPushpin(FR(1), R(I, tspRS.tspRS_Name))
            WP = M.ActiveRoute.Waypoints.Add(Pin, Pin.Name)
            If I <> UBound(R) Then
                WP.StopTime = (R(I, tspRS.tspRS_StopTime) + R(I + 1, tspRS.tspRS_Delay)) / 60 * geoOneHour
            Else
                WP.StopTime = 0
            End If
        Next

        M.ActiveRoute.Waypoints(1).PreferredDeparture = DateAdd("n", R(1, tspRS.tspRS_Delay) * 2, TimeValue(Network.StartTime))
        M.ActiveRoute.Calculate()
        M.ActiveRoute.Directions.Location.GoTo()
        M.Saved = True
        mapDelivery.Visible = True

        ProgressForm()

        DoControls(True)
    End Sub

    Private Sub SetStopInfoByLI(ByRef Li As ListViewItem, ByVal Location As Integer, ByVal StopType As String, ByVal StopID As String, ByVal StopName As String, ByVal StopMail As Integer, ByVal StopStart As String, ByVal StopEnd As String, ByVal Cubes As Double)
        Dim CI As clsMailRec
        'Li.SubItems(1) = "" & Location
        Li.SubItems.Add("" & Location)
        'Li.SubItems(2) = StopType
        Li.SubItems.Add(StopType)
        'Li.SubItems(3) = StopID
        Li.SubItems.Add(StopID)
        'Li.SubItems(4) = StopName
        Li.SubItems.Add(StopName)
        'Li.SubItems(5) = StopMail
        Li.SubItems.Add(StopMail)


        CI = New clsMailRec
        CI.DataAccess.DataBase = GetDatabaseAtLocation(Location)
        If CI.Load(StopMail, "#Index") Then
            'Li.SubItems(6) = CI.Address
            Li.SubItems.Add(CI.Address)
            'Li.SubItems(7) = CI.City & " " & CI.Zip
            Li.SubItems.Add(CI.City & " " & CI.Zip)
        End If
        'Li.SubItems(8) = StopStart
        Li.SubItems.Add(StopStart)
        'Li.SubItems(9) = StopEnd
        Li.SubItems.Add(StopEnd)
        'Li.SubItems(10) = Format(Cubes, "0.00")
        Li.SubItems.Add(Support.Format(Cubes, "0.00"))
        DisposeDA(CI)
    End Sub

    Private Sub SelectStop(ByVal key As String, Optional ByVal Remove As Boolean = False)
        Dim Li As ListViewItem, LI2 As ListViewItem
        On Error Resume Next
        'Li = lvwAllStops.ListItems(key)
        Li = lvwAllStops.Items.Item(key)

        If Li Is Nothing Then Exit Sub

        'LI2 = lvwThisTruck.ListItems(key)
        LI2 = lvwThisTruck.Items.Item(key)

        If Not LI2 Is Nothing Then  ' already in, watch for remove
            If Remove Then
                'lvwThisTruck.ListItems.Remove LI2.key
                lvwThisTruck.Items.RemoveByKey(LI2.ImageKey)
                'Li.Ghosted = False ----> COMMENTED THIS LINE BECAUSE GHOSTED PROPERTY IS NOT AVAILABLE IN VB.NET. NEED TO FIND REPLACEMENT.
                'Li.StateImageIndex = 0  ---> ADDED THIS LINE AS A REPLACEMENT GHOSTED PROPERTY. NEED TO TEST BY ADDING ONE MORE DISABLED TYPE OF IMAGE TO IMAGELIST CONTROL.
            End If
        Else
            If Not Remove Then
                'lvwThisTruck.ListItems.Add , Li.key, Li.Text, , IIf(LCase(Left(Li.Text, 4)) = "sale", "stop", "service")
                lvwThisTruck.Items.Add(Li.ImageKey, Li.Text, IIf(LCase(Microsoft.VisualBasic.Left(Li.Text, 4)) = "sale", 2, 1))
                'Li.Ghosted = True  - COMMENTED THIS LINE BECAUSE GHOSTED PROPERTY IS NOT AVAILABLE IN VB.NET. NEED TO FIND A REPLACEMENT FOR THIS PROPERTY.
                'Li.StateImageIndex = 0  ---> ADDED THIS LINE AS A REPLACEMENT GHOSTED PROPERTY. NEED TO TEST BY ADDING ONE MORE DISABLED TYPE OF IMAGE TO IMAGELIST CONTROL.
            End If
        End If
        UpdateCubes()
    End Sub

    Private Sub DoControls(ByVal Enabled As Boolean, Optional ByVal Working As Boolean = False)
        'MousePointer = IIf(Enabled, vbDefault, vbHourglass)
        Me.Cursor = IIf(Enabled, Cursors.Default, Cursors.WaitCursor)
        cmdAddAll.Enabled = Enabled
        cmdDetails.Enabled = Enabled
        cmdDone.Enabled = Enabled
        cmdPrint.Enabled = Enabled
        cmdRemoveAll.Enabled = Enabled
        cmdShow.Enabled = Enabled
        cmdSplit.Enabled = Enabled
        cmbPrintType.Enabled = Enabled
        cmdConfigure.Enabled = Enabled
        cmdAdjust.Enabled = Enabled
        cmdManifest.Enabled = Enabled

        cmdCancel.Visible = Working
        cmdConfigure.Visible = Not Working
        cmdAdjust.Visible = Not Working
    End Sub

    Private Sub frmDeliveryMap_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActiveLog("frmDeliveryMap::Form_Load...", 8)
        SetButtonImage(cmdDone)
        SetButtonImage(cmdPrint)
        SetButtonImage(cmdSplit, "back")
        SetButtonImage(cmdCancel)
        SetButtonImage(cmdConfigure, "calendar")
        SetButtonImage(cmdAdjust, "plus")
        On Error Resume Next
        LoadHiddenMap()

        mapDelivery.NewMap(MapPointPTT)
        mapDelivery.Toolbars.Item("Standard").Visible = True
        mapDelivery.PaneState = GeoPaneState.geoPaneRoutePlanner

        LoadPrintTypes()
        ActiveLog("frmDeliveryMap::...Form_Load", 8)
    End Sub

    Public Sub LoadPrintTypes()
        cmbPrintType.Items.Clear()
        cmbPrintType.Items.Add("Strips")
        cmbPrintType.Items.Add("Turns")
        cmbPrintType.Items.Add("Full")
        cmbPrintType.Items.Add("Map")
        cmbPrintType.Items.Add("CSV")
        cmbPrintType.Text = "Strips"
    End Sub

    Private Sub LoadHiddenMap()
        ActiveLog("frmDeliveryMap::LoadHiddenMap...", 6)
        HiddenMapControl = CreateObject("Mappoint.Application")
        HiddenMap = HiddenMapControl.NewMap(MapPointPTT)
        ActiveLog("frmDeliveryMap::LoadHiddenMap - Complete", 6)
        '  Set XYMap = HiddenMapControl.NewMap(MapPointPTT)
    End Sub

    Public ReadOnly Property MapPointPTT() As String
        Get
            'Dim Sources(1 To 5) As String, sC as integer
            Dim Sources(0 To 4) As String, sC As Integer
            'Dim EndPoints(1 To 7) As String, Ep as integer
            Dim EndPoints(0 To 6) As String, Ep As Integer
            Dim F As String

            'Sources(1) = SpecialFolder(FolderEnum.feProgramFiles)
            'Sources(2) = Replace(Sources(1), " (x86)", "")
            'Sources(3) = SpecialFolder(FolderEnum.feUserAppData) & "\"
            'Sources(4) = ParentDirectory(SpecialFolder(FolderEnum.feUserAppData) & "\")
            'Sources(5) = SpecialFolder(FolderEnum.feCommonAppData) & "\"

            Sources(0) = SpecialFolder(FolderEnum.feProgramFiles)
            Sources(1) = Replace(Sources(1), " (x86)", "")
            Sources(2) = SpecialFolder(FolderEnum.feUserAppData) & "\"
            Sources(3) = ParentDirectory(SpecialFolder(FolderEnum.feUserAppData) & "\")
            Sources(4) = SpecialFolder(FolderEnum.feCommonAppData) & "\"

            'EndPoints(1) = "Microsoft MapPoint\Templates\"
            'EndPoints(2) = "Microsoft\Microsoft MapPoint\16.0\Templates\"
            'EndPoints(3) = "Microsoft\Microsoft MapPoint\17.0\Templates\"
            'EndPoints(4) = "Microsoft\Microsoft MapPoint\18.0\Templates\"
            'EndPoints(5) = "Microsoft\Microsoft MapPoint\19.0\Templates\"
            'EndPoints(6) = "Microsoft\Microsoft MapPoint\20.0\Templates\"
            'EndPoints(7) = "Local\VirtualStore\Program Files (x86)\Microsoft MapPoint\Templates\"

            EndPoints(0) = "Microsoft MapPoint\Templates\"
            EndPoints(1) = "Microsoft\Microsoft MapPoint\16.0\Templates\"
            EndPoints(2) = "Microsoft\Microsoft MapPoint\17.0\Templates\"
            EndPoints(3) = "Microsoft\Microsoft MapPoint\18.0\Templates\"
            EndPoints(4) = "Microsoft\Microsoft MapPoint\19.0\Templates\"
            EndPoints(5) = "Microsoft\Microsoft MapPoint\20.0\Templates\"
            EndPoints(6) = "Local\VirtualStore\Program Files (x86)\Microsoft MapPoint\Templates\"

            F = "New North American Map.ptt"

            For sC = LBound(Sources) To UBound(Sources)
                For Ep = LBound(EndPoints) To UBound(EndPoints)
                    MapPointPTT = Sources(sC) & EndPoints(Ep) & F
                    '      Debug.Print MapPointPTT
                    If FileExists(MapPointPTT) Then Exit Property
                Next
            Next

            If Not FileExists(MapPointPTT) Then
                MapPointPTT = InputBox("Enter North American Template File:", "Template File Not Found", MapPointPTT)
            End If

            If Not FileExists(MapPointPTT) Then MapPointPTT = ""
        End Get
    End Property

    Private Sub frmDeliveryMap_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        Dim H As Integer, L As Integer
        On Error Resume Next
        'fraMapContainer.Move 120, 120, ScaleWidth - cmdDone.Width - 240, ScaleHeight - 240
        fraMapContainer.Location = New Point(120, 120)
        fraMapContainer.Size = New Size(Width - cmdDone.Width - 240, Height - 240)
        'fraSplitLoads.Move fraMapContainer.Left, fraMapContainer.Top, fraMapContainer.Width, fraMapContainer.Height
        fraSplitLoads.Location = New Point(fraMapContainer.Left, fraMapContainer.Top)
        fraSplitLoads.Size = New Size(fraMapContainer.Width, fraMapContainer.Height)
        'mapDelivery.Move 60, 60, fraMapContainer.Width - 120, fraMapContainer.Height - 120
        mapDelivery.Location = New Point(60, 60)
        mapDelivery.Size = New Size(fraMapContainer.Width - 120, fraMapContainer.Height - 120)
        H = cmdDone.Height
        L = Width - cmdDone.Width - 60
        'cmdDone.Move L, ScaleHeight - H - 60
        cmdDone.Location = New Point(L, Height - H - 60)
        'cmdPrint.Move L, cmdDone.Top - H
        cmdPrint.Location = New Point(L, cmdDone.Top - H)
        'cmbPrintType.Move L, cmdPrint.Top - cmbPrintType.Height - 120
        cmbPrintType.Location = New Point(L, cmdPrint.Top - cmbPrintType.Height - 120)

        'cmdSplit.Move L, cmdDone.Top - 2200
        cmdSplit.Location = New Point(L, cmdDone.Top - 2200)
        'cmdConfigure.Move L, cmdSplit.Top - H - 60
        cmdConfigure.Location = New Point(L, cmdSplit.Top - H - 60)
        'cmdCancel.Move L, cmdConfigure.Top
        cmdCancel.Location = New Point(L, cmdConfigure.Top)
        'cmdAdjust.Move L, cmdConfigure.Top - H - 60
        cmdAdjust.Location = New Point(L, cmdConfigure.Top - H - 60)
    End Sub

    Private Sub frmDeliveryMap_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'This code block is from FORM QUERYUNLOAD event of vb6.0
        Select Case e.CloseReason
            'Case vbFormControlMenu, vbFormCode
            Case e.CloseReason = CloseReason.UserClosing
                If Not Printed And Not IsDevelopment() Then
                    If MessageBox.Show("Do you want to print the route first?", "Print Route - WinCDS", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        'cmdPrint.Value = True
                        cmdPrint.PerformClick()
                        If Not Printed Then e.Cancel = True
                    End If
                End If
        End Select

        'The below code is from FORM UNLOAD event of vb6.0
        On Error Resume Next
        HiddenMap = Nothing
        HiddenMapControl.ActiveMap.Saved = True
        HiddenMapControl.Application.Quit()
        HiddenMapControl = Nothing

        mapStops.ActiveMap.Saved = True
        mapStops.CloseMap()

        mapDelivery.ActiveMap.Saved = True
        mapDelivery.CloseMap()
        If Not mapDelivery.ActiveMap Is Nothing Then e.Cancel = 1

        'Unload frmOptimizeRoute
        frmOptimizeRoute.Close()
        'Unload frmOptimizeConfig
        frmOptimizeConfig.Close()
        'Unload frmOptimize
        frmOptimize.Close()
        Calendar.Show()
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Dim GPA As GeoPrintArea
        On Error GoTo PrintFailure
        Select Case LCase(cmbPrintType.Text)
            Case "turns" : GPA = GeoPrintArea.geoPrintTurnByTurn
            Case "full" : GPA = GeoPrintArea.geoPrintFullPage
            Case "dirs" : GPA = GeoPrintArea.geoPrintDirections
            Case "map" : GPA = GeoPrintArea.geoPrintMap
            Case "csv"
                Dim F As String
                F = DevOutputFolder() & "route.csv"
                Network.SaveNetwork(F)
                MessageBox.Show("Route saved to: " & F)
                Printed = True
                TruckIsRouted
                Exit Sub
            Case Else : GPA = GeoPrintArea.geoPrintStripMaps
        End Select
        mapDelivery.ActiveMap.PrintOut(, Me.Text, , GPA)
        Printed = True
        TruckIsRouted
        Exit Sub
PrintFailure:
        MessageBox.Show(Err.Description, "Printer Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub TruckIsRouted()
        'lvwThisTruck.ListItems.Clear
        lvwThisTruck.Items.Clear()
    End Sub

    Private Function OptimizeStops() As Object
        Dim I As Integer, LC As Integer, Ty As String, ID As String, Nm As String, MI As Integer
        Dim WF As String, WT As String
        Dim CI As clsMailRec, Shipping As MailNew2
        Dim StopTime As Integer, WFrom As Date, WTo As Date

        Network.AddLocation(StoreSettings.Name, StoreSettings.Address, GetWinCDSCity(StoreSettings.City), GetWinCDSState(StoreSettings.City), GetWinCDSZip(StoreSettings.City))
        For I = 1 To lvwThisTruck.Items.Count
            GetStopInfo(lvwThisTruck.Items(I).ImageKey, LC, Ty, ID, Nm, MI, WF, WT)
            CI = New clsMailRec
            CI.DataAccess.DataBase = GetDatabaseAtLocation(LC)
            If CI.Load(MI, "#Index") Then
                modMail.Mail2_GetAtIndex(CStr(CI.Index), Shipping, LC)
                StopTime = GetOptimizationSetting("TimePerStop")
                WFrom = IIf(IsDate(WF), WF, #12:00:00 AM#)
                WTo = IIf(IsDate(WT), WT, #11:59:59 PM#)
                If DateAfter(WFrom, WTo, True, "n") Then    ' make sure it's ok
                    WFrom = #12:00:00 AM#
                    WTo = #11:59:59 PM#
                End If

                If Shipping.Address2 <> "" Then
                    Network.AddLocation(Ty & " " & ID & ", " & Shipping.ShipToLast, Shipping.Address2, GetWinCDSCity(Shipping.City2), GetWinCDSState(Shipping.City2), Shipping.Zip2, StopTime, WFrom, WTo)
                Else
                    Network.AddLocation(Ty & " " & ID & ", " & CI.Last, CI.Address, GetWinCDSCity(CI.City), GetWinCDSState(CI.City), CI.Zip, StopTime, WFrom, WTo)
                End If
            End If
            DisposeDA(CI)
        Next

        Network.Solve()
        OptimizeStops = Network.GetResultSet
        'Unload frmOptimize
        frmOptimize.Close()
    End Function

    Private Sub UpdateCubes()
        Dim I As Integer, X As Double, LC As Integer, Ty As String, ID As String, Cb As Double
        X = 0
        'For I = 1 To lvwAllStops.ListItems.Count
        For I = 1 To lvwAllStops.Items.Count
            'GetStopInfo lvwAllStops.ListItems(I).key, LC, Ty, ID, , , , , Cb
            GetStopInfo(lvwAllStops.Items(I).ImageKey, LC, Ty, ID, , , , , Cb)
            If Ty = "Sale" Then
                X = X + Cb
            End If
        Next
        lblAllStopsCubes.Text = "Total Cubes: " & Support.Format(X, "0.00")

        X = 0
        'For I = 1 To lvwThisTruck.ListItems.Count
        For I = 1 To lvwThisTruck.Items.Count
            'GetStopInfo lvwThisTruck.ListItems(I).key, LC, Ty, ID, , , , , Cb
            GetStopInfo(lvwThisTruck.Items(I).ImageKey, LC, Ty, ID, , , , , Cb)
            If Ty = "Sale" Then
                X = X + Cb
            End If
        Next
        lblCurrentTruckCubes.Text = "Total Cubes: " & Support.Format(X, "0.00")
    End Sub

    Private Sub GetStopInfo(ByVal key As String, Optional ByRef Location As Integer = 0, Optional ByRef StopType As String = "", Optional ByRef StopID As String = "", Optional ByRef StopName As String = "", Optional ByRef StopMail As Integer = 0, Optional ByRef StopStart As String = "", Optional ByRef StopEnd As String = "", Optional ByRef Cubes As Double = 0#)
        Dim Li As ListViewItem

        On Error Resume Next
        'Li = lvwAllStops.ListItems(key)
        Li = lvwAllStops.Items.Item(key)

        If Li Is Nothing Then Exit Sub

        Location = Val(Li.SubItems(1).Text)
        'StopType = Li.SubItems(2)
        StopType = Li.SubItems.Item(2).Text
        'StopID = Li.SubItems(3)
        StopID = Li.SubItems(3).Text
        'StopName = Li.SubItems(4)
        StopName = Li.SubItems(4).Text
        StopMail = Val(Li.SubItems(5).Text)
        StopStart = Li.SubItems(8).Text
        StopEnd = Li.SubItems(9).Text
        Cubes = Val(Li.SubItems(10))
    End Sub

End Class