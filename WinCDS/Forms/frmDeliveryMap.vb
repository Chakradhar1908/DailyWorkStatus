Imports MapPoint
Imports MapPoint.GeoCountry
Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class frmDeliveryMap
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
        'SelectAllStops
        If UseOrderList Then
            'cmdSplit.Value = True
            cmdSplit.PerformClick()
            Show() 'vbModal
            Exit Function
        End If

        'RouteThisTruck
        Show()
    End Function

    Private Function AddWaypoint(ByRef M As Map, Optional ByVal Street As String = "", Optional ByVal City As String = "", Optional ByVal OtherCity As String = "", Optional ByVal State As String = "", Optional ByVal Zip As String = "", Optional ByVal Country As Long = 0, Optional ByVal PointName As String = "", Optional ByVal DoSelect As Boolean = False) As Waypoint
        Dim FR As FindResults, Pin As Pushpin, I As Long, L As Location
        On Error GoTo BadWpt
        FR = M.FindAddressResults(Street, City, OtherCity, State, Zip, Country)
        On Error GoTo 0
        If FR.Count > 0 Then
            If FR.Count > 1 Then
                If FR.ResultsQuality > 1 Then
                    ' Choose an option, or cancel?
                    Dim AddrArray() As Object
                    'ReDim AddrArray(1 To FR.Count)
                    For I = 1 To FR.Count
                        AddrArray(I) = FR.Item(I).Name
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

    Private Sub LoadAllStops(ByVal DeliveryDate As String, ByVal FirstStore As Long, ByVal LastStore As Long)
        Dim WP As Waypoint
        Dim GMDB As String, SQL As String
        Dim I As Long, K As String
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
                    Li = lvwAllStops.Items.Add("Sale " & Sales.SaleNo & ", " & GetMailLastNameByIndex(Sales.Index, I, True) & " (L" & I & ", " & Support.Format(Cubes, "##,##0.00"), K)
                    'SetStopInfoByLI(Li, I, "Sale", Sales.SaleNo, GetMailLastNameByIndex(Sales.Index, I, True), Sales.Index, Sales.StopStart, Sales.StopEnd, Cubes)
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
                Li = lvwAllStops.Items.Add("Serv " & Service.ServiceOrderNo & ", " & Service.LastName & " (L" & I & ")")
                'SetStopInfoByLI(Li, I, "Service", Service.ServiceOrderNo, Service.LastName, Service.MailIndex, Service.StopStart, Service.StopEnd, 0)
            Loop
            DisposeDA(Service)
        Next

        DisposeDA(Li)
    End Sub

    '    Private Sub SelectAllStops(Optional ByVal Remove As Boolean = False)
    '        Dim I As Long, Li As ListViewItem
    '        If Remove Then
    '            'For I = lvwThisTruck.ListItems.Count To 1 Step -1
    '            For I = lvwThisTruck.Items.Count To 1 Step -1
    '                SelectStop lvwThisTruck.ListItems(I).key, True
    '                    lvwThisTruck.Items.Add()
    '            Next
    '                Else
    '            For I = 1 To lvwAllStops.ListItems.Count
    '      Set Li = lvwAllStops.ListItems(I)
    '      If Not Li.Ghosted Then SelectStop Li.key
    '    Next
    '        End If
    '    End Sub

    '    Public Sub RouteThisTruck(Optional ByVal DontRoute As Boolean = False)
    '        Dim I As Long, R As Variant, M As Map
    '        Dim FR As FindResults, Pin As Pushpin, WP As Waypoint

    '        DoControls False

    '  If Not DontRoute Then
    '    Set Network = New TSPNetwork
    '    Network.Setup GetOptimizationSetting("StartTime"), GetOptimizationSetting("CostPerMile"), GetOptimizationSetting("CostPerHour"), GetOptimizationSetting("TimePerStop"), HiddenMap
    '  End If

    '        DoControls False, True
    '  If DontRoute Then
    '            R = Network.GetResultSet        ' they might have already built this!
    '        Else
    '            R = OptimizeStops               ' this could take a while...
    '        End If
    '        If IsEmpty(R) Then
    '            DoControls True
    '    Exit Sub
    '        End If
    '        DoControls False

    '  If Network.Count <= 1 Then
    '            MsgBox "No stops selected.", vbExclamation, "Nothing To Do"
    '    DoControls True
    '    Exit Sub
    '        End If

    '        If mapDelivery.ActiveMap Is Nothing Then
    '            DoControls True
    '    Exit Sub
    '        End If

    '        mapDelivery.Visible = False
    '        ProgressForm 0, 1, "Drawing Map..."
    '  Set M = mapDelivery.ActiveMap

    '  M.ActiveRoute.Clear()

    '        For I = LBound(R) To UBound(R)
    '    Set FR = M.FindAddressResults(R(I, tspRS_Address), R(I, tspRS_City), , R(I, tspRS_State), R(I, tspRS_Zip), geoCountryUnitedStates)
    '    If FR.Count < 1 Then
    '                Err.Raise -1, , "Bad Address after optimization: " & R(I, tspRS_Address) & " | " & R(I, tspRS_City) & ", " & R(I, tspRS_State) & " | " & R(I, 13)
    '    End If
    '    Set Pin = M.AddPushpin(FR(1), R(I, tspRS_Name))
    '    Set WP = M.ActiveRoute.Waypoints.Add(Pin, Pin.Name)
    '    If I <> UBound(R) Then
    '                WP.StopTime = (R(I, tspRS_StopTime) + R(I + 1, tspRS_Delay)) / 60 * geoOneHour
    '            Else
    '                WP.StopTime = 0
    '            End If
    '        Next

    '        M.ActiveRoute.Waypoints(1).PreferredDeparture = DateAdd("n", R(1, tspRS_Delay) * 2, TimeValue(Network.StartTime))
    '        M.ActiveRoute.Calculate()
    '        M.ActiveRoute.Directions.Location.GoTo()
    '        M.Saved = True
    '        mapDelivery.Visible = True

    '        ProgressForm()

    '        DoControls True
    'End Sub

    '    Private Sub SetStopInfoByLI(ByRef Li As ListViewItem, ByVal Location As Long, ByVal StopType As String, ByVal StopID As String, ByVal StopName As String, ByVal StopMail As Long, ByVal StopStart As String, ByVal StopEnd As String, ByVal Cubes As Double)
    '        Dim CI As clsMailRec
    '        Li.SubItems(1) = "" & Location
    '        Li.SubItems(2) = StopType
    '        Li.SubItems(3) = StopID
    '        Li.SubItems(4) = StopName
    '        Li.SubItems(5) = StopMail
    '  Set CI = New clsMailRec
    '  CI.DataAccess.DataBase = GetDatabaseAtLocation(Location)
    '        If CI.Load(StopMail, "#Index") Then
    '            Li.SubItems(6) = CI.Address
    '            Li.SubItems(7) = CI.City & " " & CI.Zip
    '        End If
    '        Li.SubItems(8) = StopStart
    '        Li.SubItems(9) = StopEnd
    '        Li.SubItems(10) = Format(Cubes, "0.00")
    '        DisposeDA CI
    'End Sub

End Class