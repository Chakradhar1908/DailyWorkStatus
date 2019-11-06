Module modCDSCustomers
    Public ReadOnly Property IsUFO() As Boolean
        Get
            IsUFO = CheckStoreName("united", "ufo", "the warehouse")
        End Get
    End Property
    Public ReadOnly Property IsPalazzo() As Boolean
        Get
            IsPalazzo = CheckStoreName("Palazzo")
        End Get
    End Property
    Public ReadOnly Property IsGrizzlys() As Boolean
        Get
            IsGrizzlys = CheckStoreName("GRIZZLY'S", "BEAR NAKED")
        End Get
    End Property
    Public ReadOnly Property IsFurnOne() As Boolean
        Get
            IsFurnOne = CheckStoreName("FURNITURE ONE")
        End Get
    End Property

    Public ReadOnly Property IsWilkenfeld() As Boolean
        Get
            IsWilkenfeld = CheckStoreName("Wilkenfeld")
        End Get
    End Property
    Public ReadOnly Property IsBFMyer() As Boolean
        Get
            IsBFMyer = CheckStoreName("B. F. Myers Furniture", "BFMyer")
        End Get
    End Property
    Public ReadOnly Property IsBarrs() As Boolean
        Get
            IsBarrs = CheckStoreName("Barr")
        End Get
    End Property
    Public ReadOnly Property IsSleepCity() As Boolean
        Get
            IsSleepCity = CheckStoreName("Sleep City")
        End Get
    End Property
    Public ReadOnly Property IsSleepingSystems() As Boolean
        Get
            IsSleepingSystems = CheckStoreName("Sleeping Systems")
        End Get
    End Property
    Public Function CheckStoreName(ByVal CHK1 As String, Optional ByVal CHK2 As String = "", Optional ByVal CHK3 As String = "", Optional ByVal CHK4 As String = "", Optional ByVal CHK5 As String = "", Optional ByVal CHK6 As String = "", Optional ByVal CHK7 As String = "", Optional ByVal CHK8 As String = "", Optional ByVal CHK9 As String = "", Optional ByVal CHK10 As String = "") As Boolean

        CheckStoreName = False
        '* Determines whether the store name in the store settings (INI) matches one of the check string
        If CheckStoreName1(CHK1) Then CheckStoreName = True
        If CheckStoreName1(CHK2) Then CheckStoreName = True
        If CheckStoreName1(CHK3) Then CheckStoreName = True
        If CheckStoreName1(CHK4) Then CheckStoreName = True
        If CheckStoreName1(CHK5) Then CheckStoreName = True
        If CheckStoreName1(CHK6) Then CheckStoreName = True
        If CheckStoreName1(CHK7) Then CheckStoreName = True
        If CheckStoreName1(CHK8) Then CheckStoreName = True
        If CheckStoreName1(CHK9) Then CheckStoreName = True
        If CheckStoreName1(CHK10) Then CheckStoreName = True
    End Function
    Private Function CheckStoreName1(ByVal CHK As String, Optional ByVal SN As String = "") As Boolean
        Dim R As String, I as integer

        CheckStoreName1 = False
        If IP_CONTROL Then
            If CHK = ExternalIPAddress Then CheckStoreName1 = True
        End If

        If SN = "" Then
            SN = UCase(StoreSettings(Optimize:=True).Name)
            If SN = "" Then SN = UCase(StoreSettings(1).Name)
            If SN = "" Then Exit Function
        End If

        SN = UCase(SN)
        R = ""
        For I = 1 To Len(SN)
            If Asc(Mid(SN, I, 1)) >= 65 And Asc(Mid(SN, I, 1)) <= 90 Then R = R & Mid(SN, I, 1)
        Next
        SN = R

        CHK = UCase(CHK)
        R = ""
        For I = 1 To Len(CHK)
            If Asc(Mid(CHK, I, 1)) >= 65 And Asc(Mid(CHK, I, 1)) <= 90 Then R = R & Mid(CHK, I, 1)
        Next
        CHK = R

        If Len(CHK) > 0 And (Left(SN, Len(CHK)) = CHK) Then CheckStoreName1 = True
    End Function
    Public ReadOnly Property IsDoddsLtd() As Boolean
        Get
            IsDoddsLtd = CheckStoreName("Dodd's Furniture Ltd.", "Dodds")
        End Get
    End Property

    Public ReadOnly Property IsPitUSA() As Boolean
        Get
            IsPitUSA = CheckStoreName("Pitusa")
        End Get
    End Property

    Public ReadOnly Property IsSidesFurniture() As Boolean
        Get
            IsSidesFurniture = CheckStoreName("Sides Furniture")
        End Get
    End Property

    'Public Property Get IsSleepWorks() As Boolean :  IsSleepWorks = CheckStoreName("Sleepworks"): End Property
    Public ReadOnly Property IsSleepWorks() As Boolean
        Get
            IsSleepWorks = CheckStoreName("Sleepworks")
        End Get
    End Property

    Public ReadOnly Property IsCanadian() As Boolean
        Get
            '* Determines whether the current customer is Canadian (and hence, requires GST, etc).
            IsCanadian = IsDoddsLtd()
        End Get
    End Property

    Public ReadOnly Property IsChandlers() As Boolean
        Get
            IsChandlers = CheckStoreName("Chandler")
        End Get
    End Property

    Public ReadOnly Property IsAuthenTeak() As Boolean
        Get
            IsAuthenTeak = CheckStoreName("AuthenTeak")
        End Get
    End Property

    Public ReadOnly Property IsLapeer() As Boolean
        Get
            IsLapeer = CheckStoreName("LAPEER")
        End Get
    End Property

    Public ReadOnly Property IsPuritan() As Boolean
        Get
            IsPuritan = CheckStoreName("Puritan")
        End Get
    End Property

    Public ReadOnly Property IsRockyMountain() As Boolean
        Get
            IsRockyMountain = CheckStoreName("Rocky Mountain")
        End Get
    End Property

    Public ReadOnly Property IsDecoratingOnADime() As Boolean
        Get
            IsDecoratingOnADime = CheckStoreName("Decorating on a Dime")
        End Get
    End Property

    Public ReadOnly Property IsParkPlace() As Boolean
        Get
            IsParkPlace = CheckStoreName("Park Place")
        End Get
    End Property

    Public ReadOnly Property IsThorntons() As Boolean
        Get
            IsThorntons = CheckStoreName("Thornton")
        End Get
    End Property

    Public ReadOnly Property IsElmore() As Boolean
        Get
            IsElmore = CheckStoreName("Elmore")
        End Get
    End Property

    Public ReadOnly Property IsKenLu() As Boolean
        Get
            IsKenLu = CheckStoreName("KENLU")
        End Get
    End Property

    Public ReadOnly Property IsLott() As Boolean
        Get
            IsLott = CheckStoreName("Lott")
        End Get
    End Property

    Public ReadOnly Property IsBoyd() As Boolean
        Get
            IsBoyd = CheckStoreName("Boyd")
        End Get
    End Property

    Public ReadOnly Property IsMcClure() As Boolean
        Get
            IsMcClure = CheckStoreName("McClure")
        End Get
    End Property

    Public ReadOnly Property IsYeatts() As Boolean
        Get
            IsYeatts = CheckStoreName("Yeatts")
        End Get
    End Property

    Public ReadOnly Property IsCarroll() As Boolean
        Get
            IsCarroll = CheckStoreName("Carroll")
        End Get
    End Property

    Public ReadOnly Property IsShaw() As Boolean
        Get
            IsShaw = CheckStoreName("shaw")
        End Get
    End Property

    Public ReadOnly Property IsWesternDiscount() As Boolean
        Get
            IsWesternDiscount = CheckStoreName("Western Discount")
        End Get
    End Property

    Public ReadOnly Property IsPricesFurniture() As Boolean
        Get
            IsPricesFurniture = CheckStoreName("Prices")
        End Get
    End Property

    Public Function UseIUI() As Boolean
        '* Installment Option -- Involuntary Unemployment Insurance
        UseIUI = IsTreehouse Or IsBlueSky
    End Function

    Public ReadOnly Property IsTreehouse() As Boolean
        Get
            IsTreehouse = CheckStoreName("Tree house")
        End Get
    End Property

    Public ReadOnly Property IsBlueSky() As Boolean
        Get
            IsBlueSky = CheckStoreName("Blue Sky", "Texas Discount")
        End Get
    End Property

    Public ReadOnly Property IsEvridge() As Boolean
        Get
            IsEvridge = CheckStoreName("Evridge")
        End Get
    End Property

    Public ReadOnly Property IsJeffros() As Boolean
        Get
            IsJeffros = CheckStoreName("Jeffro")
        End Get
    End Property

    Public ReadOnly Property IsChicago() As Boolean
        Get
            IsChicago = CheckStoreName("New Age", "Chicago")
        End Get
    End Property

    Public ReadOnly Property IsCarpet() As Boolean
        Get
            IsCarpet = CheckStoreName("Carpet")
        End Get
    End Property

    Public ReadOnly Property IsMichaels() As Boolean
        Get
            IsMichaels = CheckStoreName("Michael")
        End Get
    End Property

    Public ReadOnly Property IsRogers() As Boolean
        Get
            IsRogers = CheckStoreName("Roger")
        End Get
    End Property

    Public ReadOnly Property IsWoodPeckers() As Boolean
        Get
            IsWoodPeckers = CheckStoreName("Woodpeckers")
        End Get
    End Property

    Public ReadOnly Property IsStudioD() As Boolean
        Get
            IsStudioD = CheckStoreName("Studio D")
        End Get
    End Property

    Public ReadOnly Property IsLoft() As Boolean
        Get
            IsLoft = CheckStoreName("Loft Home")
        End Get
    End Property

    Public ReadOnly Property IsFurnitureWorld() As Boolean
        Get
            IsFurnitureWorld = CheckStoreName("Furniture World")
        End Get
    End Property

    Public ReadOnly Property IsPayless() As Boolean
        Get
            IsPayless = CheckStoreName("Payless")
        End Get
    End Property

    Public ReadOnly Property IsClassicInteriors() As Boolean
        Get
            IsClassicInteriors = CheckStoreName("Classic Interior")
        End Get
    End Property

    Public ReadOnly Property IsHouseOfBedroomsKids() As Boolean
        Get
            IsHouseOfBedroomsKids = CheckStoreName("House of Bedrooms kids")
        End Get
    End Property

    Public ReadOnly Property IsHouseOfBedrooms() As Boolean
        Get
            IsHouseOfBedrooms = CheckStoreName("House of Bedrooms") And Not CheckStoreName("House of Bedrooms kids")
        End Get
    End Property

    Public ReadOnly Property IsTenPenny() As Boolean
        Get
            IsTenPenny = CheckStoreName("Tenpenny")
        End Get
    End Property

    Public ReadOnly Property IsFurnitureStoreOfKansas() As Boolean
        Get
            IsFurnitureStoreOfKansas = CheckStoreName("Furniture Store of Kansas", "The Furniture Store of Kansas")
        End Get
    End Property
End Module
