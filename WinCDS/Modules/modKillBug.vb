Module modKillBug
    Private Const KILLBUG_NOTIFY_DEVELOPER as integer = 30
    Private Const KILLBUG_NOTIFY_USER as integer = 7

    Private InformCripple As Boolean
    Private mKillInit As Boolean
    Private mKillDate As Date
    Private mCrippleDate As Date

    Public Function CrippleBug(Optional ByVal Feature As String = "", Optional ByVal Silent As Boolean = False) As Boolean
        Dim S As String
        CrippleBug = IsCrippled()
        If TestCrippleBug Then CrippleBug = True

        If CrippleBug Then
            If Feature <> "" Or Not InformCripple Then
                InformCripple = True
                If Not Silent Then frmCrippleBugNotify.CrippleBug(Feature)
            End If
        End If
    End Function

    Private ReadOnly Property TestCrippleBug() As Boolean
        Get
            ' because of the potential for catastrophic errors if this is left in,
            ' it is recommended that this is enabled only for indiviual computers.
            ' Simply use the format below, or uncomment one of the existing tests:/

            '  If IsCDSComputer("LAPTOP") Then TestCrippleBug = True
            '  If IsCDSComputer("INVENTORY2") Then TestCrippleBug = True
        End Get
    End Property

    Public Function IsCrippled() As Boolean
        IsCrippled = Not PrvKill And DateAfter(Today, CrippleDate)
    End Function

    Public ReadOnly Property CrippleDate() As Date
        Get
            KillBugInit()
            CrippleDate = mCrippleDate
        End Get
    End Property

    Private Function KillBugInit() As Boolean
        ' Two useful dates you can put in anywhere for anyone:  EffectivelyNever, AlwaysOn

        If mKillInit Then Exit Function
        mKillInit = True

        mKillDate = EffectivelyNever      ' Started 5/14/2018
        '  mKillDate = #6/10/2018#          'Started 11/3/2017
        '  mKillDate = #12/11/2017#         'Started 5/9/2017
        '  mKillDate = #6/12/2017#          'Started 10-31-2016
        '  mKillDate = #12/12/2016#         'Started 05-15-2016
        '  mKillDate = #6/14/2016#          'Started 11-09-2015
        '  mKillDate = #12/14/2015#         'Started 05-29-2015
        ' on this day software dies.

        '  use this format
        '  If CheckStoreName("XXX FURN STORE") Then KillDate = #4/5/2006#
        ' If CheckStoreName("E-Z Credit") Then KillDate = #3/20/2013#
        ' If CheckStoreName("Moss") Then KillDate = #4/8/2014#
        '  If CheckStoreName("Tip Top Furniture") Then KillDate = #7/25/2014#
        If CheckStoreName("Tempo Furniture") Then mKillDate = #2/18/2016#
        If CheckStoreName("Furniture At 7707") Then mKillDate = #5/2/2016#
        If CheckStoreName("Adams") Then mKillDate = #9/2/2016#
        If CheckStoreName("Kids Only Furniture") Then mKillDate = #11/12/2016#

        If CheckStoreName("Home Fashions") Then mKillDate = #10/12/2017#
        If CheckStoreName("Jim's Home Furnish") Then mKillDate = #10/12/2017#
        If CheckStoreName("The Room Loft") Then mKillDate = #10/12/2017#

        If IsLoft Then mKillDate = #1/24/2018#
        If IsSleepWorks Then mKillDate = #1/24/2018#
        If IsFurnitureWorld Then mKillDate = #1/24/2108#
        If IsPayless Then mKillDate = #4/30/2018#
        If IsRogers Then mKillDate = #6/5/2018#

        If CheckStoreName("Ashley Home") Then mKillDate = #10/24/2018#
        If CheckStoreName("tenpenny") Then mKillDate = #10/24/2018#
        If IsClassicInteriors Then mKillDate = EffectivelyNever
        If CheckStoreName("75.159.216.237") Then mKillDate = #5/14/2017#                      ' Country corner

        If CheckStoreName("jeffro") Then mKillDate = #11/21/2018#
        If CheckStoreName("Lafayette") Then mKillDate = #11/21/2018#
        If CheckStoreName("organic sleep products") Then mKillDate = #11/21/2018#
        If IsHouseOfBedroomsKids Then mKillDate = #12/1/2018#
        If IsHouseOfBedrooms Then mKillDate = #12/1/2018#

        If CheckStoreName("Pina Furniture") Then mKillDate = #12/2/2018#


        ' Cripple dates..

        mCrippleDate = EffectivelyNever                                                       ' This line goes first..  this is always in the future!!  It will be used if no stores match

        If IsTenPenny Then mCrippleDate = #6/1/2018#                                          '
        If CheckStoreName("Puritan") Then mCrippleDate = #9/7/2015#                           '
        If CheckStoreName("Evridge's") Then mCrippleDate = #10/1/2015#                        '
        If CheckStoreName("Cranes") Then mCrippleDate = #2/1/2016#                            '
        If CheckStoreName("Bayshore Furniture & Mattress") Then mCrippleDate = #1/1/2016#     '
        If CheckStoreName("Home Wood Furniture") Then mCrippleDate = #3/22/2016#              '
        If CheckStoreName("Lucas") Then mCrippleDate = #6/22/2016#                            '
        If CheckStoreName("Hill") And IsIn(StoresSld, 2, 3) Then mCrippleDate = #7/20/2016#   ' Would cripple Hill store 2 and 3 after 2009
        If CheckStoreName("Johnson") Then mCrippleDate = #6/30/2017#                          '

        If CheckStoreName("Sleep City") Then mCrippleDate = #10/24/2018#


        If CheckStoreName("KILLED") Then mKillDate = AlwaysOn                                 ' This allows testing the software in crippled mode.
        If CheckStoreName("CRIPPLED") Then mCrippleDate = AlwaysOn                            ' This allows testing the software in crippled mode.

    End Function

    Private Function EffectivelyNever() As Date
        EffectivelyNever = YearAdd(Today, 15)
    End Function

    Private Function AlwaysOn() As Date
        AlwaysOn = NullDate
    End Function

    Public Function KillBug(Optional ByVal Silent As Boolean = False) As Boolean
        ' Warns a user their software will expire soon
        If Not Silent Then UserKillBugNotify
        KillBug = IsExpired
        If TestKillBug Then KillBug = True

        If KillBug Then
            HideSplash()
            If Not Silent Then frmKillBugNotify.Show vbModal
  Else
            CrippleBug , True
  End If
    End Function

End Module
