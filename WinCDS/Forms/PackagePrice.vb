Public Class PackagePrice
    Dim ItemList As Object  '(800, 7) As String
    Dim Lines As Integer
    Dim Row As Integer
    Dim NewRecord As String

    Dim TotLanded As Decimal
    Dim TotOnSale As Decimal
    Dim TotList As Decimal
    Dim PackageGM As Double
    Dim List As String

    Private OnSale As String
    Private PackageSale As String
    Private Desc As String
    Private Comments As String
    Private Landed As String

    Private sLoading As Boolean

    Private WithEvents mDBInvKit As CDbAccessGeneral
    Private WithEvents mDBAccess As CDbAccessGeneral

    Private Const pW1 As Integer = 495
    Private Const pH1 As Integer = 495
    Private Const pL1 As Integer = 10200
    Private Const pT1 As Integer = 7200
    Private Const pW As Integer = 10335
    Private Const pH As Integer = 7935

    Public Function FindKits(Optional ByVal StyleNum As String = "") As Boolean
        Clear()

        Row = 0
        If StyleNum = "" Then Exit Function
        mDBAccess_Init(StyleNum)
        mDBAccess.GetRecord   ' this gets the record
        mDBAccess.dbClose
        mDBAccess = Nothing

        mDBInvKit_Init()
        mDBInvKit_SqlSet(StyleNum)
        FindKits = mDBInvKit.GetRecord
        mDBInvKit.dbClose
        mDBInvKit = Nothing
    End Function

    Private Sub Clear()
        'clear
        UGridIO1.Clear()
        'Unload Me
        Me.Close()
        'PackagePrice.Show()
        Me.Show()
        lstItems.items.Clear
        Lines = 0
        TotLanded = 0
        TotOnSale = 0
        TotList = 0
        Landed = 0
        OnSale = 0
        List = 0

        PackageGM = 0
        txtTotLanded.Text = ""
        txtTotOnSale.Text = ""
        txtTotList.Text = ""
        txtGM.Text = ""
        txtPackagePrice.Text = ""
        txtOrigGM.Text = ""
    End Sub

    Private Sub mDBAccess_Init(ByVal Tid As String)
        On Error GoTo HandleErr
        mDBInvKit_Init()
        mDBAccess = New CDbAccessGeneral
        mDBAccess.dbOpen(GetDatabaseAtLocation(KIT_LOC)) ' Kits are only at location 1.
        mDBAccess.SQL = "SELECT * From InvKit WHERE KitStyleNo=""" & ProtectSQL(Tid) & """"
        Exit Sub

HandleErr:
        MsgBox(" Access Init: " & Err.Description, , Err.Number)
    End Sub

    Private Sub mDBInvKit_Init()
        mDBInvKit = New CDbAccessGeneral
        mDBInvKit.dbOpen(GetDatabaseAtLocation(1))  ' Kits are only at location 1.
    End Sub

    Public Sub mDBInvKit_SqlSet(ByVal Tid As String)
        'Set mDBAccess = New CDbAccess
        mDBInvKit.SQL =
       "SELECT InvKit.*" _
       & " From InvKit" _
       & " WHERE (((InvKit.KitStyleNo)  =""" & ProtectSQL(Tid) & """))"
    End Sub

    Public Sub EditPackages()
        Show()
        GetKitInfo
    End Sub

    Private Sub GetKitInfo()
        ' If mInvCkStyle is made non-modal, cleanup is required!
        ' If InvCkStyle is changed to not include this form's code, it needs to be defined in the form and withevents.
        Dim mInvCkStyle As InvCkStyle
        mInvCkStyle = New InvCkStyle
        '  mInvCkStyle.ParentForm = Name
        'mInvCkStyle.Show vbModal, Me
        mInvCkStyle.ShowDialog(Me)
        'Unload mInvCkStyle
        mInvCkStyle.Close()
    End Sub

End Class