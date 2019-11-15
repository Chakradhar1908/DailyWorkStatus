Imports WinCDS

Public Class cDataConvert
    Private mSubClass As cDataConvert
    Private mDatabase As String
    Private mTable As String
    Private mIndex As String


    'Friend Property Get SubClass() As cDataConvert :   SubClass = mSubClass: End Property
    'Friend Property Let SubClass(ByRef T As cDataConvert) : Set mSubClass = T: End Property
    Friend Property SubClass As cDataConvert
        Get
            SubClass = mSubClass
        End Get
        Set(value As cDataConvert)
            mSubClass = value
        End Set
    End Property
    'Friend Property Get DataBase() As String :  DataBase = mDatabase: End Property
    'Friend Property Let DataBase(ByRef T As String) :   mDatabase = T: End Property
    Friend Property DataBase As String
        Get
            DataBase = mDatabase
        End Get
        Set(value As String)
            mDatabase = value
        End Set
    End Property
    'Friend Property Get Table() As String :  Table = mTable: End Property
    'Friend Property Let Table(ByRef T As String) :   mTable = T: End Property
    Friend Property Table As String
        Get
            Table = mTable
        End Get
        Set(value As String)
            mTable = value
        End Set
    End Property
    'Friend Property Get Index() As String :  Index = mIndex: End Property
    'Friend Property Let Index(ByRef T As String) :  mIndex = T: End Property
    Friend Property Index As String
        Get
            Index = mIndex
        End Get
        Set(value As String)
            mIndex = value
        End Set
    End Property
    Private Sub aFromFile()

    End Sub
    Public Sub FileOpen() : End Sub
    'Public Property Get FileRecords() as integer : End Sub
    Public ReadOnly Property FileRecords as integer
        Get
            Return 0
        End Get
    End Property
    Public Sub FileClose() : End Sub
    Public Sub SetRecordSet(ByRef Index as integer, ByRef RS As ADODB.Recordset) : End Sub
    Public Sub ConvertExceptions(ByRef RS As ADODB.Recordset) : End Sub

    Friend Function Count() as integer
        Count = getRecordsetCountByTable(File:=mDatabase, Table:=mTable)  'OVERFLOW
    End Function
    Friend Function ConvertData() As Boolean
        Dim Count As Integer

        ConvertData = False
        Count = getRecordsetCountByTable(File:=DataBase, Table:=Table)
        If (Count > 0) Then
            Debug.Print("There is data in the table")
            Debug.Print("COUNT: " & Count)
            Exit Function
        End If
        FromFile()
    End Function
    Private Sub FromFile()
        Dim DoForce As Boolean, Special As Boolean, Res As Boolean
        Dim Data As Calendar
        Dim I as integer
        Dim maxIndex as integer
        Dim RS As ADODB.Recordset
        'Dim O As cPODetail, P As cPOReceived
        Dim O As Object = New cPODetail
        Dim p As Object = New cPOReceived

        'NOTE: Commented below line because frmProgress form is using ucPBar custom control to show the progree of delivery calendar data while loading.
        'The ucPBar custom control's most of the code is not working in vb.net. Need to find an alternative in vb.net later.
        'Dim PR As frmProgress

        mSubClass.FileOpen()
        maxIndex = mSubClass.FileRecords()
        RS = getRecordsetByTableLabelIndex(File:=DataBase,
             Table:=Table, Label:=Index, Index:="-1", Always:=True)
        RS.Delete()

        If Table = "PO" Or Table = "POReceived" Then Special = True
        DoForce = False
        ' tmp patch - BFH20050705
        '  DoForce = True
        On Error Resume Next

        'Refer the above note for the below three lines commented code reason.
        'PR = New frmProgress
        'PR.AltPrg = Practice.ConversionPrg
        'PR.Progress(0, maxIndex, "Processing...", True, True)
        On Error GoTo 0

        For I = 0 To maxIndex - 1
            RS.AddNew()

            If Special Then
                If Table = "PO" Then
                    O = mSubClass
                    O.ForceUpdate = DoForce
                End If
                If Table = "POReceived" Then
                    p = mSubClass
                    p.ForceUpdate = DoForce
                End If
            End If

            mSubClass.SetRecordSet(Index:=I, RS:=RS)
            If Special Then
                If Table = "PO" Then
                    If Not O.CancelledUpdate Then DoForce = True
                End If
                If Table = "POReceived" Then
                    If Not p.CancelledUpdate Then DoForce = True
                End If
            End If
            On Error Resume Next
            'PR.Progress(I)
            On Error GoTo 0
            Application.DoEvents()
        Next

        On Error Resume Next

        'Refer the above note in this sub at top for the below two lines commented code reason.
        'PR.ProgressClose()
        'PR = Nothing
        On Error GoTo 0

        SetRecordsetByTableLabelIndex(File:=DataBase, RS:=RS, Table:=Table, Label:=Index, Index:="-1")
        On Error Resume Next
        RS.Filter = ADODB.FilterGroupEnum.adFilterPendingRecords
        If RS.RecordCount > 0 Then mSubClass.ConvertExceptions(RS)
        RS = Nothing
        mSubClass.FileClose()
    End Sub

    'Public Shared Widening Operator CType(v As CGrossMargin) As cDataConvert
    '    Throw New NotImplementedException()
    'End Operator
End Class
