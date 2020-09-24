Imports WinCDS
Public Class CDataAccess
    Private mSubClass As CDataAccess
    Private mDatabase As String
    Private mTable As String
    Private mIndex As String
    Private Mrs As ADODB.Recordset
    Private mUpdated As Boolean
    Public Event RecordUpdated()
    Private mCurrentIndex As Integer

    Public Sub New()
        mCurrentIndex = -1
        mUpdated = False
    End Sub

    'Friend Property Get SubClass() As CDataAccess :  SubClass = mSubClass: End Property
    'Friend Property Let SubClass(ByRef T As CDataAccess) : Set mSubClass = T: End Property
    Friend Property SubClass As CDataAccess
        Get
            SubClass = mSubClass
        End Get
        Set(value As CDataAccess)
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
    'Friend Property Let Table(ByRef T As String) :  mTable = T: End Property
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

    Friend Function Record_Count() As Integer
        On Error Resume Next
        Record_Count = Mrs.RecordCount
    End Function

    Friend Sub Records_Add()
        ' This is slow on large tables because it's actually running a query
        ' intended to get no records.  Let's find a quicker way.
        'Set mrs = getRecordsetByTableLabelIndex
        If Mrs Is Nothing Or Not MJKTEST Then
            ' Original behavior always reset the recordset.
            ' Now, if there's a recordset already in place, we'll use it.
            Mrs = GetEmptyRecordsetByTable(File:=DataBase, Table:=Table, Always:=True)
        Else
            ' MJKTEST activates this code when the recordset is pre-set.
            ' To save querying time, it just adds a new record to the
            ' already-open recordset.
            Mrs.AddNew()
        End If
        ' Load the subclass's data into the recordset.
        mSubClass.SetRecordSet(Mrs)
        mUpdated = True
    End Sub

    Friend Sub Record_Update()
        mSubClass.SetRecordSet(Mrs)
        mUpdated = True
    End Sub

    Friend Sub Records_Update()
        If MJKTEST And Not Mrs Is Nothing Then
            ' Recordset is all ready, we just need to supply it with a connection.
            Dim tmpDBAG As New CDbAccessGeneral
            tmpDBAG.dbOpen(DataBase)
            tmpDBAG.UpdateRecordSet(Mrs)
            tmpDBAG.dbClose()
            DisposeDA(tmpDBAG)
        Else
            'SetRecordsetByTableLabelIndex(File:=DataBase, RS:=Mrs, Table:=Table, Label:=Index, Index:="-1")
            SetRecordsetByTableLabelIndex(File:=DataBase, RS:=Mrs, Table:=Table, Label:=Index, Index:="-1")
        End If
        mUpdated = False
        RaiseEvent RecordUpdated()
        Exit Sub
    End Sub

    Friend Function Records_OpenIndexAt(ByRef Index As String, Optional ByRef OrderBy As String = "") As Boolean
        Records_OpenIndexAt = Me.Records_OpenSQL(SQL:=Me.getIndexSQL(CStr(Index), OrderBy))
        mUpdated = False
    End Function

    Friend Function Records_OpenSQL(ByRef SQL As String) As Boolean
        On Error Resume Next
        Mrs = GetRecordsetBySQL _
        (File:=DataBase _
        , SQL:=SQL _
        , Always:=False)
        mCurrentIndex = -1
        mUpdated = False
        If Err.Number <> 0 Then Records_OpenSQL = False Else Records_OpenSQL = True
    End Function

    Friend Function getIndexSQL(ByRef Index As String, Optional ByRef OrderBy As String = "") As String
        getIndexSQL = getFieldIndexSQL(Field:=Me.Index, Index:=Index, OrderBy:=OrderBy)
    End Function

    Friend Function getFieldIndexSQL(ByRef Field As String, ByRef Index As String, Optional ByRef OrderBy As String = "") As String
        getFieldIndexSQL = " SELECT [" & Me.Table & "].*" _
        & " From [" & Me.Table & "]" _
        & " Where ((([" & Me.Table & "]." & Field & ") = """ & ProtectSQL(Index) & """))"
        If OrderBy <> "" Then getFieldIndexSQL = getFieldIndexSQL & " Order By " & ProtectSQL(OrderBy)
    End Function

    Friend Function Records_OpenFieldIndexAtNumber(ByRef Field As String, ByRef Index As String, Optional ByRef OrderBy As String = "") As Boolean
        Records_OpenFieldIndexAtNumber _
        = Me.Records_OpenSQL(
            SQL:=Me.getFieldIndexSQLNumber(Field, CStr(Index), OrderBy)
            )
        mUpdated = False
    End Function

    Friend Function getFieldIndexSQLNumber(ByRef Field As String, ByRef Index As String, Optional ByRef OrderBy As String = "") As String
        If Index = "" Then Index = "-1"
        getFieldIndexSQLNumber = " SELECT [" & Me.Table & "].*" _
        & " From [" & Me.Table & "]" _
        & " Where ((([" & Me.Table & "]." & Field & ") = " & Index & "))"
        If OrderBy <> "" Then getFieldIndexSQLNumber = getFieldIndexSQLNumber & " Order By " & ProtectSQL(OrderBy)
    End Function

    Friend Function Records_OpenFieldIndexAtDate(ByRef Field As String, ByRef Index As String) As Boolean
        Records_OpenFieldIndexAtDate _
        = Me.Records_OpenSQL(
            SQL:=Me.getFieldIndexSQLDate(Field, CStr(Index))
            )
        mUpdated = False
    End Function

    Friend Function getFieldIndexSQLDate(ByRef Field As String, ByRef Index As String) As String
        If Index = "" Then Index = "-1"
        getFieldIndexSQLDate = " SELECT [" & Me.Table & "].*" _
        & " From [" & Me.Table & "]" _
        & " Where ((([" & Me.Table & "]." & Field & ") = #" & Index & "#))"
    End Function

    Friend Function Records_OpenFieldIndexAt(ByRef Field As String, ByRef Index As String, Optional ByRef OrderBy As String = "") As Boolean
        Records_OpenFieldIndexAt _
        = Me.Records_OpenSQL(
            SQL:=Me.getFieldIndexSQL(Field, CStr(Index), OrderBy)
            )
        mUpdated = False
    End Function

    Friend Function Records_Available() As Boolean
        On Error GoTo HandleErr
        If (mCurrentIndex <> -1) Then
            Records_MoveNext()
        End If
        mCurrentIndex = mCurrentIndex + 1
        'IF RECORD EMPTY....
        ' mrs.RecordCount = 0...
        If Mrs.EOF = True Then
            Records_Available = False
            Exit Function
        End If
        mSubClass.getRecordset(Mrs)
        Records_Available = True
        Exit Function

HandleErr:
        Records_Available = False
        Err.Clear()
    End Function
    Public Sub getRecordset(ByRef RS As ADODB.Recordset) : End Sub
    Public Sub SetRecordSet(ByRef RS As ADODB.Recordset) : End Sub

    Friend Sub Records_MoveNext()
        On Error Resume Next
        Mrs.MoveNext()
        mSubClass.getRecordset(Mrs)
    End Sub

    Friend Sub Dispose()
        On Error Resume Next
        Mrs.Close()
        Mrs = Nothing
    End Sub

    Friend Function Record_EOF() As Boolean
        If Mrs Is Nothing Then Record_EOF = True : Exit Function
        Record_EOF = Mrs.EOF
    End Function

    Friend Sub Records_Close()
        If mUpdated Then Records_Update()
        Mrs.Close()
        Mrs = Nothing
    End Sub

    Friend Function Records_Open(Optional ByRef OrderBy As String = "", Optional ByRef ErrMsg As String = "") As Boolean
        Mrs = getRecordsetByTableLabelIndex _
        (File:=DataBase _
        , Table:=Table _
        , Label:=Index _
        , Index:="" _
        , Always:=True _
        , OrderBy:=OrderBy _
        , QuietErrors:=False _
        , ErrMsg:=ErrMsg)
        mCurrentIndex = -1
        mUpdated = False
    End Function

    Friend Property CurrentIndex() As String
        Get
            CurrentIndex = mCurrentIndex
        End Get
        Set(value As String)
            mCurrentIndex = value
        End Set
    End Property

    Friend Function RS() As ADODB.Recordset
        RS = Mrs
    End Function

    Friend Sub Records_MovePrevious()
        On Error Resume Next
        Mrs.MovePrevious()
        mSubClass.getRecordset(Mrs)
    End Sub

    Friend Sub Records_AddAndClose1()
        Records_Add()
        'Records_Close()
    End Sub

    Friend Sub Records_AddAndClose2()
        Records_Close()
    End Sub

    Friend Function Records_MoveAbsolute(ByRef Index As Integer) As Boolean
        On Error GoTo BadMove
        Mrs.AbsolutePosition = Index
        mSubClass.getRecordset(Mrs)
        Records_MoveAbsolute = True
BadMove:
    End Function

    Friend Function Value(ByRef Field As String) As String
        On Error Resume Next
        Value = Mrs(Field).Value
    End Function

    'Public Shared Widening Operator CType(v As CGrossMargin) As CDataAccess
    '    Throw New NotImplementedException()
    'End Operator
    'Public Sub getRecordset(ByRef RS As ADODB.Recordset)
    '    If FromAdvertisingType = True Then
    '        On Error Resume Next
    '        ID = RS("AdvertisingTypeID").Value
    '        AdType = IfNullThenNilString(Trim(RS("AdvertisingType").Value))
    '        OldTypeID = RS("OldTypeID").Value
    '    End If

    '    If FromEmployees = True Then
    '        On Error Resume Next
    '        ID = RS("ID")
    '        LastName = Trim(IfNullThenNilString(RS("LastName")))
    '        SalesID = Trim(IfNullThenNilString(RS("SalesID")))
    '        CommRate = Trim(IfNullThenNilString(RS("CommRate")))
    '        PassWord = Decrypt(EncryptionKey, IfNullThenNilString(RS("Pwd")))
    '        Privs = Decrypt(EncryptionKey, IfNullThenNilString(RS("Privs")))
    '        Active = RS("Active")
    '        CommTable = RS("CommTable")

    '    End If

    'End Sub

    Friend Function Record_BOF() As Boolean
        If Mrs Is Nothing Then Record_BOF = True : Exit Function
        Record_BOF = Mrs.BOF
    End Function
End Class
