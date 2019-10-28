Public Class CDbTypeAhead
    Private mRS As ADODB.Recordset
    Private mTable As String
    Private mField As String
    Private mOperator As String
    Private mValue As String
    Private mMatch As String
    Private mMinLen As Integer
    Private OrigSearch As String
    Private LastSearch As String
    Private mOrderBy As String
    Private mExtraCondition As String
    Private mExtraSort As String
    Public mUpdated As Boolean

    Public Event Refresh()
    Public Event BuildLine(RS As ADODB.Recordset, returnLine As String)
    Public Event BuildKeyedLine(RS As ADODB.Recordset, returnLine As String, returnKey As Integer)

    Public Sub New()
        Table = ""
        Field = ""
        Operators = "LIKE"
        Value = ""
        Match = 1
        MinLength = 0
        ExtraCondition = ""
        ExtraSort = ""
        mUpdated = False
        mRS = Nothing
    End Sub

    Public Property Operators() As String '->This property name is "Operator" in vb6. Changed it to "Operators" here. Because "Operator" is a keyword in vb.net.
        Get
            Operators = mOperator
        End Get
        Set(value As String)
            mOperator = value
        End Set
    End Property

    Public Property Table As String
        Get
            Table = mTable
        End Get
        Set(value As String)
            mTable = value
        End Set
    End Property

    Public Property Field As String
        Get
            Field = mField
        End Get
        Set(value As String)
            mField = value
        End Set
    End Property

    Public Property Value As String
        Get
            Value = mValue
        End Get
        Set(value As String)
            mValue = value
        End Set
    End Property

    Public Property Match As Integer
        Get
            Match = mMatch
        End Get
        Set(value As Integer)
            mMatch = value
        End Set
    End Property

    Public Property MinLength As Integer
        Get
            MinLength = mMinLen
        End Get
        Set(value As Integer)
            mMinLen = value
        End Set
    End Property

    Public Property ExtraCondition() As String
        Get
            ExtraCondition = mExtraCondition
        End Get
        Set(value As String)
            mExtraCondition = value
        End Set
    End Property

    Public Property ExtraSort() As String
        Get
            ExtraSort = mExtraSort
        End Get
        Set(value As String)
            mExtraSort = value
        End Set
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            If (mRS Is Nothing) Then
                Count = -1
            Else
                Count = mRS.RecordCount
            End If
        End Get
    End Property

    Public Sub RefreshObject()
        RaiseEvent Refresh()
    End Sub

    Public Function ListToListBox(ByRef Find As String, ByRef lstBox As ListBox) As Boolean
        Dim RS As ADODB.Recordset
        Dim RET As String, vKEY As Integer
        Dim El As Object

        RS = getRecordset(Find)
        lstBox.Items.Clear()

        If Not (RS Is Nothing) Then
            'LockWindowUpdate lstBox.hwnd
            LockWindowUpdate(lstBox.Handle)
            Do Until RS.EOF
                RET = ""
                vKEY = 0

                RaiseEvent BuildKeyedLine(RS, RET, vKEY)

                If RET <> "(skip)" Then
                    If (RET = "") Then RET = RS(mField).Value
                    For Each El In Split(RET, vbCrLf)
                        'lstBox.AddItem El 'ret -> These two commented lines are replaced with the third line.
                        'lstBox.itemData(lstBox.NewIndex) = vKEY
                        lstBox.Items.Add(New ItemDataClass(El, vKEY))
                    Next
                End If
                RS.MoveNext()
            Loop
            'LockWindowUpdate 0
            LockWindowUpdate(IntPtr.Zero)
            On Error Resume Next
            'lstBox.Selected(0) = True
            lstBox.SelectedIndex = 0
            ListToListBox = True
        End If
    End Function

    Public Function getRecordset(Optional ByRef Value As String = "") As ADODB.Recordset
        Dim WC As String
        If MinLength < 0 Then MinLength = 0
        ' We're trying to get a matching recordset.
        If Len(Value) < MinLength Then
            ' The search string is too short, clear mrs.
            mRS = Nothing
        Else
            ' If the base string is different than the original query string, requery.
            ' Otherwise filter.

            ' We know the string is long enough to search on!
            ' Now we need to check: if there's no recordset or the base is different, requery.
            ' mValue is the original (or last) search?  We need to clarify.

            ' OrigSearch is the Query text.
            ' LastSearch is the current filter.
            ' Value is what we want.

            WC = "%"
            If InStr(1, Value, OrigSearch) <> 1 Or mRS Is Nothing Then
                ' Value is based on the original search.
                mValue = Value
                OrigSearch = mValue
                mRS = GetRecordsetBySQL("SELECT * FROM " & mTable & " WHERE [" &
        Field & "] LIKE """ & ProtectSQL(mValue) & "" & WC & """" &
        IIf(mExtraCondition = "", "", " AND " & mExtraCondition) &
        " ORDER BY [" & mField & "]" & IIf(Len(ExtraSort) > 0, "," & ExtraSort, ""), , GetDatabaseAtLocation)
            Else
                mValue = Value
                LastSearch = mValue
                mRS.Filter = "[" & mField & "] " & mOperator & " '" & Replace(mValue, "'", "''") & "" & WC & "'"
            End If
        End If
        getRecordset = mRS


        '  If (Value <> "") Or (mMatch = 0) Then
        '    Debug.Print Value, mValue, InStr(1, Value, mValue)
        '    If (mid(Value, 1, mMatch) <> mid(mValue, 1, mMatch)) _
        '    Or (mMatch = 0) Then
        '      Set mrs = Nothing
        '    End If
        '    mValue = Value
        '  End If
        '  If ((Value = "") Or (mUpdated = False) Or _
        '    ((Len(Value) >= mMatch) And ((mrs Is Nothing) Or (Value = "")))) Then
        '    mUpdated = True
        '    Debug.Print "Get new Recordset"
        '    Set mrs = getRecordsetByTable( _
        '        Table:=mTable _
        '        , field:=mField _
        '        , Operator:=mOperator _
        '        , Value:=mValue & "%")
        '  Else
        '    If Not (mrs Is Nothing) Then
        '        mrs.Filter = mField & " " & mOperator & " '" & mValue & "%' "
        '    End If
        '  End If
        '  Set getRecordset = mrs
    End Function

    Public Function List(ByRef Find As String) As Object
        Dim T As String : T = "-->"
        Dim RS As ADODB.Recordset

        RS = getRecordset(Find)
        If Not (RS Is Nothing) Then
            Do Until RS.EOF
                Dim RET As String : RET = ""
                RaiseEvent BuildLine(RS, RET)
                If RET = "-" Then
                    ' This item shouldn't add to recordcount..  but it's a very minor bug.
                    'Count = Count - 1
                Else
                    If (RET = "") Then RET = RS(mField).Value
                    T = T & RET & vbCrLf
                End If
                RS.MoveNext()
            Loop
        End If
        List = T
    End Function

    Public Function Locate(ByRef Find As Integer) As ADODB.Recordset
        If Not (mRS Is Nothing) Then
            If mRS.RecordCount > 0 Then
                mRS.AbsolutePosition = Find
            End If
        End If
        Locate = mRS
        mUpdated = False
    End Function

End Class
