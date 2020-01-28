Public Class clsHashTable
    Dim slotTable() As SlotType     ' the array that holds the data
    Dim m_Count as integer             ' items in the slot table
    Dim m_HashSize as integer          ' size of hash table
    Dim hashTbl() as integer
    Dim FreeNdx as integer             ' pointer to first free slot
    Private mAutoIndex as integer      ' For AutoIndexing
    Private m_IgnoreCase As Boolean ' member variable for IgnoreCase property
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Object, Source As Object, ByVal bytes as integer)
    Dim resx As Object
    Dim m_ListSize As Integer          ' size of slot table
    Dim m_ChunkSize As Integer         ' chunk size

    Private Structure SlotType
        Dim key As String
        Dim Value As Object
        Dim NextItem As Integer      ' 0 if last item
    End Structure

    Public Sub Add(ByVal key As String, ByVal Value As Object)
        Dim Ndx As Integer, Create As Boolean

        Try
            If key = "" Then key = AutoIndex

            ' get the index to the slot where the value is
            ' (allocate a new slot if necessary)
            Create = True
            Ndx = GetSlotIndex(key, Create)

            If Create Then
                ' the item was actually added
                'If IsObject(Value) Then
                If Not Value Is Nothing Then
                    slotTable(Ndx).Value = Value
                Else
                    slotTable(Ndx).Value = Value
                End If
            Else
                ' raise error "This key is already associated with an item of this
                ' collection"
                Err.Raise(457)
            End If

        Catch ex As NullReferenceException
            Exit Sub
        End Try
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Count = m_Count
        End Get
    End Property

    Public ReadOnly Property Keys(Optional ByVal Sort As TriState = vbUseDefault) As Object
        Get
            Dim I As Integer, J As Integer, Ndx As Integer
            Dim N As Integer, S As String
            On Error Resume Next
            ReDim resx(0 To m_Count - 1)

            For I = 0 To m_HashSize - 1
                ' take the pointer from the hash table
                Ndx = hashTbl(I)
                ' walk the slottable() array
                Do While Ndx
                    resx(N) = slotTable(Ndx).key
                    N = N + 1
                    Ndx = slotTable(Ndx).NextItem
                Loop
            Next

            If Sort <> vbUseDefault Then
                For I = 0 To m_Count - 2
                    For J = I + 1 To m_Count - 1
                        If Sort = vbTrue Then
                            If StrComp(resx(I), resx(J), vbTextCompare) > 0 Then
                                S = resx(J)
                                resx(J) = resx(I)
                                resx(I) = S
                            End If
                        Else
                            If StrComp(resx(I), resx(J), vbTextCompare) < 0 Then
                                S = resx(J)
                                resx(J) = resx(I)
                                resx(I) = S
                            End If
                        End If
                    Next
                Next
            End If

            ' assign to the result
            Keys = resx()
        End Get
    End Property

    Public Sub Remove(ByVal key As String)
        Dim Ndx As Integer, HCode As Integer, LastNdx As Integer
        Ndx = GetSlotIndex(key, False, HCode, LastNdx)
        ' raise error if no such element
        If Ndx = 0 Then Err.Raise(5)

        If LastNdx Then
            ' this isn't the first item in the slotTable() array
            slotTable(LastNdx).NextItem = slotTable(Ndx).NextItem
        ElseIf slotTable(Ndx).NextItem Then
            ' this is the first item in the slotTable() array
            ' and is followed by one or more items
            hashTbl(HCode) = slotTable(Ndx).NextItem
        Else
            ' this is the only item in the slotTable() array
            ' for this hash code
            hashTbl(HCode) = 0
        End If

        ' put the element back in the free list
        slotTable(Ndx).NextItem = FreeNdx
        FreeNdx = Ndx
        ' we have deleted an item
        m_Count = m_Count - 1

    End Sub

    Public ReadOnly Property Item(ByVal key As String) As Object
        Get
            Dim Ndx As Integer

            Item = Nothing
            ' get the index to the slot where the value is
            Ndx = GetSlotIndex(key)
            If Ndx = 0 Then
                ' return Empty if not found
                'ElseIf IsObject(slotTable(Ndx).Value) Then
            ElseIf Not slotTable(Ndx).Value Is Nothing Then
                Item = slotTable(Ndx).Value
            Else
                Item = slotTable(Ndx).Value
            End If
        End Get
    End Property

    Private ReadOnly Property AutoIndex() As Integer
        Get
            Do While Exists(mAutoIndex)
                mAutoIndex = mAutoIndex + 1
            Loop
            AutoIndex = mAutoIndex
        End Get
    End Property

    Public Function Exists(ByVal key As String) As Boolean
        Exists = GetSlotIndex(key) <> 0
    End Function

    Private Function GetSlotIndex(ByVal key As String, Optional ByVal Create As Boolean = False, Optional ByRef HCode As Integer = 0, Optional ByRef LastNdx As Integer = 0) As Integer
        Dim Ndx As Integer

        Try
            ' raise error if invalid key
            If Len(key) = 0 Then Err.Raise(1001, , "Invalid key")

            ' keep case-unsensitiveness into account
            If m_IgnoreCase Then key = UCase(key)
            ' get the index in the hashTbl() array
            HCode = HashCode(key) Mod m_HashSize
            ' get the pointer to the slotTable() array
            Ndx = hashTbl(HCode)

            ' exit if there is no item with that hash code
            Do While Ndx
                ' compare key with actual value
                If slotTable(Ndx).key = key Then Exit Do
                ' remember last pointer
                LastNdx = Ndx
                ' check the next item
                Ndx = slotTable(Ndx).NextItem
            Loop

            ' create a new item if not there
            If Ndx = 0 And Create Then
                Ndx = GetFreeSlot()
                PrepareSlot(Ndx, key, HCode, LastNdx)
            Else
                ' signal that no item has been created
                Create = False
            End If
            ' this is the return value
            GetSlotIndex = Ndx

        Catch ex As DivideByZeroException
            Exit Function
        End Try
    End Function

    Private Function HashCode(ByVal key As String) As Integer
        Dim lastEl As Integer, I As Integer
        Dim codes() As Integer
        ' copy ansi codes into an array of long
        lastEl = (Len(key) - 1) \ 4
        ReDim codes(lastEl)
        ' this also converts from Unicode to ANSI
        CopyMemory(codes(0), key, Len(key))

        ' XOR the ANSI codes of all characters
        For I = 0 To lastEl
            HashCode = HashCode Xor codes(I)
        Next

    End Function

    Private Function GetFreeSlot() As Integer
        ' allocate new memory if necessary
        'If FreeNdx = 0 Then (ExpandSlotTable m_ChunkSize)
        ' use the first slot
        GetFreeSlot = FreeNdx
        ' update the pointer to the first slot
        FreeNdx = slotTable(GetFreeSlot).NextItem
        ' signal this as the end of the linked list
        slotTable(GetFreeSlot).NextItem = 0
        ' we have one more item
        m_Count = m_Count + 1
    End Function

    Private Sub PrepareSlot(ByVal Index As Integer, ByVal key As String, ByVal HCode As Integer, ByVal LastNdx As Integer)
        ' assign the key
        ' keep case-sensitiveness into account
        If m_IgnoreCase Then key = UCase(key)
        slotTable(Index).key = key

        If LastNdx Then
            ' this is the successor of another slot
            slotTable(LastNdx).NextItem = Index
        Else
            ' this is the first slot for a given hash code
            hashTbl(HCode) = Index
        End If
    End Sub

    Public Function ContentString(Optional ByVal Separator As String = vbCrLf) As String
        Dim x As String, L As Object
        On Error GoTo NoVars
        For Each L In Keys(vbTrue)
            x = x & L & "=" & Item(L) & Separator
        Next
        ContentString = x
        Exit Function
NoVars:
        ContentString = "No Variables."
    End Function

    Public Function LoadQueryString(ByVal Q As String) As Boolean
        Dim C As clsHashTable, L As Object
        C = QueryStringParse(Q)

        RemoveAll()

        If C.Count > 0 Then
            For Each L In C.Keys
                Add(L, C.Item(L))
            Next
        End If

        LoadQueryString = True
    End Function

    Public Sub RemoveAll()
        SetSize(m_HashSize, m_ListSize, m_ChunkSize)
    End Sub

    Public Sub SetSize(ByVal HashSize As Integer, Optional ByVal ListSize As Integer = 0, Optional ByVal ChunkSize As Integer = 0)
        ' provide defaults
        If ListSize <= 0 Then ListSize = m_ListSize
        If ChunkSize <= 0 Then ChunkSize = m_ChunkSize
        ' save size values
        m_HashSize = HashSize
        m_ListSize = ListSize
        m_ChunkSize = ChunkSize
        m_Count = 0
        ' rebuild tables
        FreeNdx = 0
        ReDim hashTbl(0 To HashSize - 1)
        ReDim slotTable(0)
        ExpandSlotTable(m_ListSize)
    End Sub

    Private Sub ExpandSlotTable(ByVal numEls As Integer)
        Dim newFreeNdx As Integer, I As Integer
        newFreeNdx = UBound(slotTable) + 1

        ReDim Preserve slotTable(0 To UBound(slotTable) + numEls)
        ' create the linked list of free items
        For I = newFreeNdx To UBound(slotTable)
            slotTable(I).NextItem = I + 1
        Next
        ' overwrite the last (wrong) value
        slotTable(UBound(slotTable)).NextItem = FreeNdx
        ' we now know where to pick the first free item
        FreeNdx = newFreeNdx
    End Sub

End Class
