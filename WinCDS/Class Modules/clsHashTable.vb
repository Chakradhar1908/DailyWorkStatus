Public Class clsHashTable
    Dim slotTable() As SlotType     ' the array that holds the data
    Dim m_Count as integer             ' items in the slot table
    Dim m_HashSize as integer          ' size of hash table
    Dim hashTbl() as integer
    Dim FreeNdx as integer             ' pointer to first free slot
    Private mAutoIndex as integer      ' For AutoIndexing
    Private m_IgnoreCase As Boolean ' member variable for IgnoreCase property
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Object, Source As Object, ByVal bytes as integer)

    Private Structure SlotType
        Dim key As String
        Dim Value As Object
        Dim NextItem as integer      ' 0 if last item
    End Structure

    Public Sub Add(ByVal key As String, ByVal Value As Object)
        Dim Ndx as integer, Create As Boolean

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
    End Sub
    Public ReadOnly Property Count() as integer
        Get
            Count = m_Count
        End Get

    End Property
    Dim resx As Object
    Public ReadOnly Property Keys(Optional ByVal Sort As TriState = vbUseDefault) As Object

        Get
            Dim I as integer, J as integer, Ndx as integer
            Dim N as integer, S As String
            On Error Resume Next
            ReDim Resx(0 To m_Count - 1)

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
        Dim Ndx as integer, HCode as integer, LastNdx as integer
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
            Dim Ndx as integer

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
    Private ReadOnly Property AutoIndex() as integer
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

    Private Function GetSlotIndex(ByVal key As String, Optional ByVal Create As Boolean = False, Optional ByRef HCode as integer = 0, Optional ByRef LastNdx as integer = 0) as integer
        Dim Ndx as integer

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

    End Function
    Private Function HashCode(ByVal key As String) as integer
        Dim lastEl as integer, I as integer
        Dim codes() as integer
        ' copy ansi codes into an array of long
        lastEl = (Len(key) - 1) \ 4
        ReDim codes(lastEl)
        ' this also converts from Unicode to ANSI
        CopyMemory(codes(0), key, Len(key))

        ' XOR the ANSI codes of all characters
        For I = 0 To lastEl
            HashCode = HashCode Xor Codes(I)
        Next

    End Function
    Private Function GetFreeSlot() as integer
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
    Private Sub PrepareSlot(ByVal Index as integer, ByVal key As String,
    ByVal HCode as integer, ByVal LastNdx as integer)
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

End Class
