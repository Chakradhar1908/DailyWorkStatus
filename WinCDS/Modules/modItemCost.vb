Module modItemCost
    Private mUseItemCost As Boolean, mCheckedForItemCostTable As Boolean
    Public Const Setup_UseCost_Curr As String = "Use Current"       ' update to most recent PO
    Public Const Setup_UseCost_FIFO As String = "FIFO"              ' Use Fifo Method
    Public Const Setup_UseCost_LIFO As String = "LIFO"              '
    'Public Const Setup_UseCost_NewC As String = "Use Newest Cost"
    Public Const Setup_UseCost_Aver As String = "Average Cost"
    Public Const Setup_UseCost_Manu As String = "Manual Only"

    Public Sub AddItemCost(ByVal StyleNo As String, ByVal Location as integer, ByVal Cost As Decimal, Optional ByVal Detail as integer = 0, Optional ByVal StockDate As Date = NullDate, Optional ByVal Quantity as integer = 1)
        Dim I as integer, SQL As String, dB As String
        If Not UseItemCost() Then Exit Sub              ' for transition...

        ' minimum date value..
        ' note that all auto-entries will come in as 2005, so this will denote software anomalies from auto-entry (i.e. modPatches) ones
        If DateDiff("d", #1/1/2006#, StockDate) < 0 Then StockDate = #1/1/2006#
        SQL = "INSERT INTO ItemCost (StyleNo, Location, Cost, StockDate, Detail) VALUES ('" & ProtectSQL(StyleNo) & "', " & Location & ", " & Cost & ", #" & StockDate & "#, " & Detail & ")"
        dB = GetDatabaseInventory()

        For I = 1 To Quantity
            ExecuteRecordsetBySQL(SQL, , dB)
        Next
    End Sub
    Private Function UseItemCost(Optional ByVal Recalculate As Boolean = False) As Boolean
        If Recalculate Then mCheckedForItemCostTable = False
        If Not mCheckedForItemCostTable Then mUseItemCost = TableExists(0, "ItemCost") : mCheckedForItemCostTable = True
        UseItemCost = mUseItemCost
    End Function
    Public Function GetItemCost(ByVal StyleNo As String, Optional ByVal Location as integer = 0, Optional ByVal DeleteEntry As Boolean = True, Optional ByVal Count As Double = 1, Optional ByRef DetailID as integer = 0, Optional ByVal CatalogOnly As Boolean = False) As Decimal
        Dim Cond As String, dB As String
        Dim R As ADODB.Recordset, ID as integer, User As Boolean
        Dim C As New CInvRec

        GetItemCost = 0

        If Not UseItemCost() Or CatalogOnly Then                       ' for transition...
            C.Load(StyleNo, "Style")
            GetItemCost = C.Cost * Count
            DisposeDA(C)
            Exit Function
        End If

        dB = GetDatabaseInventory()

        If Count <= 0 Then  ' BFH20060508 - Yet another clause...  have to handle negatives..
            GetItemCost = Count * GetItemCost(StyleNo, Location, False, 1)
            DisposeDA(C)
            Exit Function
        End If

        ' This segment handles the most difficult case..
        ' we aren't deleting entries, so we can't recurse, but we want more than 1..
        ' so, we have to actually examine the table entries to see what we have
        ' also, if there aren't enough, we have to approximate for the overage
        ' this, however, should fully bulletproof the routine and let us do
        ' the inventory cost report more efficiently later
        If Count > 1 And Not DeleteEntry Then
            Dim Cnt as integer
            Cond = ""
            Cond = Cond & "SELECT"
            Cond = Cond & " TOP " & Count & " *"
            Cond = Cond & " FROM [ItemCost]"
            Cond = Cond & " WHERE StyleNo='" & ProtectSQL(StyleNo) & "'"
            If Location <> 0 Then Cond = Cond & " AND Location=" & Location
            Cond = Cond & " ORDER BY [StockDate] "

            If IsIn(UseCost, Setup_UseCost_Aver, Setup_UseCost_Curr, Setup_UseCost_Manu) Then
                C.Load(StyleNo, "Style")
                GetItemCost = C.Cost * Count
                DisposeDA(C)
                Exit Function
            ElseIf StoreSettings.UseCost = Setup_UseCost_FIFO Then
                R = GetRecordsetBySQL(Cond & "ASC", , dB)
            ElseIf StoreSettings.UseCost = Setup_UseCost_LIFO Then
                R = GetRecordsetBySQL(Cond & "DESC", , dB)
            Else
                DisposeDA(C, R)
                Exit Function ' ??
            End If
            Do While Not R.EOF
                GetItemCost = GetItemCost + R("Cost").Value
                Cnt = Cnt + 1
                R.MoveNext
            Loop
            DisposeDA(R)

            If Cnt < Count Then
                If C.Load(StyleNo, "Style") Then
                    GetItemCost = GetItemCost + C.Cost * (Count - Cnt)
                Else
                    If Cnt > 0 Then ' we tried to get the standard cost, but if that fails, we get the average of what was listed
                        GetItemCost = GetItemCost + (GetItemCost / Cnt) * (Count - Cnt)
                    Else
                        DisposeDA(C, R)
                        Exit Function ' failing all else, we protect this function and return 0
                    End If
                End If
                DisposeDA(C)
            End If

            Exit Function
        End If
        ' If we can delete entries, we'll just use recursion
        If Count > 1 Then
            For ID = 1 To Count
                GetItemCost = GetItemCost + GetItemCost(StyleNo, Location, DeleteEntry, 1)
            Next
            Exit Function
        End If

        C.Load(StyleNo, "Style")

        Cond = ""
        Cond = Cond & "SELECT"
        Cond = Cond & " TOP 1 *"
        Cond = Cond & " FROM [ItemCost]"
        Cond = Cond & " WHERE StyleNo='" & ProtectSQL(StyleNo) & "'"
        If Location <> 0 Then Cond = Cond & " AND Location=" & Location
        Cond = Cond & " ORDER BY [StockDate] "

        Select Case StoreSettings.UseCost
            Case Setup_UseCost_FIFO
                R = GetRecordsetBySQL(Cond & "ASC", , dB)
            Case Setup_UseCost_LIFO
                R = GetRecordsetBySQL(Cond & "DESC", , dB)
            Case Setup_UseCost_Aver
            Case Setup_UseCost_Curr
            Case Setup_UseCost_Manu
        End Select

        User = True
        If R Is Nothing Then
            User = False
        Else
            If R.EOF Then User = False
        End If

        If User Then
            GetItemCost = R("Cost").Value * Count 'still need count to handle "0.5"
            ID = R("ItemCostID").Value
            DetailID = R("Detail").Value
            If DeleteEntry Then DeleteItemCostByID(ID)
        Else
            GetItemCost = C.Cost * Count  'still need count to handle "0.5"
            If DeleteEntry Then
                R = GetRecordsetBySQL(Cond & "ASC", , dB)
                If Not R.EOF Then
                    DeleteItemCostByID(R("ItemCostID").Value)
                End If
                R = Nothing
            End If
        End If

        On Error Resume Next
        DisposeDA(C, R)
    End Function

    Public Function DeleteItemCostByID(ByVal ID as integer) As Boolean
        If Not UseItemCost() Then Exit Function
        ExecuteRecordsetBySQL("DELETE FROM [ItemCost] WHERE ItemCostID=" & ID, , GetDatabaseInventory)
        DeleteItemCostByID = True
    End Function

    Public Function UseCost() As String
        UseCost = StoreSettings.UseCost
    End Function
End Module