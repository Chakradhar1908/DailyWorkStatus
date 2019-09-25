Module modArrays
    Public Enum IsInOptions
        isinNone = 0
        isinIgnoreCase = 1
        isinLeftCheck = 2
    End Enum

    'BFH20050906
    ' aka member_array(), et al
    Public Function IsIn(ByVal What, ParamArray Rest()) As Boolean
        Dim L
        For Each L In Rest
            If TypeName(What) = "String" Or TypeName(L) = "String" Then
                If "" & What = "" & L Then IsIn = True : Exit Function
            Else
                If What = L Then IsIn = True : Exit Function
            End If
        Next
        IsIn = False
    End Function
    Public Function FitList(ByVal What, ByRef Arr, Optional ByVal Dflt = "#")
        Dim A()
        A = Arr
        If IsInArray(What, A) Then
            FitList = What
        Else
            If Dflt = "#" Then
                FitList = A(LBound(A))
            Else
                FitList = Dflt
            End If
        End If
    End Function
    Public Function IsInArray(ByVal What, ByRef Arr()) As Boolean
        Dim L
        On Error GoTo NoArray
        For Each L In Arr
            If What = L Then IsInArray = True : Exit Function
        Next
NoArray:
        IsInArray = False
    End Function
    Public Function SubArr(ByVal sourceArray As Object, ByVal fromIndex As Integer, ByVal copyLength As Integer)
        SubArr = ArrSlice(sourceArray, fromIndex, fromIndex + copyLength - 1)
    End Function
    Public Function ArrSlice(ByRef sourceArray As Object, ByVal fromIndex As Integer, ByVal toIndex As Integer)
        Dim Idx As Integer
        Dim tempList() = Nothing

        ArrSlice = Nothing
        If Not IsArray(sourceArray) Then Exit Function

        fromIndex = FitRange(LBound(sourceArray), fromIndex, UBound(sourceArray))
        toIndex = FitRange(fromIndex, toIndex, UBound(sourceArray))

        For Idx = fromIndex To toIndex
            ArrAdd(tempList, sourceArray(Idx))
        Next

        ArrSlice = tempList
    End Function
    Public Sub ArrAdd(ByRef Ar(), ByRef Item)
        Dim X As Integer
        Dim Arr() As Object

        Err.Clear()
        On Error Resume Next
        X = UBound(Ar)
        If Err.Number <> 0 Then
            'Arr = Array(Item)
            ReDim Preserve Arr(Item)

            Exit Sub
        End If
        ReDim Preserve Arr(UBound(Ar) + 1)
        Arr(UBound(Ar)) = Item
    End Sub
    Public Function IsNotIn(ByVal What, ParamArray Rest()) As Boolean
        Dim A()
        A = Rest
        IsNotIn = Not IsInArray(What, A)
    End Function
    Public Function NotIsIn(ByVal What, ParamArray Rest()) As Boolean
        Dim A()
        A = Rest
        NotIsIn = Not IsInArray(What, A)
    End Function
    Public Function IsNotInArray(ByVal What, ByRef Rest()) As Boolean
        Dim A()
        A = Rest
        IsNotInArray = Not IsInArray(What, A)
    End Function
    Public Function NotIsInArray(ByVal What, ByRef Rest()) As Boolean
        Dim A()
        A = Rest
        NotIsInArray = Not IsInArray(What, A)
    End Function
    Public Sub AddToArray(ByRef Arr, ByRef El)
        If Not IsArray(Arr) Then ReDim Arr(0) Else ReDim Preserve Arr(UBound(Arr) + 1)
        Arr(UBound(Arr)) = El
    End Sub

    Public Function IsInOpt(ByVal Src As String, ByVal Options As IsInOptions, ParamArray Rest() As Object) As Boolean
        Dim L, A As String, B As String
        For Each L In Rest
            A = Src
            B = L
            If Options And IsInOptions.isinIgnoreCase Then
                A = UCase(A)
                B = UCase(B)
            End If
            If Options And IsInOptions.isinLeftCheck Then
                A = Left(A, Min(Len(A), Len(B)))
                B = Left(B, Min(Len(A), Len(B)))
            End If
            If "" & A = "" & B Then
                IsInOpt = True
                Exit Function
            End If
        Next
    End Function

End Module
