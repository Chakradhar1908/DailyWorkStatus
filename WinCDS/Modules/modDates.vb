Module modDates
    Public Const NullDate As Date = #1/1/2001#
    Public Const NullDateString As String = "1/1/2001"
    Public Function DateInRange(ByVal TestDate As String, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
        DateInRange = False
        If Not IsDate(TestDate) Then Exit Function
        If DateDiff("d", CDate(TestDate), StartDate) > 0 Then Exit Function
        If DateDiff("d", CDate(TestDate), EndDate) < 0 Then Exit Function
        DateInRange = True
    End Function
    ' answers the question, 'Is the date <check> after <against>?'
    ' be sure to adjust IncludeBound and Unit accordingly

    'Public Function DateAfter(ByVal Check As Date, ByVal Against As Date, Optional ByVal IncludeBound As Boolean = True, Optional ByVal Unit As String = "d") As Boolean
    Public Function DateAfter(ByVal Check As Date, ByVal Against As Date, Optional ByVal IncludeBound As Boolean = True, Optional ByVal Unit As DateInterval = DateInterval.Day) As Boolean
        Dim R As Integer
        R = DateDiff(Unit, Check, Against)
        DateAfter = IIf(IncludeBound, R <= 0, R < 0)
    End Function
    Public Function DateEqual(ByVal Check1 As Date, ByVal Check2 As Date, Optional ByVal Unit As String = "d") As Boolean
        DateEqual = DateDiff(Unit, Check1, Check2) = 0
    End Function
    Public Function DateFormat(ByVal dteDate As Object, Optional ByVal Separator As String = "") As String
        Dim S As String
        S = DateFormatString
        If Separator <> "" Then S = Replace(S, "/", Separator)

        If IsNothing(dteDate) Then dteDate = ""
        If Not IsDate(dteDate) Then
            DateFormat = Space(Len(S))
        Else
            DateFormat = Format(dteDate, S)
        End If
        'DateFormat = "05/14/2019"
    End Function
    Public Function DateFormatString() As String
        DateFormatString = "MM/dd/yyyy"
    End Function
    Public Function GetDay(ByVal dteDay As Date) As String
        GetDay = Nothing
        'If dteDay = 0 Then Exit Function
        If Len(dteDay) = 1 Then Exit Function
        GetDay = UCase(Format(dteDay, "DDD"))
    End Function
    Public Function TimeFormat(ByVal dteDate As Object) As String
        Dim S As String
        S = TimeFormatString()

        If IsNothing(dteDate) Then dteDate = ""

        If Not IsDate(dteDate) Then
            TimeFormat = Space(Len(S))
        Else
            TimeFormat = Format(dteDate, S)
        End If
    End Function

    Public Function TimeFormatString() As String
        TimeFormatString = "hh:mm ampm"
    End Function
    Public Function DateTimeStamp(Optional ByVal D As Date = Nothing) As String
        If CLng(D.ToString) = 0 Then D = Now
        DateTimeStamp = Format(D, "YYYYMMDDHHmm")
    End Function
    Public Function DateStampFile(ByVal S As String, Optional ByVal DateAndTime As Boolean = False) As String
        DateStampFile = Replace(S, "$", IIf(DateAndTime, DateTimeStamp, DateStamp))
    End Function
    Public Function DateStamp(Optional ByVal D As Date = Nothing) As String
        If CLng(D.ToString) = 0 Then D = Now
        DateStamp = Format(D, "YYYYMMDD")
    End Function
    Public Function DateBefore(ByVal Check As Date, ByVal Against As Date, Optional ByVal IncludeBound As Boolean = True, Optional ByVal Unit As String = "d") As Boolean
        Dim R as integer
        R = DateDiff(Unit, Check, Against)
        DateBefore = IIf(IncludeBound, R >= 0, R > 0)
    End Function

    Public Function DateBetween(ByVal Check As Date, ByVal Lower As Date, ByVal Upper As Date, Optional ByVal IncludeBound As Boolean = True, Optional ByVal Unit As String = "d") As Boolean
        DateBetween = DateAfter(Check, Lower, IncludeBound, Unit) And DateBefore(Check, Upper, IncludeBound, Unit)
    End Function

    Public Function YearStart(Optional ByVal D As String = "", Optional ByVal YearOffset As Integer = 0) As Date
        If Not IsDate(D) Then D = Today
        YearStart = DateValue("01/01/" & Year(D))
        YearStart = DateAdd("yyyy", YearOffset, YearStart)
    End Function

    Public Function DaySeek(ByVal mDate As Date, ByVal SeekDay As Integer, Optional ByVal DirForward As Boolean = True) As Date
        DaySeek = mDate
        Do Until DateAndTime.Day(DaySeek) = SeekDay
            DaySeek = IIf(DirForward, DayAfter(DaySeek), DayBefore(DaySeek))
        Loop
    End Function

    Public Function DayAfter(ByVal mDate As Date, Optional ByVal Value As Integer = 1) As Date
        DayAfter = DayAdd(mDate, Value)
    End Function
    Public Function DayBefore(ByVal mDate As Date, Optional ByVal Value As Integer = 1) As Date
        DayBefore = DayAdd(mDate, -Value)
    End Function

    Public Function DayAdd(ByVal mDate As Date, ByVal Value As Integer) As Date
        DayAdd = DateAdd("d", Value, mDate)
    End Function

    Public Function CheckNullDate(ByRef ValDate As Date, Optional ByVal NewDate As Date = NullDate, Optional ByVal wTime As Boolean = False) As Date
        If DateEqual(ValDate, NullDate) Then
            If DateEqual(NewDate, NullDate) Then NewDate = Now
            ValDate = NewDate
        End If

        If Not wTime Then ValDate = DateValue(ValDate)
        CheckNullDate = ValDate
    End Function

    Public Function MonthAdd(ByVal mDate As Date, ByVal Value As Integer) As Date
        MonthAdd = DateAdd("m", Value, mDate)
    End Function

    Public Function MonthAfter(ByVal mDate As Date, Optional ByVal nMonths As Integer = 1) As Date
        MonthAfter = MonthAdd(mDate, nMonths)
    End Function

    Public Function Age(ByVal BDay As Date, Optional ByVal When1 As Date = #1/1/1901#) As Integer
        Dim Y As Integer, Before As Boolean, T As String, Adj As Integer

        If When1 = #1/1/1901# Then When1 = Today
        Y = DateDiff("yyyy", BDay, When1)
        T = Format(BDay, "mm/dd/" & Format(When1, "yyyy"))
        Do While Not IsDate(T)    ' leap year fails...  "2/29/" + "2010" isn't a valid date
            Adj = Adj + 1
            T = Format(BDay, "mm/dd/" & Format(DateAdd("yyyy", Adj, When1), "yyyy"))
            Adj = Adj + 1
        Loop
        If DateDiff("d", T, When1) < 0 Then Y = Y - 1
        Age = Age - Adj
        Age = IIf(Y >= 0, Y, 0)
    End Function

End Module
