Module modHolidays
    Public Function BlackFridayDate(Optional ByVal Yr As Integer = -1) As Date
        '::::BlackFridayDate
        ':::SUMMARY
        ': Used to return Black Friday Date.
        ':::DESCRIPTION
        ': This function is used to display Black Friday Date by using DateAdd function.
        ':::PARAMETERS
        ': - Yr - Indicates the Year.
        ':::RETURN
        ': Date - Return Black Friday Date.
        If Yr < 0 Then Yr = Year(Today)
        BlackFridayDate = DateAdd("d", 1, ThanksgivingDate(Yr))
    End Function

    Public Function ThanksgivingDate(Optional ByVal Yr As Integer = -1) As Date ' 4th Thu in Nov
        '::::ThanksgivingDate
        ':::SUMMARY
        ': Used to return Thanksgiving Date.
        ':::DESCRIPTION
        ': This function is used to display Thanksgiving Date by using NDow function.
        ':::PARAMETERS
        ': - Yr - Indicates the Year.
        ':::RETURN
        ': Date - Return Thanksgiving Date.
        If Yr < 0 Then Yr = Year(Today)
        ThanksgivingDate = NDow(Yr, 11, 4, vbThursday)
        '  ThanksgivingDate = DateSerial(Yr, 11, 29 - Weekday(DateSerial(Yr, 11, 1), vbFriday))
    End Function

    Public Function NDow(ByVal Yr As Integer, ByVal Mo As Integer, ByVal Nth As Integer, ByVal DOW As VBA.VbDayOfWeek) As Date
        '::::NDow
        ':::SUMMARY
        ': Used to return Nth day of Week.
        ':::DESCRIPTION
        ': Could be used for "ThanksgivingDate" above.
        ':::PARAMETERS
        ': - Yr - Indicates the Year.
        ': - Mo - Indicates the Month.
        ': - Nth - Indicates the Nth Day of week.
        ': - DOW - Indicates the Day of Week.
        ':::RETURN
        ': Date - Return Nth day of Week.

        ' Nth day of Week
        ' Could be used for "ThanksgivingDate" above
        NDow = DateSerial(Yr, Mo, (8 - Weekday(DateSerial(Yr, Mo, 1), (DOW + 1) Mod 8)) + ((Nth - 1) * 7))
    End Function

End Module
