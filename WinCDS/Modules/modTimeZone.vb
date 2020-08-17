Module modTimeZone
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
    ' Some pages may also contain other copyrights by the author.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Distribution: You can freely use this code in your own
    '               applications, but you may not reproduce
    '               or publish this code on any web site,
    '               online service, or distribute as source
    '               on any media without express permission.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'http://vbnet.mvps.org/index.html?code/locale/gettimezonebias.htm
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Const TIME_ZONE_ID_UNKNOWN As Integer = 1
    Private Const TIME_ZONE_ID_STANDARD As Integer = 1
    Private Const TIME_ZONE_ID_DAYLIGHT As Integer = 2
    Private Const TIME_ZONE_ID_INVALID As Integer = &HFFFFFFFF

    Private Structure SYSTEMTIME
        Dim wYear As Integer
        Dim wMonth As Integer
        Dim wDayOfWeek As Integer
        Dim wDay As Integer
        Dim wHour As Integer
        Dim wMinute As Integer
        Dim wSecond As Integer
        Dim wMilliseconds As Integer
    End Structure

    Private Structure TIME_ZONE_INFORMATION
        Dim Bias As Integer
        'Dim StandardName(0 To 63) As Byte  'unicode (0-based)
        Dim StandardName() As Byte  'unicode (0-based)
        Dim StandardDate As SYSTEMTIME
        Dim StandardBias As Integer
        'Dim DaylightName(0 To 63) As Byte  'unicode (0-based)
        Dim DaylightName() As Byte  'unicode (0-based)
        Dim DaylightDate As SYSTEMTIME
        Dim DaylightBias As Integer
    End Structure
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Integer

    Public Function GetCurrentTimeZoneOffset() As String
        Dim Tzi As TIME_ZONE_INFORMATION
        Dim dwBias As Integer
        Dim Tmp As String

        Select Case GetTimeZoneInformation(Tzi)
            Case TIME_ZONE_ID_DAYLIGHT
                dwBias = Tzi.Bias + Tzi.DaylightBias
            Case Else
                dwBias = Tzi.Bias + Tzi.StandardBias
        End Select

        dwBias = -dwBias
        If dwBias < 0 Then Tmp = "" Else Tmp = "+"
        Tmp = Tmp & Format(dwBias \ 60, "00")
        Tmp = Tmp & Format(dwBias Mod 60, "00")

        GetCurrentTimeZoneOffset = Tmp
    End Function

End Module
