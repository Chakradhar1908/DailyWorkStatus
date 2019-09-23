Module modOnDemand
    Public Function Gif89Installed() As Boolean
        'If Val(GetCDSSetting("HasGif89")) <> 0 Then Gif89Installed = True : Exit Function

        'If OnDemand_Test(OnDemandEntry("Gif89.Gif89.1")) Then
        '    SaveCDSSetting "HasGif89", "1"
        '    Gif89Installed = True
        '    Exit Function
        'End If

        Gif89Installed = False
    End Function

End Module
