Module modNeoCaption
    '    Private cSK As New Collection
    '    Private Cn As cNeoCaption
    '    Private cSkin As cSkinConfiguration

    '    Public Enum NeoCaptionStyle
    '        ncNone = 0
    '        ncNeoVB
    '        ncNeoShort
    '        ncDarkMetal
    '        ncMetro
    '        ncXcursion
    '        ncBlueMedia
    '        ncBlueShort
    '        ncMacLook
    '        ncMac
    '        ncBasicTool
    '        ncBasicDialog
    '        ncBasicMsg
    '    End Enum


    '    Public Sub SetCustomFrame(ByVal F As Form, Optional ByVal Style As NeoCaptionStyle = NeoCaptionStyle.ncNone)
    '        Dim C As Control, isActive As Boolean
    '        Dim Cfg As cSkinConfiguration, NC As cNeoCaption
    '        Dim FN As String

    '        '  If Not IsDevelopment Then Exit Sub   ' disabled BFH20160329

    '        If DoCustomFrame(F) Then

    '            If fActiveForm Is F Then isActive = True
    '    Set C = F.ActiveControl

    '  On Error Resume Next
    '            FN = F.Name
    '    Set NC = cSK.Item(FN)
    '    If NC Is Nothing Then
    '                If Style = ncNone Then Exit Sub
    '      Set NC = New cNeoCaption
    '      cSK.Add NC, FN
    '    Else
    '                NC.Detach
    '            End If

    '            If Style = ncNone Then
    '                cSK.Remove FN
    '      Exit Sub
    '            End If

    '    Set Cfg = ConfigSkin(Style)
    '    If Cfg Is Nothing Then Exit Sub

    '            cSK.Remove "+" & FN
    '    cSK.Add Cfg, "+" & FN

    '    NC.Attach2 F, Cfg

    '    If isActive Then
    '                If Not (F Is fActiveForm) Then
    '                    Debug.Print "modNeoCaption - Active Form Changed on Skinning - " & F.Name
    '      End If
    '            End If

    '            On Error Resume Next
    '            '    C.SetFocus

    '        End If

    '        If DoColorForm(F) Then SetFormColor F, Style
    'End Sub

End Module
