Imports stdole

Module modMainMenu
    Private frmSplas As Form = frmSplash2
    Public Const frmSplashType As String = "frmSplash2"

    Public ReadOnly Property frmSplash As frmSplash2
        Get
            Return frmSplas
        End Get
    End Property

    Public ReadOnly Property frmSplashIsLoaded As Boolean
        Get
            frmSplashIsLoaded = IsFormLoaded(frmSplashType)
        End Get
    End Property

    Public Sub SetButtonImage(ByRef cmd As Button, Optional ByVal ImageIndex As Integer = -1, Optional ByVal MiniButton As Boolean = False)
        '::::SetButtonImage
        ':::SUMMARY
        ': Set image on CmdButton control
        ':::DESCRIPTION
        ': Initializes a command button image.  The button must already be set to type Graphical, as these cannot be set in code.
        ':
        '::Available Image Keys:
        ': - calc,gear,config,notes,none,calendar,cancel
        ': - rStop,rDelete,rInfo,rNext,rAdd,rPrefs,rReload,rSearch
        ': - ok,clear,map,import,print,menu,back,forward
        ': - zoom,preview,next,previous,next1,previous1,delete,plus,minus,refresh
        ': - south,west,east,north
        ':::PARAMETERS
        ': - cmd - Indicates the Command Button.
        ': - ImageName - Indicates the Image Name.
        ': - MiniButton - Indicates whether it is true or false.
        'Dim T As String
        'If cmd.Style <> vbButtonGraphical Then
        'If cmd.Image Is Nothing Then
        '    Debug.Print("Bad button")
        '    If IsDevelopment() Then
        '        Err.Raise(-1, "Development Code", "Not a graphical button: " & cmd.Name)
        '        Stop
        '    End If
        'End If
        'cmd.UseMaskColor = True
        'cmd.MaskColor = vbWhite
        'If ImageName = "" Then
        '    T = LCase(cmd.Name)
        '    If LCase(Left(T, 3)) = "cmd" Then T = Mid(T, 4)
        '    If IsIn(T, "ok", "apply", "post", "done", "close", "process") Then
        '        ImageName = "ok"
        '    ElseIf T Like "*preview" Then
        '        ImageName = "preview"
        '    ElseIf T = "cancel" Then
        '        ImageName = "cancel"
        '    ElseIf T = "clear" Then
        '        ImageName = "clear"
        '    ElseIf IsIn(T, "config", "settings", "setup", "options", "save") Then
        '        ImageName = "config"
        '    ElseIf T = "print" Then
        '        ImageName = "print"
        '    ElseIf T Like "*menu*" Then
        '        ImageName = "menu"
        '    ElseIf T Like "*next*" Then
        '        ImageName = "next"
        '    ElseIf T Like "*prev*" Then
        '        ImageName = "previous"
        '    ElseIf T Like "*del*" Then
        '        ImageName = "delete"
        '    ElseIf T Like "*calendar*" Then
        '        ImageName = "calendar"
        '    ElseIf T Like "*refresh*" Then
        '        ImageName = "refresh"
        '    ElseIf T Like "*down*" Then
        '        ImageName = "south"
        '    ElseIf T Like "*up*" Then
        '        ImageName = "north"
        '    Else
        '        ImageName = "ok"
        '    End If
        'End If

        If MiniButton Then
            'cmd.Picture = MiniButtonImage(LCase(ImageName))
            'cmd.Image = MiniButtonImage(LCase(ImageName))
            cmd.Image = MiniButtonImage(ImageIndex)
        Else
            'cmd.Picture = StandardButtonImage(LCase(ImageName))
            'cmd.Image = StandardButtonImage(LCase(ImageName))
            'cmd.Image = StandardButtonImage(ImageIndex)
            cmd.Image = MainMenu.imlStandardButtons.Images(ImageIndex)
            cmd.ImageAlign = ContentAlignment.MiddleCenter
            cmd.TextAlign = ContentAlignment.BottomCenter
            cmd.TextImageRelation = TextImageRelation.ImageAboveText
        End If
    End Sub

    Public Sub SetButtonImageSmall(ByRef cmd As Button, ByVal ImageIndex As Integer)
        cmd.Image = MainMenu.imlSmallButtons.Images(ImageIndex)
    End Sub

    Public Function MiniButtonImage(ByVal ImageName As String) As StdPicture
        '::::MiniButtonImage
        ':::SUMMARY
        ': Returns a MiniButton image
        ':::DESCRIPTION
        ': Pulls a specified image from the MiniButtonImageList control
        ':::PARAMETERS
        ': - ImageName
        ':::RETURN
        ': StdPicture

        On Error Resume Next
        ImageName = LCase(ImageName)
        'MiniButtonImage = MiniButtonImageList.ListImages(ImageName).Picture
        MiniButtonImage = MiniButtonImageList.Images(ImageName)

        If MiniButtonImage Is Nothing Then
            If IsDevelopment() Then MsgBox("Not a valid mini image name: " & ImageName, vbCritical, "Development Error")
            'MiniButtonImage = MiniButtonImageList.ListImages("none").Picture
            MiniButtonImage = MiniButtonImageList.Images("none")
        End If
    End Function
    Public Function StandardButtonImage(ByVal ImageIndex As Integer) As StdPicture
        '::::StandardButtonImage
        ':::SUMMARY
        ': Used to check whether the StandardButtonImage is Nothing or not.
        ':::DESCRIPTION
        ': This function is used to display the Standard Button Image using ImageName and to check whether the StandardButtonImage is Nothing or not and print the respective message.
        ':::PARAMETERS
        ': - ImageName - Indicates the Image Name.
        ':::RETURN
        ': StdPicture - Returns the result as StdPicture object.

        On Error Resume Next
        'ImageName = LCase(ImageName)
        'StandardButtonImage = StandardButtonImageList.ListImages(ImageName).Picture
        'StandardButtonImage = StandardButtonImageList.Images(ImageName)
        StandardButtonImage = StandardButtonImageList.Images.Item(ImageIndex)

        'If StandardButtonImage Is Nothing Then
        '    If IsDevelopment() Then MsgBox("Not a valid standard image name: " & ImageName, vbCritical, "Development Error")
        '    'StandardButtonImage = StandardButtonImageList.ListImages("none").Picture
        '    StandardButtonImage = StandardButtonImageList.Images("none")
        'End If
    End Function
    Public Function MiniButtonImageList() As ImageList
        '::::MiniButtonImageList
        ':::SUMMARY
        ': Return the ImageList of MiniButton images
        ':::DESCRIPTION
        ': Returns the ImageList control with the MiniButton images for display throughout the software.
        ':::RETURN
        ': ImageList

        MiniButtonImageList = MainMenu.imlMiniButtons
    End Function
    Public Function StandardButtonImageList() As ImageList
        '::::StandardButtonImageList
        ':::SUMMARY
        ': Returns the ImageList of the standard images
        ':::DESCRIPTION
        ': Returns the ImageList control with the StandardImages for display throughout software.
        ':::RETURN
        ': ImageList
        ':::SEE ALSO
        ': MiniButtonImageList
        StandardButtonImageList = MainMenu.imlStandardButtons
    End Function
    '    Public Property Get MainMenu() As MainMenu4
    '        '  If IsCDSComputer Then Set MainMenu = MainMenu4: Exit Property
    '        Set MainMenu = MainMenu4
    'End Property
    Public ReadOnly Property MainMenu() As MainMenu4
        Get
            '  If IsCDSComputer Then Set MainMenu = MainMenu4: Exit Property
            MainMenu = MainMenu4
        End Get
    End Property
End Module
