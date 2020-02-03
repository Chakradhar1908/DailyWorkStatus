Module modInitCommonControls
    Public Sub DoInitCommonControls()
        Dim icCex As InitCommonControlsExStruct, hMod As Long
        ' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507&#37;28VS.85%29.aspx
        Const ICC_ANIMATE_CLASS As Long = &H80&
        Const ICC_BAR_CLASSES As Long = &H4&
        Const ICC_COOL_CLASSES As Long = &H400&
        Const ICC_DATE_CLASSES As Long = &H100&
        Const ICC_HOTKEY_CLASS As Long = &H40&
        Const ICC_INTERNET_CLASSES As Long = &H800&
        Const ICC_LINK_CLASS As Long = &H8000&
        Const ICC_LISTVIEW_CLASSES As Long = &H1&
        Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&
        Const ICC_PAGESCROLLER_CLASS As Long = &H1000&
        Const ICC_PROGRESS_CLASS As Long = &H20&
        Const ICC_TAB_CLASSES As Long = &H8&
        Const ICC_TREEVIEW_CLASSES As Long = &H2&
        Const ICC_UPDOWN_CLASS As Long = &H10&
        Const ICC_USEREX_CLASSES As Long = &H200&
        Const ICC_STANDARD_CLASSES As Long = &H4000&
        Const ICC_WIN95_CLASSES As Long = &HFF&
        Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all values above

        icCex.lngSize = LenB(icCex)
        icCex.lngICC = ICC_STANDARD_CLASSES ' vb intrinsic controls (buttons, textbox, etc)
        ' if using Common Controls; add appropriate ICC_ constants for type of control you are using
        ' example if using CommonControls v5.0 Progress bar:
        ' .lngICC = ICC_STANDARD_CLASSES Or ICC_PROGRESS_CLASS

        On Error Resume Next ' error? InitCommonControlsEx requires IEv3 or above
        hMod = LoadLibraryA("shell32.dll") ' patch to prevent XP crashes when VB usercontrols present
        InitCommonControlsEx icCex
  If Err() Then
            InitCommonControls ' try Win9x version
            Err.Clear()
        End If
        On Error GoTo 0
        '... show your main form next (i.e., Form1.Show)

        '    If hMod Then FreeLibrary hMod


        '** Tip 1: Avoid using VB Frames when applying XP/Vista themes
        '          In place of VB Frames, use pictureboxes instead.
        '** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
        '          Doing so will prevent them from being themed.
    End Sub

End Module
