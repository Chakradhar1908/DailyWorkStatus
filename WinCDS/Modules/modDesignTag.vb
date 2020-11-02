Module modDesignTag
    Public Sub PrintCustomTags(ByVal Style As String, ByVal Quantity As Integer, ByVal TemplateName As String)
        '::::PrintCustomTags
        ':::SUMMARY
        ': Print a custom tag
        ':::DESCRIPTION
        ': This function is used to print single or multiple customised Tags in Tag Designer Form under Reports in Inventory Menu.
        ': To print multiple tags, select the Manufacturer, Department, or partial style numbers you wish to print.
        ': Select whether you wish to print one for each On Hand item or simply 1 tag per style number in your selection.
        ': The total number of tags that will be printed will be displayed in the box labeled Tags to be printed:.
        ':::PARAMETERS
        ': - Style
        ': - Quantity
        ': - TemplateName

        Dim DidLoad As Boolean
        DidLoad = Not IsFormLoaded("frmDesignTag")
        frmDesignTag.PrintCustomTags(Style, Quantity, TemplateName)
        'If DidLoad Then Unload frmDesignTag
        If DidLoad Then frmDesignTag.Close()
    End Sub

    Public Sub LoadCustomTagLayoutsToComboBox(ByRef Cbo As ComboBox)
        '::::LoadCustomTagLayoutsToComboBox
        ':::SUMMARY
        ': Load layout to destination
        ':::DESCRIPTION
        ': Loads the list of custom tag names to the combo box on the form.
        ':::PARAMETERS
        ': - Cbo
        Dim I As Integer, Pre As String, F As String
        On Error Resume Next
        Cbo.Items.Clear()
        Cbo.Items.Add("-Select From List-")
        MainMenu.flb.Path = InventFolder()
        MainMenu.flb.Path = TagLayoutFolder()
        Pre = "taglayout-"
        MainMenu.flb.Pattern = Pre & "*.txt"
        'For I = 0 To MainMenu.flb.ListCount - 1
        For I = 0 To MainMenu.flb.Items.Count - 1
            'F = MainMenu.flb.List(I)
            F = MainMenu.flb.Items(I).ToString
            F = Mid(F, Len(Pre) + 1)
            F = Left(F, Len(F) - 4)
            Cbo.Items.Add(F)
        Next

        Cbo.SelectedIndex = 0
    End Sub

    Public Sub LoadTagLayoutTemplatesToComboBox(ByRef Cbo As ComboBox)
        '::::LoadTagLayoutTemplatesToComboBox
        ':::SUMMARY
        ': Prep form for display
        ':::DESCRIPTION
        ': Load Layout Template names to combo box for UI.
        ':::PARAMETERS
        ': - Cbo

        Dim I As Integer, Pre As String, F As String
        On Error Resume Next
        Cbo.Items.Clear()
        Cbo.Items.Add("(Default)")
        MainMenu.flb.Path = TagLayoutFolder()
        Pre = "TT-"
        MainMenu.flb.Pattern = Pre & "*.txt"
        'For I = 0 To MainMenu.flb.ListCount - 1
        For I = 0 To MainMenu.flb.Items.Count - 1
            'F = MainMenu.flb.List(I)
            F = MainMenu.flb.Items(I).ToString
            F = Mid(F, Len(Pre) + 1)
            F = Left(F, Len(F) - 4)
            Cbo.Items.Add(F)
        Next

        Cbo.SelectedIndex = 0
    End Sub
End Module
