Public Class InvPo
    Public RN as integer
    Public PoNo as integer           ' Allows adding multiple items to same PO.  Cleared when a PO is printed.
    Dim Search As New CSearchNew  ' Global on purpose - saves search results for later use.
    Dim Search_Loaded As Boolean  ' Re-search prevention.
    Dim SearchMode as integer        ' All, Dept, Mfg
    Dim OldVendor As Object       ' Vendor of last PO, for validation.
    Dim OldStore As Byte          ' Store of last PO, for validation.

End Class