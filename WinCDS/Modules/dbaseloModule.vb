Module dbaseloModule
    '::::dbaseIoModule.bas
    ':::SUMMARY
    ': This module contains functions for the functioning of A/P and to perform multiple operations with its Database.
    ':::DESCRIPTION
    ': This module contains functions required to open / close the AP database.
    ': Also contains functions to update AP Transaction, Bank Account, Invoice Data, Factory email, & Vendor Name.
    Private dbGen As CDbAccessGeneral

    Public Function dbClose() As Boolean
        '::::dbClose
        ':::SUMMARY
        ': Close Accounting DB
        ':::DESCRIPTION
        ': This function is used to close the Database.
        ':
        ': Whichever connection was previously opened by the dbOpen is closed by this function.
        ':::PARAMETERS
        ':::SEE ALSO
        ': dbOpen
        ':::RETURN
        ': Returns true
        On Error Resume Next
        dbGen.dbClose
        dbGen = Nothing
        dbClose = True
    End Function

End Module
