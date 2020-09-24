Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Module modVBErros
    Dim Printer As New Printer
    Private vbErrors_Number As Integer, vbErrors_Number_Hex As String, vbErrors_Description As String

    Public Function CheckStandardErrors(Optional ByVal Operation As String = "", Optional ByVal DefaultNotification As String = "") As Boolean
        CheckStandardErrors = vbErrorsParse()
        If Not CheckStandardErrors Then Exit Function

        ' All Below, if a specific error is detected, will report the standard error message.
        ' No additional error messages will be needed on any other call to this function.
        ' Use "#" to suppress messages...  All support functions should use it.
        ' All Below will default to True on Exit Function because of above.
        Select Case vbErrors_Number
'3 Return without GoSub
'5 Invalid procedure call
'6 Overflow
'7 Out of memory
'9 Subscript out of range
'10  This array is fixed or temporarily locked
'11  Division by zero
'13  Type mismatch
'14  Out of string space
'16  Expression too complex
'17  Can 't perform requested operation
'18  User interrupt occurred
'20  Resume without error
'28  Out of stack space
'35  Sub, Function, or Property not defined
'47  Too many DLL application clients
'48  Error in loading DLL
'49  Bad DLL calling convention
'51  Internal Error
'52  Bad file name or number
'53  File Not Found
'54  Bad file mode
'55  File already open
'57  Device I/O error
'58  File already exists
'59  Bad record length
'61  Disk Full
'62  Input past end of file
'63  Bad record number
'67  Too many files
'68  Device unavailable
'70  Permission denied
'71  Disk Not Ready
'74  Can 't rename with different drive
'75  Path/File access error
'76  Path Not Found
'91  Object variable or With block variable not set
'92  For loop not initialized
'93  Invalid pattern string
'94  Invalid use of Null
'97  Can 't call Friend procedure on an object that is not an instance of the defining class
'98  A property or method call cannot include a reference to a private object, either as an argument or as a return value
'298 System DLL could not be loaded
'320 Can 't use character device names in specified file names
'321 Invalid file format
'322 Cant create necessary temporary file
'325 Invalid format in resource file
'327 Data value named not found
'328 Illegal parameter; can't write arrays
'335 Could not access system registry
'336 Component not correctly registered
'337 COMPONENT Not Found
'338 Component did not run correctly
'360 Object already loaded
'361 Can 't load or unload this object
'363 Control Not Specified
'364 Object was unloaded
'365 Unable to unload within this context
'368 The specified file is out of date. This program requires a later version
'371 The specified object can't be used as an owner form for Show
'380 Invalid property value
'381 Invalid property-array index
'382 Property Set can't be executed at run time
'383 Property Set can't be used with a read-only property
'385 Need property-array index
'387 Property Set not permitted
'393 Property Get can't be executed at run time
'394 Property Get can't be executed on write-only property
'400 Form already displayed; can't show modally
'402 Code must close topmost modal form first
'419 Permission to use object denied
'422 Property not found
'423 Property or method not found
'424 object Required
'425 Invalid object use
'429 COMPONENT Can 't create object or return reference to this object
'430 Class doesn 't support Automation
'432 File name or class name not found during Automation operation
'438 object doesn 't support this property or method
'440 Automation Error
'442 Connection to type library or object library for remote process has been lost
'443 Automation object doesn't have a default value
'445 object doesn 't support this action
'446 object doesn 't support named arguments
'447 object doesn 't support current locale setting
'448 Named Not Argument
'449 Argument not optional or invalid property assignment
'450 Wrong number of arguments or invalid property assignment
'451 Object not a collection
'452 Invalid Ordinal
'453 Specified Not Found
'454 Code Not resource
'455 Code resource lock error
'457 This key is already associated with an element of this collection
'458 Variable uses a type not supported in Visual Basic
'459 This component doesn't support the set of events
'460 Invalid Clipboard format
'461 Method or data member not found
'462 The remote server machine does not exist or is unavailable
'463 Class not registered on local machine
'480 Can 't create AutoRedraw image
'481 Invalid Picture
'482 Printer Error
'483 Printer driver does not support specified property
'484 Problem getting printer information from the system. Make sure the printer is set up correctly
'485 Invalid picture type
'486 Can 't print form image to this type of printer
'520 Can 't empty Clipboard
'521 Can 't open Clipboard
'735 Can 't save file to TEMP directory
'744 Search Not Text
'746 Replacements too long
'31001 Out of memory
'31004 No object
'31018 Class is not set
'31027 Unable to activate object
'31032 Unable to create embedded object
'31036 Error saving to file
'31037 Error loading from file
            Case 482 : If ErrNoPrinter(Operation) Then Exit Function
            Case 16389 : If ErrDBErrors(Operation) Then Exit Function
        End Select

        CheckStandardErrors = vbErrorsParse(True)
        If DefaultNotification <> "" Then vbErrorsMessage(DefaultNotification)
    End Function

    Private Function vbErrorsParse(Optional ByVal ClearOnly As Boolean = False) As Boolean
        vbErrorsParse = False

        vbErrors_Number = 0         ' make sure they're cleared
        vbErrors_Number_Hex = ""
        vbErrors_Description = ""

        If ClearOnly Then Exit Function

        vbErrors_Number = Err.Number

        If vbErrors_Number = 0 Then Exit Function

        vbErrors_Number_Hex = Hex(vbErrors_Number)
        vbErrors_Description = Err.Description

        vbErrorsParse = True        ' Return True when there is an error
    End Function

    Public Function ErrNoPrinter(Optional ByVal Operation As String = "") As Boolean
        Dim S As String
        vbErrorsParse()

        ErrNoPrinter = False
        If vbErrors_Number = 482 Then
            If IsInStrArray(Printer.DeviceName, "Print to PDF", "Microsoft XPS Document Writer", "Cute PDF") Then
                ' Do not report errors to "Microsoft Print to PDF"... They just hit ESC to abort the printing
                ' Have to return TRUE... we "handled" this error, we just didn't "do" anything
                ErrNoPrinter = True
                Exit Function
            End If
            S = ""
            S = S & "Windows Printer Error (482)" & IIf(Operation = "", "", " in task [" & Operation & "].") & vbCrLf
            S = S & "Description: " & vbErrors_Description & vbCrLf
            S = S & "You apparently do not have any printer set up on this machine." & vbCrLf2
            S = S & "Please configure a printer and try this function again."

            ErrNoPrinter = vbErrorsMessage(S, "Printer Error", vbFalse, Operation)
        End If
    End Function

    Public Function ErrDBErrors(Optional ByVal Operation As String = "") As Boolean
        Dim S As String, M As Boolean

        vbErrorsParse()
        ErrDBErrors = False
        S = "Database Error (" & vbErrors_Number_Hex & ")" & vbCrLf
        S = S & vbErrors_Description & vbCrLf2

        Select Case vbErrors_Number
            Case &H80004005
                ' 80004005 - Operation must use an updateable query.
                If InStr(Err.Description, "updateable") > 0 Then M = True : S = S & "This is almost always a permissions issue."

                ' 80004005 - Not a valid bookmark.
                If InStr(Err.Description, "valid bookmark") > 0 Then M = True : S = S & "Doing a Compact / Repair on this database usually fixes this issue."

                If Not M Then Operation = "#"
                ErrDBErrors = vbErrorsMessage(S, "Database Error", True, Operation)

                'BFH20170213 - Not going to use here yet...
                '      ReportError S, vbErrors_Number, vbErrors_Description
        End Select
    End Function

    Private Function vbErrorsMessage(ByVal Msg As String, Optional ByVal Task As String = "Error", Optional ByVal Critical As TriState = vbUseDefault, Optional ByVal Operation As String = "") As Boolean
        Dim Style As MsgBoxStyle

        vbErrorsMessage = True
        If Operation = "#" Then Exit Function

        Select Case Critical
            Case vbTrue : Style = vbCritical
            Case vbFalse : Style = vbExclamation
            Case vbUseDefault : Style = vbInformation
        End Select
        MsgBox(Msg, Style, ProgramName & " " & Trim(Task) & IIf(Operation = "", "", " - " & Operation))
    End Function
End Module
