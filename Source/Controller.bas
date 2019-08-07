Attribute VB_Name = "Controller"
Option Explicit
Option Private Module

' Requires ThisUserForm
' Requires Runtime

Public Sub Navigate( _
    ByRef vPath As String, _
    ByRef vParameters As Dictionary _
)
    ' If debug mode is enabled, print the current state to the immediate window.
    If Runtime.IsDebugModeEnabled() Then
        Debug.Print "===================== Output ====================="
        Debug.Print "[Date & Time] " & Format(Now, "yyyy-mm-dd Hh:Nn:Ss")
        Debug.Print "[User @ Computer] " & Runtime.Username() & "@" & Runtime.ComputerName()
        Debug.Print "[Navigate Path] " & Runtime.GenerateNavigatePath(vPath, vParameters)
    End If

    ' Check whether the application is running in background mode.
    If Runtime.IsBackgroundModeEnabled() Then
        ' Output the current timestamp to a file.
        With Runtime.FileSystemObject()
            With .OpenTextFile(.BuildPath(ThisWorkbook.Path, "Output.log"), ForAppending, True)
                Call .WriteLine("[Date & Time] " & Format(Now, "yyyy-mm-dd Hh:Nn:Ss"))
                Call .WriteLine("[User @ Computer] " & Runtime.Username() & "@" & Runtime.ComputerName())
                Call .WriteLine("[Navigate Path] " & Runtime.GenerateNavigatePath(vPath, vParameters))
                Call .WriteLine
                Call .Close
            End With
        End With
    Else
        ' Display the path and parameters on the loaded html page.
        Call ThisUserForm.SetInnerHtml("<h1>Navigate</h1>" _
            & "<p><b>Date & Time</b>: <code>" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") & "</code></p>" _
            & "<p><b>User @ Computer</b>: <code>" & Runtime.Username() & "@" & Runtime.ComputerName() & "</code></p>" _
            & "<p><b>Navigate Path</b>: <code>" & Runtime.GenerateNavigatePath(vPath, vParameters) & "</code></p>")
    End If
End Sub

Public Sub HandleError( _
    ByRef vPath As String, _
    ByRef vParameters As Dictionary, _
    ByRef vErrorMessage As String _
)
    ' If debug mode is enabled, print the current state to the immediate window.
    If Runtime.IsDebugModeEnabled() Then
        Debug.Print "===================== Error ======================"
        Debug.Print "[Date & Time] " & Format(Now, "yyyy-mm-dd Hh:Nn:Ss")
        Debug.Print "[User @ Computer] " & Runtime.Username() & "@" & Runtime.ComputerName()
        Debug.Print "[Navigate Path] " & Runtime.GenerateNavigatePath(vPath, vParameters)
        Debug.Print "[Error Number] " & CStr(Err.Number)
        Debug.Print "[Error Source] " & Err.Source
        Debug.Print "[Error Description] " & Err.Description
        Debug.Print "[Error Message] " & vErrorMessage
    End If

    ' Check whether the application is running in background mode.
    If Runtime.IsBackgroundModeEnabled() Then
        ' Output the current timestamp to a file.
        With Runtime.FileSystemObject()
            With .OpenTextFile(.BuildPath(ThisWorkbook.Path, "Error.log"), ForAppending, True)
                Call .WriteLine("[Date & Time] " & Format(Now, "yyyy-mm-dd Hh:Nn:Ss"))
                Call .WriteLine("[User @ Computer] " & Runtime.Username() & "@" & Runtime.ComputerName())
                Call .WriteLine("[Navigate Path] " & Runtime.GenerateNavigatePath(vPath, vParameters))
                Call .WriteLine("[Error Number] " & CStr(Err.Number))
                Call .WriteLine("[Error Source] " & Err.Source)
                Call .WriteLine("[Error Description] " & Err.Description)
                Call .WriteLine("[Error Message] " & vErrorMessage)
                Call .WriteLine
                Call .Close
            End With
        End With
    Else
        ' Display the path and parameters on the loaded html page.
        Call ThisUserForm.SetInnerHtml("<h1>Handle Error</h1>" _
            & "<p><b>Date & Time</b>: <code>" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") & "</code></p>" _
            & "<p><b>User @ Computer</b>: <code>" & Runtime.Username() & "@" & Runtime.ComputerName() & "</code></p>" _
            & "<p><b>Navigate Path</b>: <code>" & Runtime.GenerateNavigatePath(vPath, vParameters) & "</code></p>" _
            & "<p><b>Error Number</b>: <code>" & CStr(Err.Number) & "</code></p>" _
            & "<p><b>Error Source</b>: <code>" & Err.Source & "</code></p>" _
            & "<p><b>Error Description</b>: <code>" & Err.Description & "</code></p>")

        Call MsgBox(IIf(vErrorMessage = vbNullString, "An unknown unexpected error had occurred.", vErrorMessage), _
            vbCritical, "Error Message")
    End If
End Sub

Public Sub ExecuteTestCase( _
    ByRef vModuleName As String, _
    ByRef vCaseName As String _
)
    Select Case vModuleName
        Case Else
            Call Runtime.RaiseUndefinedTestModuleHandler
    End Select
End Sub

'''''''''''''''''''''''
'                     '
' Procedure Template: '
'                     '
'''''''''''''''''''''''

' Public [Sub | Function] ProcedureName()
'     ' Declare local variables.
'     ' TODO: Implement.

'     ' Setup error handling.
'     On Error GoTo HandleError:

'     ' Allocate resources.
'     ' TODO: Implement.

'     ' Implement the application logic.
'     ' TODO: Implement.

' Terminate:
'     ' Reset error handling.
'     On Error GoTo 0

'     ' Release all allocated resources if needed.
'     ' TODO: Implement.

'     ' Re-raise any stored error.
'     Call Runtime.ReRaiseError

'     ' Exit the procedure.
'     Exit [Sub | Function]

' HandleError:
'     ' Store the error for further handling.
'     Call Runtime.StoreError

'     ' TODO: Verify whether the error should be re-raised.

'     ' Resume to procedure termination.
'     Resume Terminate:
' End [Sub | Function]
