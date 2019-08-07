Attribute VB_Name = "Controller"
Option Explicit
Option Private Module

' Requires reference: Scripting
' Requires module: Runtime
' Requires module: ThisUserForm
' Requires module: ThisWorkbook

Public Sub Navigate( _
    ByRef vPath As String, _
    ByRef vParameters As Scripting.Dictionary _
)
    ' Declare local variables.
    Dim vValues As New Scripting.Dictionary
    Dim vName As Variant
    Dim vOutput As String

    ' Fill the output values.
    With vValues
        Call .Add("Date & Time", Runtime.DateTimeStamp(Now))
        Call .Add("User @ Computer", Runtime.Username() & "@" & Runtime.ComputerName())
        Call .Add("Navigate Path", Runtime.GenerateNavigatePath(vPath, vParameters))
    End With

    ' If debug mode is enabled, print the current state to the immediate window.
    If Runtime.IsDebugModeEnabled() Then
        Debug.Print "===================== Output ====================="
        For Each vName In vValues.Keys()
            Debug.Print Trim("[" & CStr(vName) & "] " & CStr(vValues(vName)))
        Next
    End If

    ' Check whether the application is running in background mode.
    If Runtime.IsBackgroundModeEnabled() Then
        ' Output the current state to a file.
        For Each vName In vValues.Keys()
            vOutput = vOutput & Trim("[" & CStr(vName) & "] " & CStr(vValues(vName))) & vbLf
        Next
        With Runtime.FileSystemObject()
            Call Runtime.AppendFile(.BuildPath(ThisWorkbook.Path, "Output.log"), vOutput, vbLf)
        End With
    Else
        ' Display the current state on the loaded html page.
        For Each vName In vValues.Keys()
            vOutput = vOutput & "<p><b>" & CStr(vName) & "</b>: <code>" & CStr(vValues(vName)) & "</code></p>"
        Next
        Call ThisUserForm.SetInnerHtml("<h1>Navigate</h1>" & vOutput)
    End If
End Sub

Public Sub HandleError( _
    ByRef vPath As String, _
    ByRef vParameters As Scripting.Dictionary, _
    ByRef vErrorMessage As String _
)
    ' Declare local variables.
    Dim vValues As New Scripting.Dictionary
    Dim vName As Variant
    Dim vOutput As String

    ' Fill the output values.
    With vValues
        Call .Add("Date & Time", Runtime.DateTimeStamp(Now))
        Call .Add("User @ Computer", Runtime.Username() & "@" & Runtime.ComputerName())
        Call .Add("Navigate Path", Runtime.GenerateNavigatePath(vPath, vParameters))
        Call .Add("Error Number", CStr(Err.Number))
        Call .Add("Error Source", Err.Source)
        Call .Add("Error Description", Err.Description)
    End With

    ' If debug mode is enabled, print the current state to the immediate window.
    If Runtime.IsDebugModeEnabled() Then
        Debug.Print "===================== Error ======================"
        For Each vName In vValues.Keys()
            Debug.Print Trim("[" & CStr(vName) & "] " & CStr(vValues(vName)))
        Next
        Debug.Print Trim("[Error Message] " & vErrorMessage)
    End If

    ' Check whether the application is running in background mode.
    If Runtime.IsBackgroundModeEnabled() Then
        ' Output the current state to a file.
        For Each vName In vValues.Keys()
            vOutput = vOutput & Trim("[" & CStr(vName) & "] " & CStr(vValues(vName))) & vbLf
        Next
        vOutput = vOutput & Trim("[Error Message] " & vErrorMessage) & vbLf
        With Runtime.FileSystemObject()
            Call Runtime.AppendFile(.BuildPath(ThisWorkbook.Path, "Error.log"), vOutput, vbLf)
        End With
    Else
        ' Display the current state on the loaded html page.
        For Each vName In vValues.Keys()
            vOutput = vOutput & "<p><b>" & CStr(vName) & "</b>: <code>" & CStr(vValues(vName)) & "</code></p>"
        Next
        Call ThisUserForm.SetInnerHtml("<h1>Navigate</h1>" & vOutput)

        ' Show dialog box with an error message.
        Call MsgBox(IIf(vErrorMessage = vbNullString, "An unknown unexpected error had occurred.", vErrorMessage), _
            vbCritical, "Error Message")
    End If
End Sub

Public Sub ExecuteTestCase( _
    ByRef vModuleName As String, _
    ByRef vCaseName As String _
)
    Select Case vModuleName
        Case "ModuleName"
            Select Case vCaseName
                Case "CaseName"
                    ' Call ModuleName.Case_CaseName
                Case Else
                    Call Runtime.RaiseUndefinedTestCaseHandler
            End Select
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
