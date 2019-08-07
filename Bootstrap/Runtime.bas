Attribute VB_Name = "Runtime"
Option Explicit
Option Private Module

' Requires Controller.

Private Declare Function GetActiveWindow Lib "user32" () As Integer

Private Declare Function ExtractIconA Lib "shell32.dll" ( _
    ByVal hInst As Long, _
    ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long _
) As Long

Private Declare Function SendMessageA Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function GetWindowLongA Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long _
) As Long

Private Declare Function SetWindowLongA Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long

Private Declare Function ShowWindow Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nCmdShow As Long _
) As Long

Private Const vGwlStyle As Long = -16
Private Const vWsMaximizeBox As Long = &H10000
Private Const vWsMinimizeBox As Long = &H20000
Private Const vWsThickFrame As Long = &H40000
Private Const vWsSystemMenu As Long = &H80000
Private Const vSwShowMaximized As Long = 3

Private Const vDefaultErrorNumber As Long = 10000

Public Const vTestNavigatePath As String = "test"
Public Const vCloseNavigatePath As String = "close"

Private Const vTestModuleNamePrefix As String = "Test"
Private Const vTestCaseDeclarationPrefix As String = "Public Sub Case_"
Private Const vTestCaseDeclarationSuffix As String = "()"

Private vIsErrorStored As Boolean
Private vIsErrorIntercepted As Boolean

Private vStoredErrorNumber As Long
Private vStoredErrorSource As String
Private vStoredErrorDescription As String
Private vStoredErrorMessage As String

Private vFileSystemObject As FileSystemObject
Private vWScriptShell As Object

Public Function FileSystemObject() As FileSystemObject
    ' Initialize the file system object for use across the project, if needed.
    If vFileSystemObject Is Nothing Then
        Set vFileSystemObject = New FileSystemObject
    End If
    Set FileSystemObject = vFileSystemObject
End Function

Public Function WScriptShell() As Object
    ' Initialize the wscript shell object for use across the project, if needed.
    If vWScriptShell Is Nothing Then
        Set vWScriptShell = CreateObject("WScript.Shell")
    End If
    Set WScriptShell = vWScriptShell
End Function

Public Function IsDebugModeEnabled() As Boolean
    IsDebugModeEnabled = WScriptShell().Environment("PROCESS")("APP_IS_DEBUG_MODE_ENABLED") = "TRUE"
End Function

Public Function IsDeployDebugModeEnabled() As Boolean
    IsDeployDebugModeEnabled = WScriptShell().Environment("PROCESS")("APP_IS_DEPLOY_DEBUG_MODE_ENABLED") = "TRUE"
End Function

Public Function IsBackgroundModeEnabled() As Boolean
    IsBackgroundModeEnabled = WScriptShell().Environment("PROCESS")("APP_IS_BACKGROUND_MODE_ENABLED") = "TRUE"
End Function

Public Function ProjectName() As String
    ProjectName = WScriptShell().Environment("PROCESS")("APP_PROJECT_NAME")
End Function

Public Function StartupNavigatePath() As String
    StartupNavigatePath = WScriptShell().Environment("PROCESS")("APP_STARTUP_NAVIGATE_PATH")
End Function

Public Function Username() As String
    Username = WScriptShell().Environment("PROCESS")("USERNAME")
End Function

Public Function ComputerName() As String
    ComputerName = WScriptShell().Environment("PROCESS")("COMPUTERNAME")
End Function

Public Function ConfigFilePath() As String
    ConfigFilePath = FileSystemObject().BuildPath(ThisWorkbook.Path, "Config.xml")
End Function

Public Function ErrorFilePath() As String
    ErrorFilePath = FileSystemObject().BuildPath(ThisWorkbook.Path, "Error.log")
End Function

Public Function IconFilePath() As String
    With FileSystemObject()
        IconFilePath = .BuildPath(.BuildPath(ThisWorkbook.Path, "Assets"), "Main.ico")
    End With
End Function

Public Function BaseHtmlFilePath() As String
    With FileSystemObject()
        BaseHtmlFilePath = .BuildPath(.BuildPath(ThisWorkbook.Path, "Assets"), "Main.html")
    End With
End Function

Public Sub SetActiveWindowIcon()
    ' Send the api message that loads and sets an icon for the currently active window.
    Call SendMessageA(GetActiveWindow(), &H80, 0, ExtractIconA(0, IconFilePath(), 0))
End Sub

Public Sub PopulateActiveWindowTitlebar()
    ' Declare local variables.
    Dim vFormHandle As Long
    Dim vWindowStyle As Long

    ' Retrieve the form handle of the currently active window.
    vFormHandle = GetActiveWindow()

    ' Retrieve the new window style information for the currently active window.
    vWindowStyle = GetWindowLongA(vFormHandle, vGwlStyle)

    ' Add the desired properties to the retrieved new window style information.
    vWindowStyle = vWindowStyle Or vWsMaximizeBox
    vWindowStyle = vWindowStyle Or vWsMinimizeBox
    vWindowStyle = vWindowStyle Or vWsThickFrame
    vWindowStyle = vWindowStyle Or vWsSystemMenu

    ' Set the configured new window style information to the currently active window.
    Call SetWindowLongA(vFormHandle, vGwlStyle, vWindowStyle)
End Sub

Public Sub MaximizeActiveWindow()
    Call ShowWindow(GetActiveWindow(), vSwShowMaximized)
End Sub

Public Sub SetErrorMessage( _
    ByVal vMessage As String _
)
    vStoredErrorMessage = vMessage
End Sub

Public Sub ClearErrorMessage()
    vStoredErrorMessage = vbNullString
End Sub

Public Sub RaiseError( _
    ByRef vSource As String, _
    ByRef vDescription As String, _
    Optional ByVal vMessage As String = vbNullString _
)
    ' Store the error message.
    If vMessage <> vbNullString Then
        vStoredErrorMessage = vMessage
    End If

    ' Raise the error with the correct number and description.
    Call Err.Raise(vDefaultErrorNumber, vSource, vDescription)
End Sub

Public Sub StoreError()
    ' Check whether the error had already been intercepted.
    If Not vIsErrorIntercepted Then
        ' Start debugging if in debug mode.
        Debug.Assert Not IsDebugModeEnabled()

        ' Set the error caught flag.
        vIsErrorIntercepted = True
    End If

    ' Set the error stored flag.
    vIsErrorStored = True

    ' Store the current error parameters.
    vStoredErrorNumber = VBA.Err.Number
    vStoredErrorSource = VBA.Err.Source
    vStoredErrorDescription = VBA.Err.Description
End Sub

Public Sub ReRaiseError()
    ' Verify that an error is stored.
    If vIsErrorStored Then
        ' Reset the error stored flag.
        vIsErrorStored = False

        ' ReRaise an error with the stored error parameters.
        Call Err.Raise(vStoredErrorNumber, vStoredErrorSource, vStoredErrorDescription)
    End If
End Sub

Public Function ParseNavigatePath( _
    ByRef vNavigatePath As String _
) As Dictionary
    ' Declare local variables.
    Dim vPath As String
    Dim vParametersPortionIndex As Long
    Dim vParameterEntry As Variant
    Dim vParsedParameterEntry() As String
    Dim vParameters As New Dictionary

    vPath = vNavigatePath
    vParametersPortionIndex = InStr(vPath, "?")
    If vParametersPortionIndex <> 0 Then
        For Each vParameterEntry In Split(Mid(vPath, vParametersPortionIndex + 1), "&")
            vParsedParameterEntry = Split(vParameterEntry, "=")
            If UBound(vParsedParameterEntry) = 0 Then
                ReDim Preserve vParsedParameterEntry(0 To 1)
                vParsedParameterEntry(1) = vbNullString
            End If
            Call vParameters.Add(vParsedParameterEntry(0), vParsedParameterEntry(1))
        Next
        vPath = Left(vPath, vParametersPortionIndex - 1)
    End If

    Set ParseNavigatePath = New Dictionary
    With ParseNavigatePath
        .Item("Path") = vPath
        Set .Item("Parameters") = vParameters
    End With
End Function

Public Function GenerateNavigatePath( _
    ByRef vPath As String, _
    ByRef vParameters As Dictionary _
) As String
    ' Declare local variables.
    Dim vParameterKey As Variant
    Dim vParameterKeyIndex As Long
    Dim vParameterEntries() As String

    GenerateNavigatePath = vPath
    If vParameters.Count > 0 Then
        ReDim vParameterEntries(0 To vParameters.Count - 1)
        vParameterKeyIndex = LBound(vParameterEntries)
        For Each vParameterKey In vParameters.Keys()
            vParameterEntries(vParameterKeyIndex) = vParameterKey & "=" & vParameters(vParameterKey)
            vParameterKeyIndex = vParameterKeyIndex + 1
        Next
        GenerateNavigatePath = GenerateNavigatePath & "?" & Join(vParameterEntries, "&")
    End If
End Function

Public Sub Navigate( _
    ByRef vNavigatePath As String _
)
    ' Declare local variables.
    Dim vHasErrorOccurred As Boolean
    Dim vPath As String
    Dim vParameters As Dictionary

    ' Extract the query parameters if available.
    With ParseNavigatePath(vNavigatePath)
        vPath = .Item("Path")
        Set vParameters = .Item("Parameters")
    End With

    ' Configure error handling.
    On Error GoTo HandleError:

    ' Pass the path and parameters to the user defined controller.
    Call Controller.Navigate(vPath, vParameters)

Terminate:
    ' Reset error handling
    On Error GoTo 0

    ' If an error had occurred and the userform is visible, hide the userform.
    If vHasErrorOccurred Then
        If ThisUserForm.Visible Then
            Call Unload(ThisUserForm)
        End If
    End If

    ' Exit the procedure.
    Exit Sub

HandleError:
    ' Set the error flag.
    vHasErrorOccurred = True

    ' Check whether the error had already been caught.
    If vIsErrorIntercepted Then
        ' Reset the error caught flag.
        vIsErrorIntercepted = False
    Else
        ' Start debugging if in debug mode.
        Debug.Assert Not IsDebugModeEnabled()
    End If

    ' Handle error reporting.
    Call Controller.HandleError(vPath, vParameters, vStoredErrorMessage)

    ' Clear the stored error.
    vIsErrorStored = False
    vStoredErrorNumber = 0
    vStoredErrorSource = vbNullString
    vStoredErrorDescription = vbNullString
    vStoredErrorMessage = vbNullString

    ' Terminate error handling.
    Resume Terminate:
End Sub

Public Sub ExecuteTests()
    ' Declare local variables.
    Dim vComponent As VBComponent
    Dim vCodeLinePosition As Long
    Dim vCodeLine As String
    Dim vCaseName As String
    Dim vProcedureBody As String
    Dim vModuleReportHtml As String
    Dim vModuleReportDots As String
    Dim vModulePassedCaseCount As Long
    Dim vModuleCaseCount As Long
    Dim vReportHtml As String
    Dim vReportDots As String
    Dim vPassedCaseCount As Long
    Dim vCaseCount As Long
    Dim vStyleAttribute As String

    ' Configure the error handler.
    On Error GoTo HandleError:

    ' Loop through all of the project's components.
    For Each vComponent In ThisWorkbook.VBProject.VBComponents
        ' Only work with test modules.
        If Left(vComponent.Name, Len(vTestModuleNamePrefix)) = vTestModuleNamePrefix Then
            ' Load the current component's code module.
            With vComponent.CodeModule
                ' Loop through each line of code.
                vCodeLinePosition = 0
                Do While vCodeLinePosition < .CountOfLines
                    ' Load the current code line.
                    vCodeLine = .Lines(vCodeLinePosition + 1, 1)

                    ' Check for the presence of a test case declaration.
                    If _
                        (Len(vCodeLine) > Len(vTestCaseDeclarationPrefix & vTestCaseDeclarationSuffix)) _
                        And (Left(vCodeLine, Len(vTestCaseDeclarationPrefix)) = vTestCaseDeclarationPrefix) _
                        And (Right(vCodeLine, Len(vTestCaseDeclarationSuffix)) = vTestCaseDeclarationSuffix) _
                    Then
                        ' Determine the current test case name.
                        vCaseName = Mid(vCodeLine, Len(vTestCaseDeclarationPrefix) + 1, _
                            Len(vCodeLine) - Len(vTestCaseDeclarationPrefix & vTestCaseDeclarationSuffix))

                        ' Increment the test case counter.
                        vModuleCaseCount = vModuleCaseCount + 1

                        ' Execute the current test case of the current test module (the placeholder is replaced during execution).
                        Call Controller.ExecuteTestCase(CStr(vComponent.Name), CStr(vCaseName))

                        ' Report the success of the current test case.
                        vModuleReportHtml = vModuleReportHtml & "<li><span style=""color: green"">[PASS]</span> <b>" & CStr(vCaseName) & "</b></li>"
                        vModuleReportDots = vModuleReportDots + "<span style=""color: green"">.</span>"
                        vModulePassedCaseCount = vModulePassedCaseCount + 1
EndTestCase:
                    End If

                    ' Increment the current code line position.
                    vCodeLinePosition = vCodeLinePosition + 1
                Loop
            End With

            ' Add a header to the report for the current test module.
            If vModulePassedCaseCount = vModuleCaseCount Then
                vStyleAttribute = "style=""color: green"""
            Else
                vStyleAttribute = "style=""color: red"""
            End If
            vModuleReportHtml = "<h2><span " & vStyleAttribute & ">(" & CStr(vModulePassedCaseCount) & " / " & CStr(vModuleCaseCount) & ")</span>" _
                & " " & CStr(vComponent.Name) & "</h2><pre>" & vModuleReportDots & "</pre>" & vModuleReportHtml

            ' Add the module's data to the overall result.
            vReportHtml = vReportHtml & vModuleReportHtml
            vModuleReportHtml = vbNullString
            vReportDots = vReportDots & vModuleReportDots
            vModuleReportDots = vbNullString
            vPassedCaseCount = vPassedCaseCount + vModulePassedCaseCount
            vModulePassedCaseCount = 0
            vCaseCount = vCaseCount + vModuleCaseCount
            vModuleCaseCount = 0
        End If
    Next

    ' Add a header to the report and output it to be rendered before exiting.
    If vPassedCaseCount = vCaseCount Then
        vStyleAttribute = "style=""color: green"""
    Else
        vStyleAttribute = "style=""color: red"""
    End If
    Call ThisUserForm.SetInnerHtml("<div style=""width: 40em; margin: 0 auto; padding: 1em; background: lightgoldenrodyellow"">" _
        & "<h1><span " & vStyleAttribute & ">(" & CStr(vPassedCaseCount) & " / " & CStr(vCaseCount) & ")</span>" _
        & " Test</h1><pre>" & vReportDots & "</pre>" & vReportHtml & "<div>")
    Exit Sub

HandleError:
    ' Report the failure of the current test case.
    vModuleReportHtml = vModuleReportHtml & "<li><span style=""color: red"">[FAIL]</span> <b>" & CStr(vCaseName) & "</b>" _
        & "<p style=""background: lightcoral; padding: 0.2em""><b>" & Err.Source & "</b><br />" & Replace(Err.Description, vbCrLf, "<br />") & "</p>"
    vModuleReportDots = vModuleReportDots + "<span style=""color: red"">X</span>"

    ' Transfer control to the end of the current test case.
    Resume EndTestCase:
End Sub

Public Sub Assert( _
    ByVal vValue As Boolean, _
    Optional ByRef vComment As String = "Undefined assertion error" _
)
    If Not vValue Then
        Call RaiseError("Runtime.Assert", vComment)
    End If
End Sub

Public Sub RaiseUndefinedTestModuleHandler()
    Call RaiseError("Controller", "Cannot find an executor section for the test module.")
End Sub

Public Sub RaiseUndefinedTestCaseHandler()
    Call RaiseError("Controller", "Cannot find an executor for the test case of the current test module.")
End Sub
