Option Explicit

' Requires module: Runtime
' Requires module: ThisUserForm

Private Const vMinimumWidth As Long = 110
Private Const vMinimumHeight As Long = 30

Private Sub pInitialize()
    ' Check whether the application is running in background mode and not running tests.
    If _
        Runtime.IsBackgroundModeEnabled() _
        And (Runtime.StartupNavigatePath() <> Runtime.vTestNavigatePath) _
    Then
        ' Execute the startup navigate path.
        Call Runtime.Navigate(Runtime.StartupNavigatePath())
    Else
        ' Load the excel application instance.
        With Application
            ' Shrink the application window.
            .WindowState = xlNormal
            .Width = vMinimumWidth
            .Height = vMinimumHeight

            ' Show the main user form.
            Call ThisUserForm.Show

            ' Determine if the application will be closed.
            If Runtime.IsDebugModeEnabled() Then
                ' Set the dimensions and position of the window.
                .WindowState = xlMaximized
            Else
                ' Prevent the flickering of the application window before closing.
                .Visible = False
            End If
        End With
    End If
End Sub

Private Sub Workbook_BeforeClose( _
    Cancel As Boolean _
)
    If _
        (Not Runtime.IsDebugModeEnabled()) _
        Or Runtime.IsDeployDebugModeEnabled() _
    Then
        Me.Saved = True
    End If
End Sub

Private Sub Workbook_Open()
    ' Load the current excel application instance.
    With Application
        ' Check whether the application is visible.
        If .Visible Then
            ' Maximize the window and set its title and icon.
            .WindowState = xlMaximized
            .ActiveWindow.Caption = vbNullString
            .Caption = Runtime.ProjectName()
            Call Runtime.SetActiveWindowIcon
        End If

        ' Disable unnecessary activities.
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    ' If the application is in debug mode, do not continue.
    If Runtime.IsDebugModeEnabled() Then
        Exit Sub
    End If

    ' Initialize the application.
    Call pInitialize

    ' Close the excel application instance.
    Call Application.Quit
End Sub

Public Sub Initialize()
    ' If the application is not in debug mode, do not continue.
    If Not Runtime.IsDebugModeEnabled() Then
        Exit Sub
    End If

    ' Save the main workbook file.
    Call Me.Save

    ' Set the startup navigate path environment variable to user input.
    Runtime.WScriptShell().Environment("PROCESS")("APP_STARTUP_NAVIGATE_PATH") = InputBox("Enter the path to navigate to")

    ' Initialize the application.
    Call pInitialize
End Sub

Public Sub Test()
    ' If the application is not in debug mode, do not continue.
    If Not Runtime.IsDebugModeEnabled() Then
        Exit Sub
    End If

    ' Save the main workbook file.
    Call Me.Save

    ' Set the startup navigate path environment variable to test.
    Runtime.WScriptShell().Environment("PROCESS")("APP_STARTUP_NAVIGATE_PATH") = Runtime.vTestNavigatePath

    ' Initialize the application.
    Call pInitialize
End Sub
