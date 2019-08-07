Option Explicit

' Requires Runtime

Private Const vMinimumWidth As Long = 110
Private Const vMinimumHeight As Long = 30

Private Const vDefaultWidth As Long = 800
Private Const vDefaultHeight As Long = 400

Private vOriginalLeft As Long
Private vOriginalTop As Long

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
            ' Store the original dimensions.
            vOriginalLeft = .Left
            vOriginalTop = .Top

            ' Shrink the application window.
            .Width = vMinimumWidth
            .Height = vMinimumHeight

            ' Show the main user form.
            Call ThisUserForm.Show

            ' Prevent the flickering of the application window, if it is to be closed.
            If Not Runtime.IsDebugModeEnabled() Then
                .Visible = False
            End If

            ' Restore the original dimensions.
            .Left = vOriginalLeft
            .Top = vOriginalTop

            ' Set the dimensions of the window.
            .Width = vDefaultWidth
            .Height = vDefaultHeight
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
            ' Set the dimensions of the window.
            .Width = vDefaultWidth
            .Height = vDefaultHeight

            ' Set the title and icon of the window.
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

    ' Set the startup navigate path environment variable to test.
    Runtime.WScriptShell().Environment("PROCESS")("APP_STARTUP_NAVIGATE_PATH") = Runtime.vTestNavigatePath

    ' Initialize the application.
    Call pInitialize
End Sub
