VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ThisUserForm
   Caption         =   "Main"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   OleObjectBlob   =   "ThisUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ThisUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Requires Runtime

Private Const vApplicationPadding As Long = 50
Private Const vWebBrowserPadding As Long = 4

Private Sub UserForm_Activate()
    ' Set the title and icon of the current window.
    Caption = Runtime.ProjectName()
    Call Runtime.SetActiveWindowIcon

    ' Populate the current window with standard controls and maximize it.
    Call Runtime.PopulateActiveWindowTitlebar
    Call Runtime.MaximizeActiveWindow

    ' Load the main HTML file as the basis of the pages to be displayed in the embedded web browser with the startup navigate path.
    Call ThisWebBrowser.Navigate(Runtime.BaseHtmlFilePath() & "#" & Runtime.StartupNavigatePath())
End Sub

Private Sub UserForm_Layout()
    ' Load the excel application instance.
    With Application
        ' Check whether there is a need to move the application window.
        If _
            (.Left < (Left + vApplicationPadding)) _
            And ((.Left + .Width) > (Left + Width + vApplicationPadding)) _
            And (.Top < (Top + vApplicationPadding)) _
            And ((.Top + .Height) > (Top + Height + vApplicationPadding)) _
        Then
            Exit Sub
        End If

        ' Move the application window to the center of the user form.
        .Left = Left + (Width - .Width) / 2
        .Top = Top + (Height - .Height) / 2
    End With
End Sub

Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer _
)
    If Runtime.StartupNavigatePath() <> Runtime.vTestNavigatePath Then
        Call Runtime.Navigate(Runtime.vCloseNavigatePath)
    End If

    Call Hide
End Sub

Private Sub UserForm_Resize()
    ' Resize the embedded web browser.
    With ThisWebBrowser
        .Width = InsideWidth + vWebBrowserPadding
        .Height = InsideHeight + vWebBrowserPadding
    End With
End Sub

Private Sub ThisWebBrowser_DocumentComplete( _
    ByVal pDisp As Object, _
    URL As Variant _
)
    If Runtime.StartupNavigatePath() = Runtime.vTestNavigatePath Then
        If Runtime.IsDebugModeEnabled() Then
            Call Runtime.ExecuteTests
        Else
            Call Hide
        End If
    Else
        Call Runtime.Navigate(Right(URL, Len(URL) - InStr(URL, "#")))
    End If
End Sub

Public Sub Navigate( _
    vNavigatePath As String _
)
    Call ThisWebBrowser.Navigate(Runtime.BaseHtmlFilePath() & "?" & CStr(CDbl(Now)) & "#" & vNavigatePath)
End Sub

Public Sub SetInnerHtml( _
    vHtmlText As String _
)
    ThisWebBrowser.Document.body.InnerHtml = vHtmlText
    DoEvents
End Sub
