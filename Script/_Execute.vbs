Option Explicit

' Define the project parameter constants.
Const vProjectName = ""
Const vIsBackgroundModeEnabled = False
Const vMainWorkbookFilePassword = ""

' Declare local variables.
Dim vWScriptShell
Dim vMainWorkbookFilePath

' Initialize the wscript shell external object.
Set vWScriptShell = CreateObject("WScript.Shell")

' Determine the path to the main workbook file.
With CreateObject("Scripting.FileSystemObject")
	vMainWorkbookFilePath = .BuildPath(.GetParentFolderName(WScript.ScriptFullName), "App.xlsm")
End With

' Set the required environment variables.
With vWScriptShell.Environment("PROCESS")
	' Indicates that the project is to be run in background mode.
	If vIsBackgroundModeEnabled Then
		.Item("APP_IS_BACKGROUND_MODE_ENABLED") = "TRUE"
	End If
	' Stores the project name.
	.Item("APP_PROJECT_NAME") = vProjectName
End With

' Inialize a backup instance of the Excel application for other workbooks to use.
With CreateObject("Excel.Application")
	' Initialize an isolated instance of the Excel application and open the main workbook within it.
	With CreateObject("Excel.Application")
		' Check whether background mode is enabled.
		If Not vIsBackgroundModeEnabled Then
			' Make the application window visible and bring it to the forefront
			.Visible = True
			Call vWScriptShell.AppActivate(.Caption)
		End If

		' Open the main workbook file in read-only mode with the prepared password.
		Call .Workbooks.Open(vMainWorkbookFilePath, , True, , vMainWorkbookFilePassword)
	End With
End With
