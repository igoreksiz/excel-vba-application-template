Option Explicit

' Declare local variables.
Dim vTaskName
Dim vProjectDirectoryPath
Dim vDeployDirectoryPath
Dim vBuildConfiguration

' Define the current task name.
vTaskName = "DEBUG"

' Retrieve the project's directory path.
vProjectDirectoryPath = GetLocalProjectDirectoryPath()

' Load the deploy directory path.
vDeployDirectoryPath = LoadDeployDirectoryPath(vProjectDirectoryPath)
If vDeployDirectoryPath = vbNullString Then
	Call TaskNotification(vTaskName, "cannot find the 'Deploy.txt' file in the project directory containing a valid directory path.")
	Call WScript.Quit(-1)
End If

' Load the build configuration from the build configuration xml document.
Set vBuildConfiguration = LoadBuildConfiguration(vFileSystemObject.BuildPath(vProjectDirectoryPath, "Build.xml"))

' Set the required environment variables.
With vWScriptShell.Environment("PROCESS")
	' Indicates that the project is to be run in debug mode.
	.Item("APP_IS_DEBUG_MODE_ENABLED") = "TRUE"
	' Indicates that the project is to be run in background mode.
	If vBuildConfiguration("IsBackgroundModeEnabled") Then
		.Item("APP_IS_BACKGROUND_MODE_ENABLED") = "TRUE"
	End If
	' Indicates that the project is to be run in deploy debug mode.
	.Item("APP_IS_DEPLOY_DEBUG_MODE_ENABLED") = "TRUE"
	' Stores the project name.
	.Item("APP_PROJECT_NAME") = "[Debug] " & vBuildConfiguration("ProjectName")
End With


' Inialize a backup instance of the Excel application for other workbooks to use.
With CreateObject("Excel.Application")
	' Open the project's main workbook in debug mode.
	With CreateObject("Excel.Application")
		' Display the application window.
		Call ShowExcelApplication(.Application)

		' Open the main workbook file the prepared password.
		Call .Workbooks.Open(GetMainWorkbookFilePath(vDeployDirectoryPath), , True, , GetMainWorkbookFilePassword(vBuildConfiguration))

		' Wait for the main workbook to be closed.
		Do While .Workbooks.Count > 0
			Call WScript.Sleep(1000)
		Loop
	End With
End With

' Report the task's success.
Call TaskSuccessNotification(vTaskName)
