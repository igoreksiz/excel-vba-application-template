Option Explicit

' Declare local variables.
Dim vTaskName
Dim vProjectDirectoryPath
Dim vBuildConfiguration

' Define the current task name.
vTaskName = "BUILD"

' Retrieve the project's directory path.
vProjectDirectoryPath = GetLocalProjectDirectoryPath()

' If the main workbook is already open, notify the user and exit.
If IsMainWorkbookOpen(vProjectDirectoryPath) Then
	Call TaskNotification(vTaskName, "the main workbook is already open in a different process and must be closed before proceeding.")
	Call WScript.Quit(-1)
End If

' Load the build configuration from the build configuration xml document.
Set vBuildConfiguration = LoadBuildConfiguration(vFileSystemObject.BuildPath(vProjectDirectoryPath, "Build.xml"))

' Create the main workbook.
Call CreateMainWorkbook(vProjectDirectoryPath, vBuildConfiguration)

' Create the execute script.
Call CreateExecuteScript(vProjectDirectoryPath, vBuildConfiguration)

' Report the task's success.
Call TaskSuccessNotification(vTaskName)
