Option Explicit

' Declare local variables.
Dim vTaskName
Dim vProjectDirectoryPath
Dim vDeployDirectoryPath
Dim vBuildConfiguration
Dim vDeployAssetDirectoryPath

' Define the current task name.
vTaskName = "DEPLOY"

' Retrieve the project's directory path.
vProjectDirectoryPath = GetLocalProjectDirectoryPath()

' Load the deploy directory path.
vDeployDirectoryPath = LoadDeployDirectoryPath(vProjectDirectoryPath)
If vDeployDirectoryPath = vbNullString Then
	Call TaskNotification(vTaskName, "cannot find the 'Deploy.txt' file in the project directory containing a valid directory path.")
	Call WScript.Quit(-1)
End If

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

With vFileSystemObject
	' Copy the main workbook file and execute script file to the deploy directory.
	Call .CopyFile(.BuildPath(vProjectDirectoryPath, "App.xlsm"), .BuildPath(vDeployDirectoryPath, "App.xlsm"), True)
	Call .CopyFile(.BuildPath(vProjectDirectoryPath, "Execute.vbs"), .BuildPath(vDeployDirectoryPath, "Execute.vbs"), True)

	' Copy the assets directory to the deploy directory.
	vDeployAssetDirectoryPath = .BuildPath(vDeployDirectoryPath, "Assets")
	If .FolderExists(vDeployAssetDirectoryPath) Then
		Call .DeleteFolder(vDeployAssetDirectoryPath)
	End If
	Call .CopyFolder(.BuildPath(vProjectDirectoryPath, "Assets"), vDeployAssetDirectoryPath, True)
End With

' Report the task's success.
Call TaskSuccessNotification(vTaskName)
