Option Explicit

' Declare local variables.
Dim vExecuteScriptFilePath

' Determine th path to the execute script file.
With CreateObject("Scripting.FileSystemObject")
	vExecuteScriptFilePath = .BuildPath(.GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName)), "Execute.vbs")
End With

' Load the wscript shell object.
With CreateObject("WScript.Shell")
	' Set the navigate path environment variable to the user's input.
	.Environment("PROCESS")("APP_NAVIGATE_PATH") = InputBox("Enter the path to navigate to")

	' Run the execute script.
	Call .Run(vExecuteScriptFilePath, 0, False)
End With
