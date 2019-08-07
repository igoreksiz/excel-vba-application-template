Option Explicit

' Declare global variables.
Dim vWScriptShell
Dim vFileSystemObject

' Initialize the wscript shell external object.
Set vWScriptShell = CreateObject("WScript.Shell")

' Initialize the file system object external object.
Set vFileSystemObject = CreateObject("Scripting.FileSystemObject")

' Define external object constants.
Const adTypeText = 2
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const fsoForReading = 1
Const vbext_ct_StdModule = 1
Const vbext_ct_ClassModule = 2
Const vbext_ct_MSForm = 3
Const xlMaximized = -4137
Const WshFinished = 1
Const xlOpenXMLWorkbookMacroEnabled = 52

Sub TaskNotification( _
	vTaskName, _
	vMessage _
)
	Call WScript.StdOut.WriteLine("-[ " & vTaskName & " ]- " & vMessage)
End Sub

Sub TaskSuccessNotification( _
	vTaskName _
)
	Call TaskNotification(vTaskName, "has ended successfully.")
End Sub

Function ExecuteShell( _
	vTaskName, _
	vCommand, _
	vIsStandardInputOutputReturned, _
	vIsContinueOnErrorEnabled _
)
	' Declare local variables.
	Dim vStandardInputText
	Dim vStandardOutputText
	Dim vStandardErrorText

	' Initialize the result dictionary.
	ExecuteShell = CreateObject("Scripting.Dictionary")

	' Initialize the execution of the given command.
	With vWScriptShell.Exec(vCommand)
		' Wait for the subprocess to end.
		Do While .Status <> WshFinished
			' If setup to do so, forward any of the recieved standard input to the subprocess.
			If _
				Not vIsStandardInputOutputReturned _
				And Not WScript.StdIn.AtEndOfStream _
			Then
				vStandardInputText = WScript.StdIn.ReadAll()
				If vStandardInputText <> vbNullString Then
					Call .StdIn.Write(vStandardInputText)
				End If
			End If

			' Give the subprocess time to execute.
			Call WScript.Sleep(100)
		Loop

		' Collect the standard output and standard error stream content.
		vStandardOutputText = .StdOut.ReadAll()
		vStandardErrorText = .StdErr.ReadAll()

		' Process the standard input output depending on the inheritance setting.
		If vIsStandardInputOutputReturned Then
			With ExecuteShell
				Call .Add("StandardOutput", vStandardOutputText)
				Call .Add("StandardError", vStandardErrorText)
			End With
		Else
			Call WScript.StdOut.Write(vStandardOutputText)
			Call WScript.StdErr.Write(vStandardErrorText)
		End If

		' Retrieve the exit code.
		Call ExecuteShell.Add("ExitCode", .ExitCode)

		' Process the exit code depending on the function's corresponding parameter.
		If Not vIsContinueOnErrorEnabled Then
			' Output all of the gathered information about the subprocess's execution.
			With WScript.StdErr
				Call .WriteLine("The child process exited with the status code: " & ExecuteShell("ExitCode"))
				Call .WriteLine("- command: " & vCommand)
				If vIsStandardInputOutputReturned Then
					Call .WriteLine("- standard output: " & ExecuteShell("StandardOutput"))
					Call .WriteLine("- standard error: " & ExecuteShell("StandardError"))
				End If
			End With

			' Report the task's failure and exit the current process.
			Call TaskNotification(vTaskName, "has failed.")
			Call WScript.Quit(-1)
		End If
	End With
End Function

Function ReadTextFile( _
	vFilePath _
)
	' Open the specified text file in ascii read mode, without creating it.
	With vFileSystemObject.OpenTextFile(vFilePath, fsoForReading, False)
		' Read and return all of the file's content.
		ReadTextFile = .ReadAll()

		' Close the text file.
		Call .Close
	End With
End Function

Sub WriteTextFile( _
	vFilePath, _
	vContent _
)
	' Create the specified text file in ascii read mode, without creating it.
	With vFileSystemObject.CreateTextFile(vFilePath)
		' Write all of the submitted content.
		Call .Write(vContent)

		' Close the text file.
		Call .Close
	End With
End Sub

Sub DownloadFileOverHttp( _
	vFilePath, _
	vMethod, _
	vUrl _
)
	' Declare local variables.
	Dim vXmlHttp

	' Create an instance of the xml http object.
	Set vXmlHttp = CreateObject("MSXML2.XMLHTTP.6.0")
	With vXmlHttp
		' Prepare the http request and send it.
		Call .open(vMethod, vUrl, False)
		Call .send
	End With

	' Write the file in binary mode to the specified path.
	With CreateObject("ADODB.Stream")
		.Type = adTypeBinary
		Call .Open
		Call .Write(vXmlHttp.responseBody)
		Call .SaveToFile(vFilePath, adSaveCreateOverWrite)
	End With
End Sub

Function GetLocalProjectDirectoryPath()
	With vFileSystemObject
		GetLocalProjectDirectoryPath = .GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName))
	End With
End Function

Function LoadDeployDirectoryPath( _
	vProjectDirectoryPath _
)
	' Declare local variables.
	Dim vTextFileContent

	' Load the file system object.
	With vFileSystemObject
		' Set the default result value.
		LoadDeployDirectoryPath = vbNullString

		' Load the contents of the deploy configuration file.
		vTextFileContent = ReadTextFile(.BuildPath(vProjectDirectoryPath, "Deploy.txt"))

		' Trim whitespace from the loaded contents.
		vTextFileContent = Trim(Replace(Replace(vTextFileContent, vbCr, vbNullString), vbLf, vbNullString))

		' Check the validity of the specified deploy directory path.
		If .FolderExists(vTextFileContent) Then
			' Set the result to be the contents of the deploy configuration file.
			LoadDeployDirectoryPath = vTextFileContent
		End If
	End With
End Function

Function LoadBuildConfiguration( _
	vFilePath _
)
	' Declare local variables.
	Dim vReferences()
	Dim vExternalModules()
	Dim vIndex
	Dim vItemNode
	Dim vItemNodes
	Dim vItem

	' Initialize the msxml dom document object.
	With CreateObject("MSXML2.DOMDocument.6.0")
		' Configure to load files asynchronously.
		.async = False

		' Load the build configuration xml file.
		Call .load(vFilePath)

		' Initialize the result.
		Set LoadBuildConfiguration = CreateObject("Scripting.Dictionary")

		' Load the root xml node.
		With .selectSingleNode("build")
			' Load the project's name.
			Call LoadBuildConfiguration.Add("ProjectName", .selectSingleNode("project-name").Text)

			' Load the flag that indicates whether background mode is enabled for the project.
			Call LoadBuildConfiguration.Add("IsBackgroundModeEnabled", .selectSingleNode("is-background-mode-enabled").Text = "True")

			' Load the required references.
			Set vItemNodes = .selectSingleNode("references").selectNodes("item")
			If vItemNodes.length = 0 Then
				Call LoadBuildConfiguration.Add("References", Null)
			Else
				Redim vReferences(vItemNodes.length - 1)
				vIndex = 0
				For Each vItemNode In vItemNodes
					Set vItem = CreateObject("Scripting.Dictionary")

					With vItemNode
						Call vItem.Add("GUID", .selectSingleNode("guid").Text)
						Call vItem.Add("Major", CLng(.selectSingleNode("major").Text))
						Call vItem.Add("Minor", CLng(.selectSingleNode("minor").Text))
					End With

					Set vReferences(vIndex) = vItem
					vIndex = vIndex + 1
				Next
				Call LoadBuildConfiguration.Add("References", vReferences)
			End If

			' Load the required external modules.
			Set vItemNodes = .selectSingleNode("external-modules").selectNodes("item")
			If vItemNodes.length = 0 Then
				Call LoadBuildConfiguration.Add("ExternalModules", Null)
			Else
				Redim vExternalModules(vItemNodes.length - 1)
				vIndex = 0
				For Each vItemNode In vItemNodes
					Set vItem = CreateObject("Scripting.Dictionary")

					With vItemNode
						Call vItem.Add("Name", .selectSingleNode("name").Text)
						Call vItem.Add("URL", .selectSingleNode("url").Text)
					End With

					Set vExternalModules(vIndex) = vItem
					vIndex = vIndex + 1
				Next
				Call LoadBuildConfiguration.Add("ExternalModules", vExternalModules)
			End If
		End With
	End With
End Function

Function GetFolderDateLastModified( _
	vFolder _
)
	' Declare local variables.
	Dim vSubFolder
	Dim vDateLastModified
	Dim vFile

	' Initialize the result to the smallest possible date value.
	GetFolderDateLastModified = DateSerial(100, 1, 1)

	' Recursively search for the greatest date last modified value among all contained subfolders.
	For Each vSubFolder In vFolder.SubFolders
		vDateLastModified = GetFolderDateLastModified(vSubFolder)
		If vDateLastModified > GetFolderDateLastModified Then
			GetFolderDateLastModified = vDateLastModified
		End If
	Next

	' Recursively search for the greatest date last modified value among all contained files.
	For Each vFile In vFolder.Files
		If vFile.DateLastModified > GetFolderDateLastModified Then
			GetFolderDateLastModified = vFile.DateLastModified
		End If
	Next
End Function

Function IsMainWorkbookOpen( _
	vProjectDirectoryPath _
)
	' The presence of the main workbook temporary file indicates that the main workbook is already open.
	With vFileSystemObject
		IsMainWorkbookOpen = .FileExists(.BuildPath(vProjectDirectoryPath, "~$App.xlsm"))
	End With
End Function

Function GetMainWorkbookFilePath( _
	vProjectDirectoryPath _
)
	GetMainWorkbookFilePath = vFileSystemObject.BuildPath(vProjectDirectoryPath, "App.xlsm")
End Function

Function GetMainWorkbookFilePassword( _
	vBuildConfiguration _
)
	' Declare local variable.
	Dim vBase64Node

	' Prepare a base64 XML node.
	Set vBase64Node = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
	vBase64Node.dataType = "bin.base64"

	' Use the stream api to encode the project's name
	With CreateObject("ADODB.Stream")
		.Type = adTypeText
		.CharSet = "us-ascii"
		Call .Open
		Call .WriteText("Pass^" & vBuildConfiguration("ProjectName") & "$ssaP")
		.Position = 0
		.Type = adTypeBinary
		.Position = 0
		vBase64Node.nodeTypedValue = .Read()
	End With

	' Return the processed string result.
	GetMainWorkbookFilePassword = vBase64Node.text
End Function

Sub CreateMainWorkbook( _
	vProjectDirectoryPath, _
	vBuildConfiguration _
)
	' Declare local variables.
	Dim vBootstrapFolder
	Dim vScriptFolder
	Dim vSourceFolder
	Dim vTestFolder
	Dim vMainWorkbookFilePath
	Dim vTempModuleFilePath
	Dim vItem
	Dim vModuleFile

	' Load the file system object.
	With vFileSystemObject
		' Load the bootstrap, script, source and test folder objects.
		Set vBootstrapFolder = .GetFolder(.BuildPath(vProjectDirectoryPath, "Bootstrap"))
		Set vScriptFolder = .GetFolder(.BuildPath(vProjectDirectoryPath, "Script"))
		Set vSourceFolder = .GetFolder(.BuildPath(vProjectDirectoryPath, "Source"))
		Set vTestFolder = .GetFolder(.BuildPath(vProjectDirectoryPath, "Test"))

		' Determine the main workbook file path.
		vMainWorkbookFilePath = GetMainWorkbookFilePath(vProjectDirectoryPath)

		' Check whether the main workbook file already exists.
		If .FileExists(vMainWorkbookFilePath) Then
			' Load the main workbook file object.
			With .GetFile(vMainWorkbookFilePath)
				' If the main workbook file is newer than any module file or this script, it doesn't need to be rebuilt.
				If ( _
					(GetFolderDateLastModified(vBootstrapFolder) < .DateLastModified) _
					And (GetFolderDateLastModified(vScriptFolder) < .DateLastModified) _
					And (GetFolderDateLastModified(vSourceFolder) < .DateLastModified) _
					And (GetFolderDateLastModified(vTestFolder) < .DateLastModified) _
				) Then
					Exit Sub
				End If
			End With

			' Remove the outdated main workbook file.
			Call .DeleteFile(vMainWorkbookFilePath, True)
		End If
	End With

	' Initialize an instance of the excel application.
	With CreateObject("Excel.Application")
		' Set the number of worksheets in new workbooks to one.
		.SheetsInNewWorkbook = 1

		' Create a new workbook.
		With .Workbooks.Add()
			' Rename the sheet of the main workbook.
			.Worksheets(1).Name = "ThisWorksheet"

			' Load the vbproject of the new workbook.
			With .VBProject
				' Add the references, defined in the build configuration, to the VBProject.
				If Not IsNull(vBuildConfiguration("References")) Then
					For Each vItem In vBuildConfiguration("References")
						Call .References.AddFromGuid(vItem("GUID"), vItem("Major"), vItem("Minor"))
					Next
				End If

				' Import the "Runtime" component.
				Call .VBComponents.Import(vFileSystemObject.BuildPath(vBootstrapFolder.Path, "Runtime.bas"))

				' Import the "ThisUserForm" component from a file.
				Call .VBComponents.Import(vFileSystemObject.BuildPath(vBootstrapFolder.Path, "ThisUserForm.frm"))
				Call .VBComponents("ThisUserForm").CodeModule.DeleteLines(1, 1)

				' Load the contents of the "ThisWorkbook" component from a file.
				With .VBComponents("ThisWorkbook").CodeModule
					Call .DeleteLines(1, 2)
					Call .AddFromFile(vFileSystemObject.BuildPath(vBootstrapFolder.Path, "ThisWorkbook.bas"))
				End With

				' Rename the sheet module of the vbproject.
				.VBComponents("Sheet1").Name = "ThisWorksheet"

				' Determine the temp folder path.
				With vFileSystemObject
					vTempModuleFilePath = .BuildPath(.BuildPath(vProjectDirectoryPath, "Temp"), "Module")
				End With

				' Download and import the external modules, defined in the build configuration.
				If Not IsNull(vBuildConfiguration("ExternalModules")) Then
					For Each vItem In vBuildConfiguration("ExternalModules")
						Call DownloadFileOverHttp(vTempModuleFilePath, "GET", vItem("URL"))
						Call .VBComponents.Import(vTempModuleFilePath)
						Call vFileSystemObject.DeleteFile(vTempModuleFilePath)
					Next
				End If

				' Import the source modules.
				For Each vModuleFile In vSourceFolder.Files
					Call .VBComponents.Import(vModuleFile.Path)
				Next

				' Import the test modules.
				For Each vModuleFile In vTestFolder.Files
					Call .VBComponents.Import(vModuleFile.Path)
				Next
			End With

			' Assign a shortcut key to the initialize macros.
			Call .Application.MacroOptions("ThisWorkbook.Initialize", , , , True, "q")
			Call .Application.MacroOptions("ThisWorkbook.Test", , , , True, "Q")

			' Save and password protect the main workbook file path.
			Call .SaveAs(vMainWorkbookFilePath, xlOpenXMLWorkbookMacroEnabled, GetMainWorkbookFilePassword(vBuildConfiguration))
		End With

		' Close the excel application instance.
		Call .Quit
	End With
End Sub

Sub FormatExportedModuleFile( _
	vFilePath _
)
	' Declare local variables.
	Dim vContent
	Dim vSpaceBeforNewlineSequence

	' Read the file content.
	vContent = ReadTextFile(vFilePath)

	' Make sure the file ends with one newline sequence.
	If Right(vContent, 2) = vbCrLf Then
		Do While Right(vContent, 4) = (vbCrLf & vbCrLf)
			vContent = Left(vContent, Len(vContent) - 2)
		Loop
	Else
		vContent = vContent & vbCrLf
	End If

	' Make sure there is no whitespace before a newline sequence.
	vSpaceBeforNewlineSequence = " " & vbCrLf
	Do While InStr(vContent, vSpaceBeforNewlineSequence) <> 0
		vContent = Replace(vContent, vSpaceBeforNewlineSequence, vbCrLf)
	Loop

	' Overwrite the file with the formatted content.
	Call WriteTextFile(vFilePath, vContent)
End Sub

Sub ExportMainWorkbookModules( _
	vProjectDirectoryPath, _
	vBuildConfiguration _
)
	' Declare local variables.
	Dim vSourceFolder
	Dim vTestFolder
	Dim vModuleFile
	Dim vExcludedModuleSet
	Dim vItem
	Dim vModuleFilePath
	Dim vVBComponent
	Dim vModuleFileDirectoryPath
	Dim vModuleFileExtension

	' Load the file system object.
	With vFileSystemObject
		' Determine the source and test folder paths.
		Set vSourceFolder = .GetFolder(.BuildPath(vProjectDirectoryPath, "Source"))
		Set vTestFolder = .GetFolder(.BuildPath(vProjectDirectoryPath, "Test"))

		' Remove all of the old source module files.
		For Each vModuleFile In vSourceFolder.Files
			Call vModuleFile.Delete
		Next

		' Remove all of the old test module files.
		For Each vModuleFile In vTestFolder.Files
			Call vModuleFile.Delete
		Next
	End With

	' Prepare a list of bootstrap and external modules that shall not be exported.
	Set vExcludedModuleSet = CreateObject("Scripting.Dictionary")
	With vExcludedModuleSet
		Call .Add("Runtime", Null)
		Call .Add("ThisUserForm", Null)
		Call .Add("ThisWorkbook", Null)
		Call .Add("ThisWorksheet", Null)
		For Each vItem In vBuildConfiguration("ExternalModules")
			Call .Add(vItem("Name"), Null)
		Next
	End With

	' Initialize an instance of the excel application.
	With CreateObject("Excel.Application")
		' Open the main workbook file.
		With .Workbooks.Open(GetMainWorkbookFilePath(vProjectDirectoryPath), , True, , GetMainWorkbookFilePassword(vBuildConfiguration))
			' Load the vbproject of the main workbook.
			With .VBProject
				' Export all of the VBProject's components.
				For Each vVBComponent In .VBComponents
					With vVBComponent
						' Check whether the current component is excluded.
						If Not vExcludedModuleSet.Exists(.Name) Then
							' Determine the module file's directory path.
							If Left(.Name, 4) = "Test" Then
								vModuleFileDirectoryPath = vTestFolder.Path
							Else
								vModuleFileDirectoryPath = vSourceFolder.Path
							End If

							' Determine the module file's extension.
							Select Case .Type
								Case vbext_ct_StdModule
									vModuleFileExtension = "bas"
								Case vbext_ct_ClassModule
									vModuleFileExtension = "cls"
								Case Else
									vModuleFileExtension = vbNullString
							End Select

							' Export the current component to the specfied module if an extension is specified.
							If vModuleFileExtension <> vbNullString Then
								vModuleFilePath = vFileSystemObject.BuildPath(vModuleFileDirectoryPath, .Name & "." & vModuleFileExtension)
								Call .Export(vModuleFilePath)
								Call FormatExportedModuleFile(vModuleFilePath)
							End If
						End If
					End With
				Next
			End With

			' Configure the workbook to ignore changes made.
			.Saved = True
		End With

		' Close the excel application instance.
		Call .Quit
	End With
End Sub

Sub CreateExecuteScript( _
	vProjectDirectoryPath, _
	vBuildConfiguration _
)
	' Declare local variables.
	Dim vTextFileContent

	' Load the file system object.
	With vFileSystemObject
		' Load the contents of the execute script template.
		vTextFileContent = ReadTextFile(.BuildPath(.BuildPath(vProjectDirectoryPath, "Script"), "_Execute.vbs"))

		' Set project specific constants.
		vTextFileContent = Replace(vTextFileContent, _
			"Const vProjectName = """"", _
			"Const vProjectName = """ & vBuildConfiguration("ProjectName") & """")
		vTextFileContent = Replace(vTextFileContent, _
			"Const vIsBackgroundModeEnabled = False", _
			"Const vIsBackgroundModeEnabled = " & CStr(vBuildConfiguration("IsBackgroundModeEnabled")))
		vTextFileContent = Replace(vTextFileContent, _
			"Const vMainWorkbookFilePassword = """"", _
			"Const vMainWorkbookFilePassword = """ & GetMainWorkbookFilePassword(vBuildConfiguration) & """")

		' Create the execute script file.
		Call WriteTextFile(.BuildPath(vProjectDirectoryPath, "Execute.vbs"), vTextFileContent)
	End With
End Sub

Sub ShowExcelApplication( _
	vExcelApplication _
)
	' Load the excel application instance.
	With vExcelApplication
		' Make the application window visible and bring it to the forefront
		.Visible = True
		Call vWScriptShell.AppActivate(.Caption)
	End With
End Sub

' Execute content of the file specified in the first argument.
Call ExecuteGlobal(ReadTextFile(WScript.Arguments(0)))
