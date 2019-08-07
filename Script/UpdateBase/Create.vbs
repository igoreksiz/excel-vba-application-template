Option Explicit

' Declare local variables.
Dim vTaskName

' Define the current task name.
vTaskName = "UPDATE BASE: CREATE"

' Fetch the base remote repository's branches.
Call ExecuteShell(vTaskName, "git fetch base", False, False)

' Make sure we are in the project's master branch.
Call ExecuteShell(vTaskName, "git checkout master", False, False)

' Rebase the project's master branch to the base remote repository's master branch.
Call ExecuteShell(vTaskName, "git rebase base/master", False, True)

' Loop until the rebase is finished or aborted.
Do While Left(ExecuteShell(vTaskName, "git status", True, False)("StandardOutput"), 18) = "rebase in progress"
	' Notify the user of the instructions to follow.
	Call TaskNotification(vTaskName, "open a new command prompt and follow the rebase instructions.")
	Call TaskNotification(vTaskName, "when the rebase is done, close the new command prompt and press ENTER.")

	' Wait for the enter key to be pressed.
	Call WScript.StdIn.ReadLine()
Loop

' Execute garbage collection upon the repository.
Call ExecuteShell(vTaskName, "git gc", False, False)

' Propagate the changes to the origin remote repository.
Call ExecuteShell(vTaskName, "git push -f", False, False)
Call ExecuteShell(vTaskName, "git push --tags -f", False, False)

' Report the task's success.
Call TaskSuccessNotification(vTaskName)
