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
	Call ReportNotification(vTaskName, "follow the rebase instructions and when done enter the `exit` command, to continue.")

	' Execute the nested shell.
	Call ExecuteShell(vTaskName, "cmd", False, False)
Loop

' Execute garbage collection upon the repository.
Call ExecuteShell(vTaskName, "git gc", False, False)

' Propagate the changes to the origin remote repository.
Call ExecuteShell(vTaskName, "git push -f", False, False)
Call ExecuteShell(vTaskName, "git push --tags -f", False, False)

' Report the task's success.
Call TaskSuccessNotification(vTaskName)
