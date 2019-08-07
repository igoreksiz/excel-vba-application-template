Option Explicit

' Declare local variables.
Dim vTaskName

' Define the current task name.
vTaskName = "UPDATE BASE: LOAD"

' Fetch the origin remote repository's branches.
Call ExecuteShell(vTaskName, "git fetch", False, False)

' Fetch the base remote repository's branches.
Call ExecuteShell(vTaskName, "git fetch base", False, False)

' Make sure we are in the project's master branch.
Call ExecuteShell(vTaskName, "git checkout master", False, False)

' Make a hard reset of the master branch to where the origin remote repository's master branch is pointing.
Call ExecuteShell(vTaskName, "git reset --hard origin/master", False, False)

' Overwrite the local repository's tags with those found in the origin remote repository.
Call ExecuteShell(vTaskName, "git fetch --tags -p", False, False)

' Execute garbage collection upon the repository.
Call ExecuteShell(vTaskName, "git gc", False, False)

' Report the task's success.
Call TaskSuccessNotification(vTaskName)
