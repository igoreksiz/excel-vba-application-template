# Excel VBA Application Template

[![License](https://img.shields.io/github/license/Player1os/excel-vba-application-template.svg)](https://github.com/Player1os/excel-vba-application-template/blob/master/LICENSE)

A template project from which Excel applications written in Visual Basic for Applications can be derived.

The use cases range from simple scripts that run in the background to interactive GUI applications.

Such applications are suitable in cases where your platform is restricted in terms of using:
- development tools
- popular third party runtimes (such as Python, Java, NodeJS, etc.)
- compiling and running custom executables

## Requirements

The target platform must be a version of Microsoft Windows 7 or newer and must have a version of Microsoft Office 2008 or newer installed.

The GUI is written in HTML and rendered by an embedded Internet Explorer instance with client side javascript disabled. All interactivity
is handled by clicking on anchor elements and changing the fragment portion of the URL.

## Creating a new project

Follow these instructions to create a new project based on this template:

1. Create a new repository for the project and clone it locally, by running:

	```
	git clone REPOSITORY_URL
	```

1. Add the template's `origin` remote to the newly created project as an additional remote repository named `base`, by running:

	```
	git remote add base git@github.com:Player1os/excel-vba-application-template.git
	```

	Alternatively, we can use the HTTPS url, if an ssh connection cannot be established, by running:

	```
	git remote add base https://github.com/Player1os/excel-vba-application-template.git
	```

1. Pull the latest commits of the `master` branch from the newly set `base` remote, by running:

	```
	git pull base master
	```

1. Make sure to change the default values found in the following files:

	- `Build.xml`
		- `project-name`
		- `is-background-mode-enabled`
	- `README.md`
	- `LICENSE`

1. If needed, create a `Deploy.txt` file containing the path to an existing directory (this indicates the location, where the production
version of the application is to be deployed).

1. Apply any other changes required by the new project. When done, commit them, by running:

	```
	git add .
	git commit -m "Transform the template boilerplate into the the project's initial form"
	```

1. Push the local `master` branch and set it to track the remote `origin/master` branch, by running:

	```
	git push -u origin master
	```

## Creating a copy of an existing project

Assuming we have already [created a new project](#creating-a-new-project) as described above, we can create a new copy of the project
(on a different machine, for instance), by following these steps:

1. Create a local clone from the `origin` remote repository, by running:

	```
	git clone REPOSITORY_URL
	```

1. Add the template's `origin` remote to the project as an additional remote repository named `base`, by running:

	```
	git remote add base git@github.com:Player1os/excel-vba-application-template.git
	```

	Alternatively, we can use the HTTPS url, if an ssh connection cannot be established, by running:

	```
	git remote add base https://github.com/Player1os/excel-vba-application-template.git
	```

1. Fetch the newest version of the `base` repository's branches, by running:

	```
	git fetch base
	```

	This should only download the `master` branch of the `base` repository.

## Implementing a change to the project

Whether we are adding a feature or fixing a bug, it is recommended to follow these steps while doing so:

1. Checkout the `master` branch of the project, by running:

	```
	git checkout master
	```

1. Create and checkout a new, appropriately named topic branch, by running:

	```
	git checkout -b XYZ
	```

1. Apply the desired changes into the newly created topic branch. If any changes occurred in the `master` branch while we were still
working on the topic branch, pull these changes and rebase the topic branch, by running:

	```
	git checkout master
	git pull
	git checkout XYZ
	git rebase master
	```

	It is recommended to do this as soon as possible when the `master` branch is updated to minimize the amount of conflicts that need
	to be addressed at any given time.

1. When done, make sure all changes to the topic branch are pushed to the remote `origin` repository, by running:

	```
	git push -u origin XYZ
	```

1. Optionally, we may want to *squash* or otherwise reorganize the branch's commits before publishing them, thus producing a more
readable commit history, by running:

	```
	git rebase -i master
	```

	It is recommended to check the repository's *git history* after the interactive rebase operation, to make sure we've achieved
	the desired changes.

1. Using the remote `origin` repository's interface, create a **pull request** from the topic branch back to the `master` branch.

1. Since the topic branch has already been rebased to the newest version of the `master` branch, there will be no conflicts and the
topic branch can be simply merged by fast-forwarding the `master` branch.

1. Delete the local topic branch after checking out and pulling the newest version of the `master` branch, by running:

	```
	git checkout master
	git pull
	git branch -d XYZ
	```

1. If the `origin` remote repository has already deleted its topic branch during the **pull request** operation, the local remote refs
can simply be fetched and pruned, by running:

	```
	git fetch -p
	```

	Otherwise, we need to remove the `origin` remote topic branch manually, by running:

	```
	git push -d origin XYZ
	```

## Updating the project with changes from the template

We can apply any changes introduced to the template after a project has been derived from it, by running:

```
Task\UpdateBase\Create.bat
```

This rebases the `master` branch of the project's repository to the newest commit of the template's `master` branch, creating a new tree.
This new tree is then pushed to to the `origin` remote repository. The script also ensures all of the repository's tags remain unchanged
during this process and any newly introduced dependencies are installed.

As is the case with any rebase operation, it is recommended that we check the current *git history*, to make sure the update has been
applied correctly.

Optionally, we may wish to [update the project's version](#update-the-project%27s-version) as explained above.

If the project is being developed on another machine, the local repository on it will need to be updated, by running:

```
Task\UpdateBase\Load.bat
```

This overwrites the local repository tree with the contents of the newly rebased tree located in the `origin` remote repository and
ensures any newly introduced dependencies are installed.

---

## Update the project's version

At some point (usually after [implementing one or more changes to the project](#implementing-a-change-to-the-project), as described above)
we may want to update the project's version, by running:

```
Task\Version.bat [patch | minor | major]
```

> TODO.
The corresponding `preversion` and `postversion` scripts defined in the `package.json` file ensure that the *new version* value is stored
in all the relevant locations, which include:

- Modifying the `package.json` and `package-lock.json` files.
- Tagging in the latest `master` branch commit.

## Contents

> TODO:
- execute script file name
- main workbook file name
- password protection of VBAProject in main workbook
- password protection of config workbook

### VBA Project

- Document worksheet naming convention (in code based on normal)

### External references

The contained modules and classes require the VBA project to have references to the following standard libraries:

- **Microsoft Scripting Runtime**
- **Microsoft Visual Basic for Applications Extensibility 5.3**

## Alleged garbage collection bug

There is some uncertainty among the VBA community as to whether the reference counter is correct implemented within the interpreter. The assumption is that the references on objects are not decremented when an object variable (a reference) goes out of scope. The solution to this problem would be to manually ensure the clearing of every such locally declared variable as to avoid memory leaking.

For more information please refer to the following discussions:
- https://stackoverflow.com/questions/517006/is-there-a-need-to-set-objects-to-nothing-inside-vba-functions
- https://stackoverflow.com/questions/19038350/when-should-an-excel-vba-variable-be-killed-or-set-to-nothing
- https://www.mrexcel.com/forum/excel-questions/619445-vba-setting-objects-nothing.html

Although I have not tested all object types and Microsoft Excel versions, I can confirm the following about the VBA interpreter found in Microsoft Excel 2013:
- It correctly deallocates dynamically created objects (including those nested within other objects) as soon as all variables that reference them cease to exist, without the need to manually clear said variables.
- It does not handle circular references correctly, i.e. if two objects reference each other internally and all variables that reference them cease to exists, they will not be deallocated unless one of the internal references is cleared manually. This naturally also applies in more general circular referencing cases with more objects involved.

From the above, I believe that manual clearing of variables is required only when dealing with specific ADO objects. That is why all procedures that deal with said ADO objects contain a section that takes care of manually clearing all declared variables that reference them at the end of their lifetime, even if an unhandled error was encountered within the procedure.

This can also be a good way to showcase the error handling boilerplate code found in the `Module.bas` module, which attempts to emulate the functionality of the modern error handling syntax found in other languages:
```javascript
try {
	// ...
} catch (exception) {
	// ...
} finally {
	// ...
}
```

> TODO: Add IDE Options (color, syntax check, tabs, error handling method, ...).
> TODO: Describe how to alter the build xml file (adding references, importing external modules).
> TODO: Implementation of tests.
