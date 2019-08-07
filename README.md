# Excel VBA Application Template

[![License](https://img.shields.io/github/license/Player1os/excel-vba-application-template.svg)](https://github.com/Player1os/excel-vba-application-template/blob/master/LICENSE)

A template project from which Excel applications written in Visual Basic for Applications can be derived.

## TODO:

### Naming convention:
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

>> TODO: Add IDE Options (color, syntax check, tabs, error handling method, ...)
>> TODO: Assume the presence of Outlook, Internet Explorer, ...

This repository contains a framework for creating MS Excel VBA applications, that simplifies the following tasks:

- Native and UTF-8 file I/O
- JSON serialization and de-serialization
- Base64 serialization and de-serialization
- Database communication through *ActiveX Data Objects*
- Sending EMail (if MS Outlook is available)
- Sending HTTP requests
- Helper functions that are not part of the standard API for the following classes:
	- Workbook
	- Worksheet
	- Range
	- UserForm
- Abstracted processing of a cell range as a table with headers
- General data-types such as:
	- List (Collection)
	- Dictionary
	- Set
- Displaying progress for long running processes
- Disabling application event handling for increased performance
- Automatic project unlocking in debug mode for password protected projects
- Automatic extraction of all project files from the workbook
- Error handling and reporting
- Deploying the application to a remote location as a read only application
- Specific behavior when running in debug mode
- Organization of the application's code into:
	- Event handlers
	- Routed controllers (Actions)
	- Application state
	- Utility modules and classes

## Third party code

Some of the modules contain code that was not implemented by me. The original authors are credited in all of those instances. If you see your code here and I didn't credit you, please let me know.

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

This project also contains a `CUtilityRecordset` class, which acts as a wrapper for the `ADO.Recordset` class, since that is the only class I found to be useful in user code outside of the provided methods to interact with an external database. The wrapper has one task and that is internally holding, transparently providing and at the end of its lifetime manually clearing an internal instance of the `ADO.Recordset` class.

As for the circular referencing bug, I tend to use simple objects with no circular references whenever possible to avoid it.

>> TODO: Describe how to alter the build xml file
