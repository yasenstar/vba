# VBA - Return Values from Function

[TOC]

## Method 1: Return multiple values by using passing arguments By Reference

Passing argument By Reference is is probably the most common way to return multiple values from a function.

It uses the `ByRef` keyword to tell the compiler that the variable passed to the function is only a pointer to a memory location where the actual value of the variable is stored. This way, the code inside the function itself can modify the value of the variable. Even through the function does not explicitly return the changed value, it can be retrieved by using the same variable name within the code that calls the function.

By contrast, passing argument By Value `ByVal` instructs the function that the variable is read-only so the code inside the function can't change its value.

Note that in VBA and Visual Basic 6.0, if you do not specify `ByVal` or `ByRef` for a function or procedure argument, the default passing mechanism is `ByRef`. In Visual Basic .NET, the default behavior is passing arguments by Value.

It is a good coding practice for VBA and VB6 to include either the `ByVal` or `ByRef` keyword for each function argument (also called parameter).

My following sample is pop up folder picker dialog, let user to choose the destination folder, and return the folder path:

```vb
Function browseFolderPath(ByRef folderPath as String) As String
    On Error Goto ErrorHandler
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fileExplorer
        .Title = "Choose or Create a Folder for Save the Schedules"
        'To disable to multi select
        .AllowMultiSelect = False
        If .Show = -1 Then 'Any folder is selected
            folderPath = .SelectedItems(1)
        Else ' else dialog is calcelled
            'MsgBox "You have cancelled the folder picker"
            folderPath = "" 'when cancelled set blank as file path
        End If
    End With
    
ErrorHandler:
    Exit Function
End Function
```

Firstly, I put nothing in the argument part just as `browseFolderPath()`, it's not working, then I use `ByVal`, not working as well, only `ByRef` can really return the value needed.

Note: to utilize the function, you need to declare one global variable, like below

`Public Print_Path As String`

Then you can use `Call browseFolderPath(Print_Path)` in any Sub of the Excel coding.



------

Ref: http://www.geeksengine.com/article/vba-function-multiple-values.html, thanks!