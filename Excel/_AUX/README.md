## Auxiliary Functions

VBA is a verbose language. Business Users who record macros can create scripts 100s of lines long in a matter of minutes. This compounds to create steep barriers of entry for aspiring VBA developers. 

### Modulation with `Public`
The universal programming solution to trimming the verbose (among other things) is modulating your code. To do so in your Personal Macrobook, you'll likely want to start with an Auxiliary file. This auxiliary file can contain functions and subroutines that are accessible to other code. Inter-module development is possible by declaring the subroutine/function with the `Public` keyword. For instance, 

```vb
Public Sub RemoveAllSelectionBorders()
' Remove All Borders
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
```
![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/aux_funcs/RemoveBorders.png?raw=true)

### Calling Public Code

This function can then be called in other modules. For instance, I use a "Formatting." The below code references this module to Remove all Borders on the Selection and then change the font size to 14.

```vb
Sub LargeFont()
    ' Remove all borders & make font size 14
    RemoveAllSelectionBorders
    Selection.Font.Size = 14
End Sun
```

![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/aux_funcs/CallPublicExample.PNG?raw=true)
