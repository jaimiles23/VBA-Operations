Attribute VB_Name = "Format"
Sub format_section()
Attribute format_section.VB_ProcData.VB_Invoke_Func = "S\n14"
    ' Format selection to Section Style
    ' Keyboard Shortcut: Ctrl+Shift+s
    RemoveAllSelectionBorders
    ApplyBottomBorder
    
    With Selection
        With .Font
            .Size = 14
            .ThemeColor = xlThemeColorDark1
        End With
        
        .Interior.Color = 8388608
    End With
End Sub

Sub format_sort_ascending()
Attribute format_sort_ascending.VB_ProcData.VB_Invoke_Func = "H\n14"
    ' Format selection to Header Style
    ' Keyboard Shortcut: Ctrl+Shift+h
    RemoveAllSelectionBorders
    ApplyBottomBorder
    
    With Selection
        With .Font          ' Font
            .Size = 13
        End With
        
        With .Interior      ' Shade
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
        End With
    End With
End Sub

Sub format_alphabetize()
Attribute format_alphabetize.VB_Description = "="
Attribute format_alphabetize.VB_ProcData.VB_Invoke_Func = "A\n14"
' alphabetize selection.
' Keyboard Shortcut: Ctrl+Shift+A
With ActiveSheet.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Selection.Columns(1), Order:=xlAscending
        'the key you want to use is the column to sort on. I used column 1, which is "A", column "B" is 2, etc
    .SetRange Selection
    .Apply
End With
End Sub

Sub format_red()
Attribute format_red.VB_ProcData.VB_Invoke_Func = "r\n14"
    ' Fill Selection Red. Useful for calling attention to field for editing
    ' Keyboard Shortcut: Ctrl+r
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
    End With
End Sub

