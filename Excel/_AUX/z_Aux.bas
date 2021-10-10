Attribute VB_Name = "z_Aux"
Public Sub RemoveAllSelectionBorders()
' Remove All Borders from Selection
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub


Public Sub ApplyBottomBorder()
    ' Apply bottom border to selection
    With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
    End With
End Sub

Public Sub MoveDownEnterValue(Value As String)
    ' Move down 1 & enter value
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = Value
End Sub

Public Sub MoveRightEnterValue(Value As String)
    ' Move Right & enter Value
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = Value
End Sub
