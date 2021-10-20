Attribute VB_Name = "Clean"

Private Sub FindAndReplaceSelection(find_str As Variant, replace_str As String)
    ' Replace selection values.
    Selection.Replace What:=find_str, _
            Replacement:=replace_str, LookAt:=xlPart _
            , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

Private Sub ID_Empty()
' TEMP FIXES - REGEX WILL COUNT BLANK CELLS>
    FindAndReplaceSelection "", "EMPTY"
End Sub

Private Sub RM_Empty()
    FindAndReplaceSelection "EMPTY", ""
    FindAndReplaceSelection "Em", ""                ' m* will change empty.
End Sub



Private Sub CleanDigits()
    ' Only allow digits in range. Otherwise, remove. ALSO APPLIES TO STRING.
    ' Note: only select values to clean - otherwise, will crash

    ''''' Change Me <-------
    MIN_ALLOWED = 1
    MAX_ALLOWED = 4
    REPLACE_VAL = ""
    
    Debug.Print "MinVal: " & MIN_ALLOWED
    Debug.Print "MaxVal: " & MAX_ALLOWED
    Debug.Print "Replace: " & REPLACE_VAL
    
    Dim Rng As Range
    Dim WorkRng As Range
    Set WorkRng = Application.Selection
    
    On Error Resume Next
    For Each Rng In WorkRng
    
        Rng.Value = CInt(Rng.Value)         ' Convert values to integers
        
        If ( _
            Not IsNumeric(Rng.Value) Or _
            Rng.Value > MAX_ALLOWED Or _
            Rng.Value < MIN_ALLOWED _
        ) Then
            Rng.Value = REPLACE_VAL
        End If
    Next
End Sub



Private Sub FixAssessmentPeriod()
    ' Makes assessment periods uniform, e.g., initial, updated, DC, etc.
    ' Uses wildcards for cleaning. Must fill empty spaces before and after, else these are filled.
    ID_Empty        ' Identify & fill empty cells.

    Dim replacement_i As String, replacement_u As String, replacement_d As String, replacement_ad As String, _
        replacement_day30 As String, replacement_day60 As String
    replacement_i = "initial"
    replacement_u = "update"
    replacement_d = "discharge"
    replacement_ad = "admin_dc"
    
    replacement_day30 = "30 day"
    replacement_day60 = "60 day"
    
    '''''''''' Initial
    Dim Initial(4) As Variant
    Initial(0) = "ini*"
    Initial(1) = "intake"
    Initial(2) = "int*"
    Initial(3) = "adm*"
    
    '''''''''' Update
    Dim Update(6) As Variant
    Update(0) = "ann*"    'Annual
    Update(1) = "up*"
    Update(2) = "reass*"
    Update(3) = "re-ass*"
    Update(4) = "subs*"
    Update(5) = "six*"
    
    '''''''''' DISCHARGE
    Dim Discharge(5) As Variant
    Discharge(0) = "dc"
    Discharge(1) = "dis*"
    Discharge(2) = "dsc*"
    Discharge(3) = "fin*"
    Discharge(4) = "dic*"
    
    Dim AdminDischarge(1) As Variant
    AdminDischarge(0) = "admin dc*"
    
    '''''''''' 30 day
    Dim Day30(3) As Variant
    Day30(0) = "30"
    Day30(1) = "30 *"
    Day30(2) = "30Days"
    
    '''''''''' 60 day
    Dim Day60(3) As Variant
    Day60(0) = "60"
    Day60(1) = "60 *"
    Day60(2) = "60Days"
    
    
    '''''''''' Loops
    ' Loop through each array and replace all finds of that string with the replacement str
    For i = 0 To UBound(Initial) - LBound(Initial)
        FindAndReplaceSelection Initial(i), replacement_i
    Next
    
    For i = 0 To UBound(Update) - LBound(Update)
        FindAndReplaceSelection Update(i), replacement_u
    Next
    
    '' Admin DC BEFORE DC
    For i = 0 To UBound(AdminDischarge) - LBound(AdminDischarge)
        FindAndReplaceSelection AdminDischarge(i), replacement_ad
    Next
    
    For i = 0 To UBound(Discharge) - LBound(Discharge)
        FindAndReplaceSelection Discharge(i), replacement_d
    Next
    
    For i = 0 To UBound(Day30) - LBound(Day30)
        FindAndReplaceSelection Day30(i), replacement_day30
    Next
    
    For i = 0 To UBound(Day60) - LBound(Day60)
        FindAndReplaceSelection Day60(i), replacement_day60
    Next
    
    RM_Empty
End Sub


Private Sub makeLower()
    ' Makes Selection Lowercase. NOTE: Slow processing, only select needed range. Best with Table column selections.
    Dim cell As Range
    Set cell = Selection
    
    For Each cell In Selection.Cells
        If cell.HasFormula = False Then
            cell = LCase(cell)
        End If
    Next
End Sub


Sub Fix_YN()
    ' Makes Y/N rseponses uniform
    makeLower
    ID_Empty
    
    Dim str_y As String, str_n As String, str_neither As String
    str_y = "y"
    str_n = "n"
    str_neither = ""
    
    Dim a_y(2) As Variant, a_n(2) As Variant, a_neither(1) As Variant
    a_y(0) = "y*"
    a_y(1) = 1
    a_n(0) = "n*"
    a_n(1) = 0
    a_neither(0) = "y/n"
    
    ''''' Loops
    ' Neither goes first
    For i = 0 To UBound(a_neither) - LBound(a_neither)
        FindAndReplaceSelection a_neither(i), str_neither
    Next
    
    For i = 0 To UBound(a_n) - LBound(a_n)
        FindAndReplaceSelection a_n(i), str_n
    Next
    
    For i = 0 To UBound(a_y) - LBound(a_y)
        FindAndReplaceSelection a_y(i), str_y
    Next
    
    '' Remove all Not y/n
    Dim WorkRng As Range
    Set WorkRng = Application.Selection
    
    On Error Resume Next
    For Each Rng In WorkRng
        If (Rng.Value <> str_y And Rng.Value <> str_n) Then
            Rng.Value = str_neither
        End If
    Next
    
    RM_Empty
End Sub


Private Sub Clean_Gender()
    ' Cleans gender column. 
    ID_Empty
    
    '''''''''' Constants
    Dim replacement_m As String, replacement_f As String, replacement_o As String, replacement_t As String
    
    replacement_m = "m"
    replacement_f = "f"
    replacement_o = "o"
    replacement_t = "trans"
    
    '''''''''' Arrays
    Dim a_f(2) As Variant
    a_f(0) = "f*"
    a_f(1) = 2 
    
    Dim a_m(2) As Variant
    a_m(0) = "m*"
    a_m(1) = 1
    
    Dim a_t(1) As Variant
    a_t(0) = "trans*"
    
    Dim a_o(5) As Variant
    a_o(0) = "genderf*"
    a_o(1) = "n/a"
    a_o(2) = "nonbin*"
    a_o(3) = "non bin*"
    a_o(4) = "othe*"
    
    ''''''''''' Loops
    For i = 0 To UBound(a_f) - LBound(a_f)
        FindAndReplaceSelection a_f(i), replacement_f
    Next
    For i = 0 To UBound(a_m) - LBound(a_m)
        FindAndReplaceSelection a_m(i), replacement_m
    Next
    For i = 0 To UBound(a_o) - LBound(a_o)
        FindAndReplaceSelection a_o(i), replacement_o           ' Other regex must go after - e.g., "gender_F_luid
    Next
    For i = 0 To UBound(a_t) - LBound(a_t)
        FindAndReplaceSelection a_t(i), replacement_t
    Next
    
    RM_Empty
End Sub

Private Sub Clean_Language()
    ''''' Clean Languages
    ID_Empty
    
    '''''''''' Constants
    Dim replacement_eng As String, replacement_span As String, replacement_oth As String
    replacement_eng = "english"
    replacement_span = "spanish"
    replacement_oth = "other"
    
    '''''''''' Arrays
    Dim a_eng(1) As Variant, a_span(1) As Variant, a_other(1) As Variant
    a_other(0) = "*/*"
    a_eng(0) = "eng*"
    a_span(0) = "spa*"
    
    '''''''''' Loops
    For i = 0 To UBound(a_other) - LBound(a_other)
        FindAndReplaceSelection a_other(i), replacement_oth
    Next
    For i = 0 To UBound(a_eng) - LBound(a_eng)
        FindAndReplaceSelection a_eng(i), replacement_eng
    Next
    For i = 0 To UBound(a_span) - LBound(a_span)
        FindAndReplaceSelection a_span(i), replacement_span
    Next
    
    RM_Empty
End Sub


Private Sub Clean_Ethnicity()
    ' Clean race & ethnicity.
    ' DOUBLE CHECK CODES BEFORE RUNNING
    ' Start with mixed beforehand. Order matters.
    ID_Empty
    
    '''''''''' Replacement Strings
    Dim str_ang As String, str_afr As String, str_asi As String, str_haw As String, str_lat As String, str_nat As String, str_other As String, str_mixed As String
    str_afr = "afr_amer"
    str_ang = "ang_amer"
    str_asi = "asi_amer"
    str_haw = "haw_pac"
    str_lat = "lat_amer"
    str_nat = "nat_amer"
    str_other = "other"
    str_mixed = "mixed"
    
    
    ''''' Arrays
    Dim a_afr(2) As Variant, a_ang(7) As Variant, a_asi(6) As Variant, a_haw(1) As Variant, a_lat(3) As Variant, a_mixed(3) As Variant, a_nat(2) As Variant, a_other(3) As Variant
    ''''' Mixed
    a_mixed(0) = "*/*"
    a_mixed(1) = "mix*"
    a_mixed(2) = "multi*"
    ''''' Other
    a_other(0) = "oth*"
    a_other(1) = "*matter*"
    a_other(2) = "^n$"
    ''''' Latin
    a_lat(0) = "lat*"
    a_lat(1) = "hisp*"
    a_lat(2) = "mex*"
    
    ''''' Native American
    a_nat(0) = "nat*"
    a_nat(1) = "amerin*"
    ''''' Hawaain
    a_haw(0) = "haw*"
    ''''' Asian
    a_asi(0) = "asi*"
    a_asi(1) = "hmon*"
    a_asi(2) = "camb*"
    a_asi(3) = "filip*"
    a_asi(4) = "fliip*"
    a_asi(5) = "flil*"
    
    ''''' African
    a_afr(0) = "afr*"
    a_afr(1) = "black*"
    ''''' Anglish
    a_ang(0) = "ang*"
    a_ang(1) = "cauc*"
    a_ang(2) = "russian*"
    a_ang(3) = "white*"
    a_ang(4) = "irish*"
    a_ang(5) = "german*"
    a_ang(6) = "norwegian*"
    
        
    '''''''''' Loops
    For i = 0 To UBound(a_mixed) - LBound(a_mixed)
        FindAndReplaceSelection a_mixed(i), str_mixed
    Next
    For i = 0 To UBound(a_other) - LBound(a_other)
        FindAndReplaceSelection a_other(i), str_other
    Next
    For i = 0 To UBound(a_lat) - LBound(a_lat)
        FindAndReplaceSelection a_lat(i), str_lat
    Next
    For i = 0 To UBound(a_nat) - LBound(a_nat)
        FindAndReplaceSelection a_nat(i), str_lat
    Next
    For i = 0 To UBound(a_haw) - LBound(a_haw)
        FindAndReplaceSelection a_haw(i), str_haw
    Next
    For i = 0 To UBound(a_asi) - LBound(a_asi)
        FindAndReplaceSelection a_asi(i), str_asi
    Next
    For i = 0 To UBound(a_afr) - LBound(a_afr)
        FindAndReplaceSelection a_afr(i), str_afr
    Next
    For i = 0 To UBound(a_ang) - LBound(a_ang)
        FindAndReplaceSelection a_ang(i), str_ang
    Next
    
    RM_Empty
End Sub


Sub ParseMultipleValuesIntoColumns()
    ' Parses multiple numeric values in a column into separate columns.
    ' Creates NUM_VALS aux columns, then fills them with the number checking in original column.
    
    ' NOTES:
        ' Starting value must be 1
        ' Only accepts values under 10
    
    NUM_VALS = 4            ' <- CHANGE ME TO NUM POSSIBLE VALUES
    Debug.Print "Creating " & NUM_VALS & " columns"
    
    '''''''''' Var Declaration
    Dim NumRows As Integer
    Dim ColName As String
    
    NumRows = Selection.Rows.Count
    Selection.End(xlUp).Select
    ColName = Selection.Value
    
    For i = 1 To NUM_VALS
        Selection.EntireColumn.Offset(0, 1).Insert      ' Insert to right of column
        ActiveCell.Offset(0, 1).Select
        Selection.Value = ColName & "_" & i
    Next
    
    ActiveCell.Offset(0, -NUM_VALS).Select  ' Return to starting column
    
    For i = 1 To NumRows
        ActiveCell.Offset(1, 0).Select
        For j = 1 To NUM_VALS
            If (ActiveCell.Value = j) Or (InStr(ActiveCell.Value, j) > 0) Then
                ActiveCell.Offset(0, j).Value = j
            End If
            
        Next
    Next
End Sub

