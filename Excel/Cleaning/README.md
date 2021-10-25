# Cleaning Scripts

## Purpose
I worked with a Behavioral Health organization across the state. A historically underfunded field, it was common practice for data entry to take place in an Excel workbook with no data validation restrictions. Each data set requires extensive cleaning and formatting before uploading it to the centralized database. With irregular formats, it makes sense to have these cleaning scripts easily available with fluid reference structures via Personal Macrobooks. These macros are similar yet independent of the ones I developed in the workplace to solve this challenge.

## String Cleaning Methodology
Please note that these subs are defined with the `Private` keyword and thus only accessible within the module. The macros are only available inside the module and must be run intentionally. These macros start with the `CleanString_` Suffix.

### Array replacement
The string cleaning macros iterates through an array to clean strings. Below is a generic method that's repeated across the string cleaning macros.
```vb
Dim CorrectWord As String
CorrectWord = "One"

Dim WordArr As Variant(2)
WordArr(0) = "Ones"
WordArr(1) = "Onee"

For i = 0 To UBound(WordArr) - LBound(WordArr)
    FindAndReplaceSelection WordArr(i), CorrectWord
Next
```

### Filling empty space
The string cleaning macros use Excel wildcards that prioritize recall (over precision). These wildcards will match empty cells. To circumvent this, the string cleaning macros call two auxiliary macros:

#### ID_Empty
Identifies Empty cells in the Selection and fills them with a generic "EMPTY" String.

#### RM_EMPTY
Removes Generic Empty String from cells and restores cells to original NULL state.


## Macro Efficient
Please note that these macros are not calculation efficient. They will take long times or crash if you select the entire column. In general, best practice for VBA macros is to only select the intended data.


## Macro showcase
Macro Demonstration.

### CleanString_AssessmentPeriod
> Select relevant column and run

Useful for cleaning longitudinal survey instruments. 
![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/Excel/Cleaning/Clean_Period.gif?raw=true)



### CleanString_YN
> Select relevant column and run

### CleanString_Gender
> Select relevant column and run

### CleanString_Language
> Select relevant column and run

### CleanString_Ethnicity
> Select relevant column and run

### Clean_Digits
> Select relevant column and run

### ParseMultipleValuesIntoColumns
> Select relevant column and run






