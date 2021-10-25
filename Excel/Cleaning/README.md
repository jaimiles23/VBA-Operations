# Cleaning Scripts
- [Cleaning Scripts](#cleaning-scripts)
  - [Purpose](#purpose)
  - [String Cleaning Methodology](#string-cleaning-methodology)
    - [Array replacement](#array-replacement)
    - [Filling empty space](#filling-empty-space)
      - [ID_Empty](#id_empty)
      - [RM_EMPTY](#rm_empty)
  - [Macro Efficient](#macro-efficient)
  - [Macro showcase](#macro-showcase)
    - [CleanString_AssessmentPeriod](#cleanstring_assessmentperiod)
    - [CleanString_YN](#cleanstring_yn)
    - [CleanString_Gender](#cleanstring_gender)
    - [CleanString_Language](#cleanstring_language)
    - [CleanString_Ethnicity](#cleanstring_ethnicity)
    - [Clean_Digits](#clean_digits)
    - [ParseMultipleValuesIntoColumns](#parsemultiplevaluesintocolumns)

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

Clean Yes/No column with different codings & misspellings.

![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/Excel/Cleaning/Clean_YesNo.gif?raw=true)


### CleanString_Gender
> Select relevant column and run

Clean gender column with misspellings. Assumes gender binary and other.

![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/Excel/Cleaning/Clean_Gender.gif?raw=true)

### CleanString_Language
> Select relevant column and run

![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/Excel/Cleaning/Clean_Language.gif?raw=true)

### CleanString_Ethnicity
> Select relevant column and run

Create uniform ethnicity for analysis.

![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/Excel/Cleaning/Clean_Ethnicity.gif?raw=true)

### Clean_Digits
> Select relevant column and run

Clean column to only accept digits within range. Example below only allows values between 1 and 4 (inclusive).

![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/Excel/Cleaning/Clean_Digits.gif?raw=true)

### ParseMultipleValuesIntoColumns
> Select relevant column and run

Parse multiple values into a column into separate columns. Ignores delimiter and instead assumes numeric codings, from 1 to 10. Must specify number of values to search.


![.](https://github.com/jaimiles23/VBA-Operations/blob/main/_images/Excel/Cleaning/ParseValues.gif?raw=true)
