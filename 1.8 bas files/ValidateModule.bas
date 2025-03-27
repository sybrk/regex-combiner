Attribute VB_Name = "ValidateModule"
Function Validate() As String
'
' validate Makro
'

'

    'Regex pattern validation
    Dim regexRange As Range
    Dim regexCell As Range
    Dim validRegex As Boolean
    Dim isLookBehind As Boolean
    Dim foundRegexErrors As Boolean
    Dim foundLookBehind As Boolean
    Dim patternToCheck As String
    Set regexRange = Range("G7:H" & Range("H" & Rows.Count).End(xlUp).Row)
    
    'clear fill colors before finding new errors.
    For Each regexCell In regexRange
    
    regexCell.Interior.pattern = xlNone
          
    Next regexCell
 
    foundRegexErrors = False
    foundLookBehind = False
    'Loop through each cell in the regex columns
    For Each regexCell In regexRange
    patternToCheck = regexCell.Value
    
    validRegex = IsValidRegex(patternToCheck)
    isLookBehind = ContainsLookBehind(patternToCheck)
    
    If isLookBehind Then
     Debug.Print "Pattern " & regexCell.Value & " is lookbehind."
     patternToCheck = ReplaceLookBehind(patternToCheck)
     Debug.Print "Pattern " & patternToCheck & " replaced"
     'validRegex = True
     validRegex = IsValidRegex(patternToCheck)
     foundLookBehind = True
     'regexCell.Interior.ColorIndex = 45 'Orange
    End If
    If validRegex Then
     'Debug.Print "Pattern " & regexCell.Value & " is valid."
    Else
    'If the value in the cell is is not valid regex pattern highlight the cell.
    regexCell.Interior.ColorIndex = 3 'Red
    foundRegexErrors = True
    Debug.Print "Pattern " & regexCell.Value & " is invalid."
    End If
    
    
    Next regexCell
    
    If foundRegexErrors Then
    Validate = Validate & vbCrLf & "* Incorrect regex patterns found. Please check all the red errors and confirm that they are false positives before creating the regex file."
   
    Else
    Debug.Print "No regex pattern errors"
    End If
    
    ' Here we start validating Description column for empty and duplicate values.
    'Declare variables
    Dim colRange As Range
    Dim cell As Range
    
    'Set the column range to be highlighted
    Set colRange = Range("E7:E" & Range("F" & Rows.Count).End(xlUp).Row)
    
    'Reset the colors
    For Each cell In colRange
    cell.Interior.ColorIndex = xlNone
    
    Next cell
    
    'Loop through each cell in the column
    For Each cell In colRange
    
    'If the cell is empty, highlight it in grey
    If IsEmpty(cell) Then
    cell.Interior.ColorIndex = 15
    End If
    
    Next cell
    
    'Count the number of empty cells in the column
    numEmptyCells = Application.WorksheetFunction.CountBlank(colRange)
    
    'Warn the user if any empty cells were found
    If numEmptyCells > 0 Then
    Validate = Validate & vbCrLf & "* Empty cells were found and highlighted in grey under Descriptions column. You cannot leave description field empty. If those rows are not needed, please delete them."
   
    End If
    
    'Declare variables
    Dim dictValues As Object
    
    'Create a dictionary to store the unique values in the column range
    Set dictValues = CreateObject("Scripting.Dictionary")
    
    'Loop through each cell in the column range
    For Each cell In colRange
    
    'If the value in the cell is not already in the dictionary, add it to the dictionary
    If Not dictValues.Exists(cell.Value) Then
    dictValues.Add cell.Value, True
    Else
    'If the value in the cell is already in the dictionary, then the column has duplicate values
    cell.Interior.ColorIndex = 3 'Red
    End If
    
    Next cell
    
    'If any duplicate values were found, display a message box warning the user
    If dictValues.Count < colRange.Cells.Count Then
    Validate = Validate & vbCrLf & "* There are duplicate values under Descriptions column. They are already highlighted in red. Make sure all descriptions are unique."
 
    End If
    
    
End Function


