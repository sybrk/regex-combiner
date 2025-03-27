Attribute VB_Name = "CombineModule"
Sub combine()
Attribute combine.VB_ProcData.VB_Invoke_Func = " \n14"
'
' combine Makro
    Dim errorMessages As String
    'Validate and get error messages if any
    errorMessages = Validate()
    Dim errorCount As Integer
    Dim userResponse As Integer
    
    ' Verify if there are any error messages.
    errorCount = Len(errorMessages)
    
    'Check if there are error messages and prompt user
    If errorCount > 0 Then
        userResponse = MsgBox(errorMessages + vbCrLf + "Do you still want to create regex file?", vbYesNo + vbQuestion, "User Confirmation")
        If userResponse = vbYes Then
        ' Export regex file
        Call ExportRegexFile
        MsgBox "Regex file forced to be created."
        ElseIf userResponse = vbNo Then
        MsgBox "Please fix the errors", vbCritical, "Errors"
        End If
    Else
    ' Export regex file
        Call ExportRegexFile
    MsgBox "Regex File Created!", vbInformation, "Success"
    End If
End Sub
