Attribute VB_Name = "HelperFunctions"
   Function ContainsLookBehind(pattern As String) As Boolean
        'checks if a pattern contains lookbehind regex
        ContainsLookBehind = False
        If InStr(1, pattern, "(?<", vbBinaryCompare) > 0 Then
        ContainsLookBehind = True
        End If
    End Function

    Function IsValidRegex(pattern As String) As Boolean
        'validated a pattern against VBA regex
        On Error GoTo InvalidPattern
        Dim regex As Object
        Dim forbiddenChars() As Variant
        'VBA regex does not understand if some special characters cause issue. These characters are listed below and checked separately.
        forbiddenChars = Array("\c", "\g", "\h", "\i", "\j", "\k", "\l", "\m", "\o", "\p", "\u", "\q", "\y", "\x")
        Set regex = CreateObject("VBScript.RegExp")
        regex.pattern = pattern
        'hello world here is not important. It is just to test if a regex pattern works or not.
        regex.Execute ("hello world!")
        IsValidRegex = True
        'here we are checking for forbidden characters.
        For i = LBound(forbiddenChars) To UBound(forbiddenChars)
        If InStr(1, pattern, forbiddenChars(i), vbBinaryCompare) > 0 Then
            IsValidRegex = False
            'Debug.Print forbiddenChars(i) & " issue in " & pattern
            ' here we create exception for \p when it is followed by {
            If forbiddenChars(i) = "\p" Then
                If Mid(pattern, InStr(1, pattern, forbiddenChars(i), vbBinaryCompare) + Len(forbiddenChars(i)), 1) = "{" Then
                    IsValidRegex = True
                    'Debug.Print "issue fixed in " & pattern
                    Exit Function
                End If
            End If
            ' here we create exception for \u when it is followed by a number
            If forbiddenChars(i) = "\u" Then
                If IsNumeric(Mid(pattern, InStr(1, pattern, forbiddenChars(i), vbBinaryCompare) + Len(forbiddenChars(i)), 1)) Then
                    IsValidRegex = True
                    'Debug.Print "issue fixed in " & pattern
                    Exit Function
                End If
            End If
            Exit Function
        End If
        
        Next i
        Exit Function
InvalidPattern:
        IsValidRegex = False
    End Function
    
    Function ReplaceLookBehind(pattern As String) As String
        'removes looksbehind chars from pattern
        ReplaceLookBehind = Replace(pattern, "?<!", "")
        ReplaceLookBehind = Replace(ReplaceLookBehind, "?<=", "")
    End Function
    Function CaseSensitiveExists(dict As Object, key As String) As Boolean
        Dim k As Variant
        CaseSensitiveExists = False
        For Each k In dict.Keys
            If k = key Then
            CaseSensitiveExists = True
            Exit Function
            End If
        Next k
    End Function


    Sub ExportRegexFile()
        'Set the ID of first regex as regexrules0
        Range("D7").Select
        ActiveCell.FormulaR1C1 = "RegExRules0"
        Range("D7").Select
        'Autofill if user entered more then one regex.
        If Range("D7:D" & Range("E" & Rows.Count).End(xlUp).Row).Rows.Count > 1 Then
        Selection.autofill Destination:=Range("D7:D" & Range("E" & Rows.Count).End(xlUp).Row)
        End If
        'Export regex file
        ActiveWorkbook.XmlMaps("SettingsBundle_Mapping").Export _
        ActiveWorkbook.Path & "\combined.sdlqasettings", True
    End Sub



