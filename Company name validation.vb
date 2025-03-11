' Substitute function doesn't work Properly
Sub ProcessData()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long
    Dim colNum As Integer
    Dim result1() As String, result2() As String
    Dim result3() As String, result4() As String
    Dim uniqueValues() As String
    Dim uniqueCount As Long
    
    ' Set active worksheet
    Set ws = ActiveSheet

    ' Get the active column from the selection
    If Selection.Columns.Count > 1 Then
        MsgBox "Please select only one column.", vbExclamation, "Error"
        Exit Sub
    End If
    colNum = Selection.Column

    ' Find last row in the selected column
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row

    ' Find last used column and determine where to start appending results
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1

    ' Initialize arrays
    ReDim result1(1 To lastRow)
    ReDim result2(1 To lastRow)
    ReDim result3(1 To lastRow)
    ReDim result4(1 To lastRow)
    ReDim uniqueValues(1 To lastRow)
    
    ' Step 1: Remove special characters
    For i = 1 To lastRow
        result1(i) = RemoveSpecialCharacters(ws.Cells(i, colNum).Value)
    Next i

    ' Step 2: Substitute diacritic letters
    For i = 1 To lastRow
        result2(i) = ReplaceDiacritics(result1(i))
    Next i

    ' Step 3: Trim leading and trailing spaces
    For i = 1 To lastRow
        result3(i) = Trim(result2(i))
    Next i
    ws.Cells(1, lastCol).Value = "Trimmed Values"
    For i = 1 To lastRow
        ws.Cells(i, lastCol).Value = result3(i)
    Next i
    lastCol = lastCol + 1

    ' Step 4: Convert to lowercase and remove spaces
    For i = 1 To lastRow
        result4(i) = Replace(LCase(result3(i)), " ", "")
    Next i
    ws.Cells(1, lastCol).Value = "Lowercase No Spaces"
    For i = 1 To lastRow
        ws.Cells(i, lastCol).Value = result4(i)
    Next i
    lastCol = lastCol + 1

    ' Step 5: Find unique values
    uniqueCount = 0
    For i = 1 To lastRow
        If Not IsInArray(result4(i), uniqueValues, uniqueCount) Then
            uniqueCount = uniqueCount + 1
            uniqueValues(uniqueCount) = result4(i)
        End If
    Next i

    ' Output unique values
    ws.Cells(1, lastCol).Value = "Unique Values"
    For i = 1 To uniqueCount
        ws.Cells(i, lastCol).Value = uniqueValues(i)
    Next i
    lastCol = lastCol + 1

    ' Step 6: Count occurrences
    ws.Cells(1, lastCol).Value = "Unique Count"
    ws.Cells(2, lastCol).Value = uniqueCount

    ' Success message
    MsgBox "Macro executed successfully!", vbInformation, "Success"
End Sub

' Function to remove special characters (without ActiveX)
Function RemoveSpecialCharacters(text As String) As String
    Dim i As Integer
    Dim cleanText As String
    Dim ch As String
    
    cleanText = ""
    
    ' Keep only letters (a-z, A-Z), numbers (0-9), and spaces
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = " " Then
            cleanText = cleanText & ch
        End If
    Next i
    
    RemoveSpecialCharacters = cleanText
End Function

' Function to replace diacritic letters
Function ReplaceDiacritics(text As String) As String
    Dim diacritics As Variant, replacements As Variant
    Dim i As Integer

    diacritics = Array("á", "à", "â", "ä", "ã", "å", "é", "è", "ê", "ë", "í", "ì", "î", "ï", _
                       "ó", "ò", "ô", "ö", "õ", "ú", "ù", "û", "ü", "ý", "ÿ", "ñ", "ç")
    replacements = Array("a", "a", "a", "a", "a", "a", "e", "e", "e", "e", "i", "i", "i", "i", _
                         "o", "o", "o", "o", "o", "u", "u", "u", "u", "y", "y", "n", "c")

    For i = LBound(diacritics) To UBound(diacritics)
        text = Replace(text, diacritics(i), replacements(i))
    Next i

    ReplaceDiacritics = text
End Function

' Function to check if a value exists in an array
Function IsInArray(value As String, arr As Variant, count As Long) As Boolean
    Dim i As Long
    For i = 1 To count
        If arr(i) = value Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function





' No error messages, overrides the data

Sub ProcessData()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim col As Long, lastRow As Long
    Dim result1 As Variant, result2 As Variant, result3 As Variant
    Dim result4 As Variant, uniqueValues As Variant
    Dim result6 As Integer
    Dim i As Long, uniqueCount As Integer
    
    ' Set worksheet
    Set ws = ActiveSheet

    ' Get selected column or default to first column
    If Selection.Columns.Count = 1 Then
        col = Selection.Column
    Else
        col = 1
    End If

    ' Find last row in selected column
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col))

    ' Step 1: Remove Special Characters
    ReDim result1(1 To lastRow)
    For i = 1 To lastRow
        result1(i) = RemoveSpecialCharacters(rng.Cells(i, 1).Value)
    Next i

    ' Step 2: Substitute Diacritic Letters
    ReDim result2(1 To lastRow)
    For i = 1 To lastRow
        result2(i) = ReplaceDiacritics(result1(i))
    Next i

    ' Step 3: Trim Leading & Trailing Spaces
    ReDim result3(1 To lastRow)
    For i = 1 To lastRow
        result3(i) = Trim(result2(i))
    Next i

    ' Step 4: Convert to Lowercase & Remove Spaces
    ReDim result4(1 To lastRow)
    For i = 1 To lastRow
        result4(i) = LCase(Replace(result3(i), " ", ""))
    Next i

    ' Step 5: Find Unique Values (Fixed)
    uniqueCount = 0
    ReDim uniqueValues(1 To lastRow)

    For i = 1 To lastRow
        If Not IsInArray(CStr(result4(i)), uniqueValues, uniqueCount) Then
            uniqueCount = uniqueCount + 1
            uniqueValues(uniqueCount) = result4(i)
        End If
    Next i

    ' Step 6: Count Unique Values
    result6 = uniqueCount

    ' Output results
    ws.Cells(1, col + 1).Resize(lastRow).Value = Application.Transpose(result1) ' Step 1
    ws.Cells(1, col + 2).Resize(lastRow).Value = Application.Transpose(result2) ' Step 2
    ws.Cells(1, col + 3).Resize(lastRow).Value = Application.Transpose(result3) ' Step 3
    ws.Cells(1, col + 4).Resize(lastRow).Value = Application.Transpose(result4) ' Step 4
    ws.Cells(1, col + 5).Resize(uniqueCount).Value = Application.Transpose(uniqueValues) ' Step 5
    ws.Cells(1, col + 6).Value = result6 ' Step 6
End Sub

' Function to Remove Special Characters
Function RemoveSpecialCharacters(text As String) As String
    Dim chars As String, i As Integer
    chars = "!@#$%^&~*()_+={}[]|\:;'" & Chr(34) & "<>,.?/~`"
    For i = 1 To Len(chars)
        text = Replace(text, Mid(chars, i, 1), "")
    Next i
    RemoveSpecialCharacters = text
End Function

' Function to Replace Diacritic Letters
Function ReplaceDiacritics(ByVal text As Variant) As String
    Dim replacements As Variant, i As Integer
    replacements = Array("ÀÁÂÃÄÅĀĂĄ", "A", "Æ", "AE", "ÇĆĈĊČ", "C", "ÐĐ", "D", _
                         "ÈÉÊËĒĔĖĘĚ", "E", "ĜĞĠĢ", "G", "ĤĦ", "H", "ÌÍÎÏĪĬĮİı", "I", _
                         "Ĵ", "J", "Ķ", "K", "ĹĻĽĿŁ", "L", "ÑŃŅŇ", "N", "ÒÓÔÕÖØŌŎŐ", "O", _
                         "Œ", "OE", "ŔŖŘ", "R", "ŚŜŞŠ", "S", "ŢŤŦ", "T", "ÙÚÛÜŪŬŮŰŲ", "U", _
                         "Ŵ", "W", "ÝŸŶ", "Y", "ŹŻŽ", "Z")

    For i = LBound(replacements) To UBound(replacements) Step 2
        Dim j As Integer
        For j = 1 To Len(replacements(i))
            text = Replace(text, Mid(replacements(i), j, 1), replacements(i + 1))
            text = Replace(text, LCase(Mid(replacements(i), j, 1)), LCase(replacements(i + 1)))
        Next j
    Next i
    ReplaceDiacritics = text
End Function

' Function to Check If Value Exists in Array (Fixed)
Function IsInArray(value As String, arr As Variant, count As Integer) As Boolean
    Dim i As Integer
    For i = 1 To count
        If arr(i) = value Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
