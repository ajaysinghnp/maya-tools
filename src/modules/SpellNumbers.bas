Attribute VB_Name = "SpellNumbers"
'Spell the numbers in English
Function SpellNumberEng(ByVal MyNumber)
Dim Dollars, Cents, temp
Dim DecimalPlace, Count
ReDim Place(9) As String
Place(2) = " Thousand "
Place(3) = " Million "
Place(4) = " Billion "
Place(5) = " Trillion "

' String representation of amount.
MyNumber = Trim(Str(MyNumber))

' Position of decimal place 0 if none.
DecimalPlace = InStr(MyNumber, ".")

' Convert cents and set MyNumber to dollar amount.
If DecimalPlace > 0 Then
Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
End If
Count = 1
Do While MyNumber <> ""
temp = GetHundreds(Right(MyNumber, 3))
If temp <> "" Then Dollars = temp & Place(Count) & Dollars
If Len(MyNumber) > 3 Then
MyNumber = Left(MyNumber, Len(MyNumber) - 3)
Else
MyNumber = ""
End If
Count = Count + 1
Loop
Select Case Dollars
Case ""
Dollars = ""
Case "One"
Dollars = "One Rupees"
Case Else
Dollars = Dollars & " Rupees"
End Select
Select Case Cents
Case ""
Cents = ""
Case "One"
If Dollars <> "" Then
Cents = " and One Paisa"
Else
Cents = "One Paisa"
End If
Case Else
If Dollars <> "" Then
Cents = " and " & Cents & " Paisa"
Else
Cents = Cents & " Paisa"
End If
End Select
SpellNumberEng = "NPR. " & Dollars & Cents & " Only."
End Function

'Spell the numbers in Nepali Bishal Font
Function SpellNumberNep(ByVal MyNumber)
Dim Dollars, Cents, temp
Dim DecimalPlace, Count
ReDim Place(9) As String
Place(2) = " xhf/ "
Place(3) = " nfv "
Place(4) = " s/f]* "
Place(5) = " ca{ "
Place(6) = " va{ "
Place(7) = " gLn "
Place(8) = " kbd "
Place(9) = " z+v "

' String representation of amount.
MyNumber = Trim(Str(MyNumber))

' Position of decimal place 0 if none.
DecimalPlace = InStr(MyNumber, ".")

' Convert cents and set MyNumber to dollar amount.
If DecimalPlace > 0 Then
Cents = GetHundredsNep(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
End If
Count = 1
Do While MyNumber <> ""
If Count = 1 Then
temp = GetHundredsNep(Right(MyNumber, 3))
If temp <> "" Then Dollars = temp & Place(Count) & Dollars
If Len(MyNumber) > 3 Then
MyNumber = Left(MyNumber, Len(MyNumber) - 3)
Else
MyNumber = ""
End If
Else
temp = GetHundredsNep(Right(MyNumber, 2))
If temp <> "" Then Dollars = temp & Place(Count) & Dollars
If Len(MyNumber) > 2 Then
MyNumber = Left(MyNumber, Len(MyNumber) - 2)
Else
MyNumber = ""
End If
End If
Count = Count + 1
Loop
Select Case Dollars
Case ""
Dollars = ""
Case "One"
Dollars = "Ps ?k}of¤"
Case Else
Dollars = Dollars & " ?k}of¤"
End Select
Select Case Cents
Case ""
Cents = ""
Case "Ps"
If Dollars <> "" Then
Cents = " / Ps k};f"
Else
Cents = "Ps k};f"
End If
Case Else
If Dollars <> "" Then
Cents = " / " & Cents & " k};f"
Else
Cents = Cents & " k};f"
End If
End Select
SpellNumberNep = "cIf/]kL " & Dollars & Cents & " dfq"
End Function

' Converts a number from 100-999 into text
Function GetHundreds(ByVal MyNumber)
Dim Result As String
If Val(MyNumber) = 0 Then Exit Function
MyNumber = Right("000" & MyNumber, 3)

' Convert the hundreds place.
If Mid(MyNumber, 1, 1) <> "0" Then
Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
End If

' Convert the tens and ones place.
If Mid(MyNumber, 2, 1) <> "0" Then
Result = Result & GetTens(Mid(MyNumber, 2))
Else
Result = Result & GetDigit(Mid(MyNumber, 3))
End If

GetHundreds = Result

End Function

' Converts a number from 100-999 into text
Function GetHundredsNep(ByVal MyNumber)
Dim Result As String
If Val(MyNumber) = 0 Then Exit Function
MyNumber = Right("000" & MyNumber, 3)

' Convert the hundreds place.
If Mid(MyNumber, 1, 1) <> "0" Then
Result = GetDigitNep(Mid(MyNumber, 1, 1)) & " ;o "
End If

' Convert the tens and ones place.
If Mid(MyNumber, 2, 1) <> "0" Then
Result = Result & GetTensNep(Mid(MyNumber, 2))
Else
Result = Result & GetDigitNep(Mid(MyNumber, 3))
End If

GetHundredsNep = Result

End Function

' Converts a number from 10 to 99 into English text.
Function GetTens(TensText)
Dim Result As String
Result = "" ' Null out the temporary function value.
If Val(Left(TensText, 1)) = 1 Then ' If value between 10-19...
Select Case Val(TensText)
Case 10: Result = "Ten"
Case 11: Result = "Eleven"
Case 12: Result = "Twelve"
Case 13: Result = "Thirteen"
Case 14: Result = "Fourteen"
Case 15: Result = "Fifteen"
Case 16: Result = "Sixteen"
Case 17: Result = "Seventeen"
Case 18: Result = "Eighteen"
Case 19: Result = "Nineteen"
Case Else
End Select
Else ' If value between 20-99...
Select Case Val(Left(TensText, 1))
Case 2: Result = "Twenty "
Case 3: Result = "Thirty "
Case 4: Result = "Forty "
Case 5: Result = "Fifty "
Case 6: Result = "Sixty "
Case 7: Result = "Seventy "
Case 8: Result = "Eighty "
Case 9: Result = "Ninety "
Case Else
End Select
Result = Result & GetDigit(Right(TensText, 1)) ' Retrieve ones place.
End If

GetTens = Result

End Function

' Converts a number from 10 to 99 into bishall font text.
Function GetTensNep(TensText)
Dim Result As String
Result = "" ' Null out the temporary function value.
Select Case Val(TensText)
Case 10: Result = "bz"
Case 11: Result = "P#f/"
Case 12: Result = "afx|"
Case 13: Result = "t]x|"
Case 14: Result = "rf}w"
Case 15: Result = "kGw|"
Case 16: Result = ";f]x|"
Case 17: Result = ";q"
Case 18: Result = "c&f/"
Case 19: Result = "pGgfO{;"
Case 20: Result = "la;"
Case 21: Result = "PSsfO{;"
Case 22: Result = "afO{;"
Case 23: Result = "t]O{;"
Case 24: Result = "rf}la;"
Case 25: Result = "kRrL;"
Case 26: Result = "%lAa;"
Case 27: Result = ";QfO{;"
Case 28: Result = "c¶fO{;"
Case 29: Result = "pGtL;"
Case 30: Result = "tL;"
Case 31: Result = "PstL;"
Case 32: Result = "aQL;"
Case 33: Result = "t}+tL;"
Case 34: Result = "rf}+tL;"
Case 35: Result = "k}+tL;"
Case 36: Result = "%tL;"
Case 37: Result = ";}+tL;"
Case 38: Result = "c+*\tL;"
Case 39: Result = "pGrfln;"
Case 40: Result = "rfln;"
Case 41: Result = "Psrfln;"
Case 42: Result = "aofln;"
Case 43: Result = "lqrfln;"
Case 44: Result = "rf}jfln;"
Case 45: Result = "k}+tfln;"
Case 46: Result = "l%ofln;"
Case 47: Result = ";}+tfln;"
Case 48: Result = "c+*\tfln;"
Case 49: Result = "pGrf;"
Case 50: Result = "krf;"
Case 51: Result = "PsfpGg"
Case 52: Result = "afpGg"
Case 53: Result = "lqkGg"
Case 54: Result = "rf}aGg"
Case 55: Result = "krkGg"
Case 56: Result = "%kGg"
Case 57: Result = ";GtfpGg"
Case 58: Result = "cG&fpGg"
Case 59: Result = "pG;f¶L"
Case 60: Result = ";f¶L"
Case 61: Result = "Ps;¶L"
Case 62: Result = "a};¶L"
Case 63: Result = "lq;¶L"
Case 64: Result = "rf}+;¶L"
Case 65: Result = "k}+;¶L"
Case 66: Result = "l%of;¶L"
Case 67: Result = ";+/;¶L"
Case 68: Result = "c+/;¶L"
Case 69: Result = "pG;Q/L"
Case 70: Result = ";Q/L"
Case 71: Result = "PsxQ/"
Case 72: Result = "axQ/"
Case 73: Result = "ltxQ/"
Case 74: Result = "rf}xQ/"
Case 75: Result = "krxQ/"
Case 76: Result = "l%xQ/"
Case 77: Result = ";txQ/"
Case 78: Result = "c&xQ/"
Case 79: Result = "pgf;L"
Case 80: Result = "c:;L"
Case 81: Result = "Psf;L"
Case 82: Result = "aof;L"
Case 83: Result = "lt/f;L"
Case 84: Result = "rf}/f;L"
Case 85: Result = "krf;L"
Case 86: Result = "l%of;L"
Case 87: Result = ";Qf;L"
Case 88: Result = "c&f;L"
Case 89: Result = "pgfggAa]"
Case 90: Result = "gAa]"
Case 91: Result = "PsfgAa]"
Case 92: Result = "aofgAa]"
Case 93: Result = "lt/fgAa]"
Case 94: Result = "rf}/fgAa]"
Case 95: Result = "kGrfgAa]"
Case 96: Result = "l%ofgAa]"
Case 97: Result = ";GtfgAa]"
Case 98: Result = "cG&fgAa]"
Case 99: Result = "pgfg;o"
Case Else
End Select
Result = Result

GetTensNep = Result

End Function

' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
Select Case Val(Digit)
Case 1: GetDigit = "One"
Case 2: GetDigit = "Two"
Case 3: GetDigit = "Three"
Case 4: GetDigit = "Four"
Case 5: GetDigit = "Five"
Case 6: GetDigit = "Six"
Case 7: GetDigit = "Seven"
Case 8: GetDigit = "Eight"
Case 9: GetDigit = "Nine"
Case Else: GetDigit = ""
End Select
End Function

' Converts a number from 1 to 9 into text.
Function GetDigitNep(Digit)
Select Case Val(Digit)
Case 1: GetDigitNep = "Ps"
Case 2: GetDigitNep = "b'O{"
Case 3: GetDigitNep = "tLg"
Case 4: GetDigitNep = "rf/"
Case 5: GetDigitNep = "kf¤r"
Case 6: GetDigitNep = "%"
Case 7: GetDigitNep = ";ft"
Case 8: GetDigitNep = "cf&"
Case 9: GetDigitNep = "gf}"
Case Else: GetDigitNep = ""
End Select
End Function

'Converts Number to UNICODE
