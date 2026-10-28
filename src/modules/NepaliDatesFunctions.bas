Attribute VB_Name = "NepaliDatesFunctions"
'==========================================================
'created by Ajay Singh
'https://ajaysingh.com.np
'contact email : ajay@mentormaya.com
'==========================================================

Public Starting_Range_AD As Date 'Public variable declration for database starting date in AD i.e. 2000-01-01 is 1943-04-14 AD
Public BS(91) 'Declaring BS database array variable
Public Array_Size As Double 'array size variable for array size calcuation and storing

Sub BS_Database()
    Starting_Range_AD = "1943-04-14" 'database starting date in AD i.e. 2000-01-01 is 1943-04-14 AD
    
    '========================================================================
    'all the BS database array starts with year and total days for each subsequent 12 months
    BS(0) = Array(2000, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(1) = Array(2001, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(2) = Array(2002, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(3) = Array(2003, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(4) = Array(2004, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(5) = Array(2005, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(6) = Array(2006, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(7) = Array(2007, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(8) = Array(2008, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31)
    BS(9) = Array(2009, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(10) = Array(2010, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(11) = Array(2011, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(12) = Array(2012, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    BS(13) = Array(2013, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(14) = Array(2014, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(15) = Array(2015, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(16) = Array(2016, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    BS(17) = Array(2017, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(18) = Array(2018, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(19) = Array(2019, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(20) = Array(2020, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(21) = Array(2021, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(22) = Array(2022, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    BS(23) = Array(2023, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(24) = Array(2024, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(25) = Array(2025, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(26) = Array(2026, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(27) = Array(2027, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(28) = Array(2028, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(29) = Array(2029, 31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30)
    BS(30) = Array(2030, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(31) = Array(2031, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(32) = Array(2032, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(33) = Array(2033, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(34) = Array(2034, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(35) = Array(2035, 30, 32, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31)
    BS(36) = Array(2036, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(37) = Array(2037, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(38) = Array(2038, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(39) = Array(2039, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    BS(40) = Array(2040, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(41) = Array(2041, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(42) = Array(2042, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(43) = Array(2043, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    BS(44) = Array(2044, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(45) = Array(2045, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(46) = Array(2046, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(47) = Array(2047, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(48) = Array(2048, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(49) = Array(2049, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    BS(50) = Array(2050, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(51) = Array(2051, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(52) = Array(2052, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(53) = Array(2053, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    BS(54) = Array(2054, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(55) = Array(2055, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(56) = Array(2056, 31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30)
    BS(57) = Array(2057, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(58) = Array(2058, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(59) = Array(2059, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(60) = Array(2060, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(61) = Array(2061, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(62) = Array(2062, 30, 32, 31, 32, 31, 31, 29, 30, 29, 30, 29, 31)
    BS(63) = Array(2063, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(64) = Array(2064, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(65) = Array(2065, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(66) = Array(2066, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31)
    BS(67) = Array(2067, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(68) = Array(2068, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(69) = Array(2069, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(70) = Array(2070, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    BS(71) = Array(2071, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(72) = Array(2072, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    BS(73) = Array(2073, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    BS(74) = Array(2074, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(75) = Array(2075, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(76) = Array(2076, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    BS(77) = Array(2077, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    BS(78) = Array(2078, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(79) = Array(2079, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    BS(80) = Array(2080, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    BS(81) = Array(2081, 31, 31, 32, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    BS(82) = Array(2082, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    BS(83) = Array(2083, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30)
    BS(84) = Array(2084, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30)
    BS(85) = Array(2085, 31, 32, 31, 32, 30, 31, 30, 30, 29, 30, 30, 30)
    BS(86) = Array(2086, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    BS(87) = Array(2087, 31, 31, 32, 31, 31, 31, 30, 30, 29, 30, 30, 30)
    BS(88) = Array(2088, 30, 31, 32, 32, 30, 31, 30, 30, 29, 30, 30, 30)
    BS(89) = Array(2089, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    BS(90) = Array(2090, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    BS(91) = Array(2091, 31, 31, 32, 31, 31, 31, 30, 30, 29, 30, 30, 30)
    
    Array_Size = UBound(BS) 'storing the value of array size
End Sub

Function AD2BS(AD As Date) As String 'Function for converting AD to BS
    Call BS_Database 'Calling the database in memory

    DD_Gap = DateDiff("d", Starting_Range_AD, AD) 'calculating the days difference between entered AD date and starting database AD date
    
    If DD_Gap < 0 Then GoTo Out_Of_Range 'condition for checking whether the entered AD date is befor the starting AD date in the database
    
    DD_total = 0 'creating variable for cummulative days in the database
    
    For i = 0 To Array_Size 'Loop for each year
        For ii = 1 To 12 'Loop for each month
            DD_total = DD_total + BS(i)(ii) 'calculating cummulative days by adding days in each months
            
            If DD_total > DD_Gap Then 'checking whether the cummulative days is greater than DD_Gap variable
                DD_total = DD_total - BS(i)(ii) 'reversing the last addition of days
                GoTo Return_Value
            End If
            
        Next
    Next

Out_Of_Range:

    AD2BS = "Error! Out of range..." 'returning error

Exit Function

Return_Value:

    AD2BS = BS(i)(0) & "-" & ii & "-" & (DD_Gap - DD_total + 1) 'returning BS value

End Function

Function BS2AD(BS_Date As String) As String 'Function for converting BS to AD
    Call BS_Database 'Calling the database in memory
    Out_Of_Range = True 'Boolean variable for checking whether entered BS falls within the database range
    
    For BS_Index = 0 To Array_Size 'Finding the BS array index in which the BS year lies
        If BS_Year(BS_Date) = BS(BS_Index)(0) Then 'checking whether BS index for the BS year is found or not
            Out_Of_Range = False
            Exit For
        End If
    Next
    
    If Out_Of_Range = True Then GoTo Out_Of_Range

    DD_total = 0 'creating variable for cummulative days in the database
    For i = 0 To BS_Index 'loop for each year until the desired BS index year is reached
        For ii = 1 To 12 'loop for each month
            If i = BS_Index And ii = BS_Month(BS_Date) Then 'checking whether we reached the entered BS date year and month
                BS2AD = DateAdd("d", DD_total + BS_Day(BS_Date) - 1, Starting_Range_AD) 'adding cummulative days+days entered in BS date with starting database AD date
                
                If BS_Year(BS_Date) <> BS_Year(AD2BS(DateValue(BS2AD))) Or _
                    BS_Month(BS_Date) <> BS_Month(AD2BS(DateValue(BS2AD))) Or _
                    BS_Day(BS_Date) <> BS_Day(AD2BS(DateValue(BS2AD))) Then 'checking whether reverse conversion is same as the orginal BS date or not
                   
                   GoTo Out_Of_Range
                End If
                BS2AD = Year(BS2AD) & "-" & Month(BS2AD) & "-" & Day(BS2AD)
                Exit Function
            End If
            DD_total = DD_total + BS(i)(ii) 'cummulative days calculation
        Next
    Next

Out_Of_Range:
BS2AD = "Error! Out of range..."

End Function
Function BS_Date_Diff(BS_Date1 As String, BS_Date2 As String) As Double 'Function for differecning two BS dates

BS_Date_Diff = DateDiff("d", BS2AD(BS_Date1), BS2AD(BS_Date2))

End Function

Function BS_Date_Day_Add(BS_Date As String, Days_To_Add As Double) As String 'Function for adding days to BS date

BS_Date_Day_Add = AD2BS(DateAdd("d", Days_To_Add, BS2AD(BS_Date)))

End Function
Function BS_Day_Of_Week(BS_Date As String) As Integer 'Function for returning day of week in BS Date

BS_Day_Of_Week = Weekday(BS2AD(BS_Date))

End Function
Function BS_Month(BS As String) As Double 'Function for returning BS Month in BS Date

For i = 1 To Len(BS) 'creating loop for each character in BS Date
    If IsNumeric(Mid(BS, i, 1)) = False Then 'checking whether year and month separator is reached or not
    
        If IsNumeric(Mid(BS, i + 2, 1)) = False Then 'checking whether there is one or two digit month
            BS_Month = Mid(BS, i + 1, 1) 'returning month value if there is one digit month
            Exit Function
        Else
            BS_Month = Mid(BS, i + 1, 2) 'returning month value if there is two digit month
            Exit Function
        End If
    
    End If
Next

End Function

Function BS_Year(BS As String) As Double 'Funcation for returning year of the BS date

BS_Year = Mid(BS, 1, 4)

End Function

Function BS_Day(BS As String) As Double 'Function for returning day of the BS date

If IsNumeric(Mid(BS, Len(BS) - 1, 1)) = True Then 'checking whether one or two digit day
    BS_Day = Right(BS, 2) 'returning two digit day
Else
    BS_Day = Right(BS, 1) 'returning one digit day
End If

End Function

