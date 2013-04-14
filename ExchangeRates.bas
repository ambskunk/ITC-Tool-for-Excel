Attribute VB_Name = "ExchangeRates"
Option Explicit

Function GetExchangeRatesFromHTML(sHTMLVal As String, Optional cERsExisting As cExchangeRates) As cExchangeRates
    '---------------------------------------------------------------------------------------------------------
    ' Description: Function to read a source of HTML and scrap the exchange rates from it
    ' Inputs:
    ' -sHTMLVal: The HTML source code
    ' -cERsExisting: A collection of Exchange Rate Objects that the new ER Objects should be added to (if
    ' empty then create a new one)
    '---------------------------------------------------------------------------------------------------------
    Dim sSearchString As String, sRegionString As String
    Dim sMonthSearch As String, sMonth As String, sYear As String, sExRate As String
    Dim lStart As Long, lEnd As Long, lStartMth As Long, lEndMth As Long, lStartYr As Long, lEndYr As Long
    Dim lStartEx As Long, lEndEx As Long, cER As cExchangeRate, cers As New cExchangeRates
    Dim vCurrencyCodes As Variant
    Dim sFirstMonthYear As String, sThisMonthYear As String
    Dim lStartPayment As Long, sPayment As String, sPaymentConverted As String
    
    
    If Not cERsExisting Is Nothing Then Set cers = cERsExisting
    
    '========================================================================
    ' Check the first month and year in the source
    '========================================================================
    sMonthSearch = "<span style=""padding-left: 7px;"">Earned</span>"
    lStartMth = VBA.InStr(1, sHTMLVal, sMonthSearch)
    lEndMth = VBA.InStr(lStartMth, sHTMLVal, "<span class=""month"">")
    sMonth = RemoveTabCharacters(VBA.Mid$(sHTMLVal, lStartMth + VBA.Len(sMonthSearch), lEndMth - lStartMth - VBA.Len(sMonthSearch)))
    lStartYr = lEndMth
    lEndYr = VBA.InStr(lEndMth, sHTMLVal, "</span>")
    sYear = VBA.Trim$(VBA.Mid$(sHTMLVal, lStartYr + VBA.Len("<span class=""month"">"), lEndYr - lStartYr - VBA.Len("<span class=""month"">")))

    'Read in the ranges of the currency and countries code
    vCurrencyCodes = CurrencyCodes          'This is a function to read in the currency codes at the bottom of the module

    '========================================================================
    ' Loop through the source code looking for the countries
    '========================================================================
    sFirstMonthYear = sMonth & sYear
    sSearchString = "<td class=""first"">"
    lStart = VBA.InStr(1, sHTMLVal, sSearchString)
    Do While lStart > 0
        lEnd = VBA.InStr(lStart, sHTMLVal, "</td>")
    
        'Find the country string and if it's not 'Currency' (ie the heading) then read in the data
        sRegionString = VBA.Mid$(sHTMLVal, lStart + VBA.Len(sSearchString), lEnd - lStart - VBA.Len(sSearchString))
        If sRegionString <> "Currency" Then
            lStartEx = VBA.InStr(lStart, sHTMLVal, "<td class=""fx-rate"">")
            lEndEx = VBA.InStr(lStartEx, sHTMLVal, "</td>")
            sExRate = RemoveTabCharacters(VBA.Mid$(sHTMLVal, lStartEx + VBA.Len("<td class=""fx-rate"">"), lEndEx - lStartEx - VBA.Len("<td class=""fx-rate"">")))
            
            'Get the unexchanged payment amount
            lStartPayment = lStartEx - 47
            sPayment = VBA.Trim$(VBA.Mid$(sHTMLVal, lStartPayment, 11))
            
            'Get the exchanged payment amount
            lStartPayment = VBA.InStr(lStart, sHTMLVal, "<td class=""payment-amount"">")
            lEndEx = VBA.InStr(lStartPayment, sHTMLVal, "</td>")
            sPaymentConverted = RemoveTabCharacters(VBA.Mid$(sHTMLVal, lStartPayment + VBA.Len("<td class=""payment-amount"">"), lEndEx - lStartPayment - VBA.Len("<td class=""payment-amount"">")))
            
            'Get the exchange rate by converting the numbers, but only if it's less than 0.1
            If VBA.Val(sExRate) < 0.1 Then
                If VBA.CDbl(sPayment) > 0 Then sExRate = VBA.CStr(VBA.CDbl(sPaymentConverted) / VBA.CDbl(sPayment))
            End If
            
            'Add the output to the Exchange Rate Object
            Set cER = Nothing
            Set cER = New cExchangeRate
            cER.CurrencyCode = sRegionString
            cER.RegionCode = Application.WorksheetFunction.VLookup(sRegionString, vCurrencyCodes, 2, False)
            cER.Region = Application.WorksheetFunction.VLookup(sRegionString, vCurrencyCodes, 3, False)
            cER.ExchangeRate = VBA.CDbl(sExRate)
            cER.Month = sMonth
            cER.Year = sYear
            
            Application.StatusBar = "Adding Exchange Rate: " & cER.Year & ", " & cER.Month & ", " & cER.Region & ", " & cER.CurrencyCode & ", " & cER.ExchangeRate
            
            cers.Add cER
        End If
        
        lStartMth = VBA.InStr(lStart, sHTMLVal, sMonthSearch)           'Check for the month again
        lStart = VBA.InStr(lEnd, sHTMLVal, sSearchString)               'Check to see if we have another country
        
        'If the next month string is less than the next country code then set it
        If lStartMth < lStart And lStartMth > 0 Then
            lEndMth = VBA.InStr(lStartMth, sHTMLVal, "<span class=""month"">")
    
            sMonth = RemoveTabCharacters(VBA.Mid$(sHTMLVal, lStartMth + VBA.Len(sMonthSearch), lEndMth - lStartMth - VBA.Len(sMonthSearch)))
            lStartYr = lEndMth
            lEndYr = VBA.InStr(lEndMth, sHTMLVal, "</span>")
            sYear = VBA.Trim$(VBA.Mid$(sHTMLVal, lStartYr + VBA.Len("<span class=""month"">"), lEndYr - lStartYr - VBA.Len("<span class=""month"">")))
        End If
        
        'If we only want the latest currency then abort here
        If bDownloadLatestReport Then
            sThisMonthYear = sMonth & sYear
            If sFirstMonthYear <> sThisMonthYear Then lStart = 0
        End If
    Loop
    
    Application.StatusBar = False
    
    'Assign the output
    If Not cers Is Nothing Then Set GetExchangeRatesFromHTML = cers
End Function

Function RemoveTabCharacters(sString As String) As String
    'Function to remove unwanted characters from a filename
    'Anything might make saving a file unhappy
    Dim sOutput As String
    
    sOutput = sString
    Do While VBA.InStr(1, sOutput, vbTab) > 0
        sOutput = VBA.Trim$(VBA.Replace(sOutput, vbTab, "", 1))
    Loop
    Do While VBA.InStr(1, sOutput, vbCr) > 0
        sOutput = VBA.Trim$(VBA.Replace(sOutput, vbCr, "", 1))
    Loop
    Do While VBA.InStr(1, sOutput, vbLf) > 0
        sOutput = VBA.Trim$(VBA.Replace(sOutput, vbLf, "", 1))
    Loop
    RemoveTabCharacters = VBA.Trim$(sOutput)
End Function


Function GetExchangeRatesFromWorksheet(wkFromSheet As Worksheet) As cExchangeRates
    '---------------------------------------------------------------------------------------------------------
    ' Description: Function to extract all the exchange rates from the exchange rate worksheet
    ' Inputs:
    ' -wkFromSheet: The exchange rate worksheet
    ' Notes: This assumes the worksheet is in the correct format.  Will return nothing if it doesn't
    ' find what it expects.
    '---------------------------------------------------------------------------------------------------------
    Dim vData As Variant
    Dim rr As Integer, cc As Integer
    Dim sDate As String, sMonth As String, sYear As String
    Dim cER As cExchangeRate, cers As New cExchangeRates
    
    vData = wkFromSheet.UsedRange
    
    'Loop through all the months
    For cc = LBound(vData, 2) To UBound(vData, 2)
        sDate = vData(1, cc)
        'If we have a string then assume it's a date
        If VBA.Len(sDate) > 0 Then
            sMonth = ReturnFullMonth(GetMonthOrYear(sDate, True))
            sYear = GetMonthOrYear(sDate, False)
            
            'Loop through the rows to check all the country codes
            For rr = LBound(vData, 1) + 1 To UBound(vData, 1)
                If VBA.Len(vData(rr, cc)) > 0 Then
                    'Set up a new exchange rate object
                    Set cER = Nothing
                    Set cER = New cExchangeRate
                    
                    cER.Month = sMonth
                    cER.Year = sYear
                    cER.RegionCode = vData(rr, 1)
                    cER.ExchangeRate = vData(rr, cc)
                    
                    cers.Add cER
                End If
            Next rr
        End If
    Next cc
    
    If Not cers Is Nothing Then Set GetExchangeRatesFromWorksheet = cers
End Function

Sub PutExchangeRatesInWorksheet(wsToSheet As Worksheet, cers As cExchangeRates)
    '---------------------------------------------------------------------------------------------------------
    ' Description: Function to write the exchange rates to a worksheet
    ' Inputs:
    ' -wsToSheet: The worksheet to write to (ie the exchange rate worksheet)
    ' -cERs: The collection of exchange rates
    '---------------------------------------------------------------------------------------------------------
    Dim vData() As Variant
    Dim cER As cExchangeRate
    Dim sMonths() As String, sCountries() As String, rSortRange As Range, rKeyRange As Range, rFormatRange As Range
    Dim ii As Integer, jj As Integer, iCtrMonths As Integer, iCtrCountries As Integer
    Dim bFoundMonths As Boolean, bFoundCountry As Boolean, iCol As Integer, iRow As Integer
    
    ReDim sMonths(1 To 50)
    ReDim sCountries(1 To 50)
    
    'Guess we should check first
    If cers Is Nothing Then Exit Sub
    
    'Start up the counters
    iCtrMonths = 1
    iCtrCountries = 1
    
    'We need to get a count of the number of months and also the number of countries so the
    'array can be dimensioned
    For Each cER In cers
        'Check for the months first
        If iCtrMonths = 1 Then
            sMonths(1) = cER.Month & cER.Year
            iCtrMonths = 2
        Else
            bFoundMonths = False
            For ii = LBound(sMonths) To UBound(sMonths)
                If sMonths(ii) = cER.Month & cER.Year Then
                    bFoundMonths = True
                    Exit For
                ElseIf VBA.Len(sMonths(ii)) = 0 Then
                    Exit For
                End If
            Next ii
            
            'If we haven't found it, then add it
            If Not bFoundMonths Then
                'Redimension is inefficient so only redim in batches of 50
                If iCtrMonths Mod UBound(sMonths) = 0 Then ReDim Preserve sMonths(1 To UBound(sMonths) + 50)
                'Then add the month to the array and increase the counter
                sMonths(iCtrMonths) = cER.Month & cER.Year
                iCtrMonths = iCtrMonths + 1
            End If
        End If
        
        'Check for the countries
        If iCtrCountries = 0 Then
            sCountries(1) = cER.RegionCode
            iCtrCountries = 2
        Else
            bFoundCountry = False
            For ii = LBound(sCountries) To UBound(sCountries)
                If sCountries(ii) = cER.RegionCode Then
                    bFoundCountry = True
                    Exit For
                ElseIf VBA.Len(sCountries(ii)) = 0 Then
                    Exit For
                End If
            Next ii
            
            'If we haven't found it, then add it
            If Not bFoundCountry Then
                If iCtrCountries Mod UBound(sCountries) = 0 Then ReDim Preserve sCountries(1 To UBound(sCountries) + 50)
                sCountries(iCtrCountries) = cER.RegionCode
                iCtrCountries = iCtrCountries + 1
            End If
        End If
    Next cER
    
    'Resize the array back down
    ReDim Preserve sMonths(1 To iCtrMonths - 1)
    ReDim Preserve sCountries(1 To iCtrCountries - 1)
    
    'Sort the countries nicely please
    sCountries = SortArray(sCountries)
    
    ReDim vData(1 To UBound(sMonths) + 1, 1 To UBound(sCountries) + 1)
    For ii = LBound(vData, 2) To UBound(vData, 2)
        For jj = LBound(vData, 1) To UBound(vData, 1)
            If ii = 1 Then
                If jj > LBound(vData, 1) Then
                    vData(jj, ii) = "'" & VBA.Left$(sMonths(jj - 1), 3) & "-" & VBA.Right$(sMonths(jj - 1), 2)
                End If
            Else
                If jj = LBound(vData, 2) Then
                    vData(jj, ii) = sCountries(ii - 1)
                End If
            End If
        Next jj
    Next ii
    
    'Next loop through the exchange rate objects and place in the data array
    For Each cER In cers
        With cER
            iRow = Application.WorksheetFunction.Match(.Month & .Year, sMonths, 0)
            iCol = Application.WorksheetFunction.Match(.RegionCode, sCountries, 0)
            vData(iRow + 1, iCol + 1) = .ExchangeRate
        End With
    Next cER
    
    'Output the data to the worksheet
    With wsToSheet
        .Cells.Clear
        .Cells.HorizontalAlignment = xlCenter
        With .Cells.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        .Range(.Cells(1, 1), .Cells(1 + UBound(vData, 2) - LBound(vData, 2), 1 + UBound(vData, 1) - LBound(vData, 1))).Value2 = Application.Transpose(vData)
        
        'It's likely that the output is not sorted so we'll use the worksheet functions to do so
        Set rKeyRange = .Range(.Cells(1, 2), .Cells(1, 1 + UBound(vData, 1) - LBound(vData, 1)))
        Set rSortRange = .Range(.Cells(1, 2), .Cells(1 + UBound(vData, 2) - LBound(vData, 2), 1 + UBound(vData, 1) - LBound(vData, 1)))
        Set rFormatRange = .Range(.Cells(2, 1), .Cells(1 + UBound(vData, 2) - LBound(vData, 2), 1 + UBound(vData, 1) - LBound(vData, 1)))
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=rKeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .SetRange rSortRange
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlLeftToRight
            .SortMethod = xlPinYin
            .Apply
        End With
        
        With rFormatRange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rFormatRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        rSortRange.NumberFormat = "0.00000"
    End With
End Sub

Function GetMonthOrYear(sFromDate As String, bGetMonth As Boolean) As String
    'Custom Function to Get the Month from a Month-Year string
    Dim lDashLoc As Long, sOutput As String
    
    lDashLoc = VBA.InStr(1, sFromDate, "-")
    If bGetMonth Then
        sOutput = VBA.Left$(sFromDate, lDashLoc - 1)
    Else
        sOutput = VBA.Right$(sFromDate, VBA.Len(sFromDate) - lDashLoc)
        'Yeah if someone is still using this in 2100, then good luck to them
        If VBA.IsNumeric(sOutput) Then sOutput = "20" & sOutput
    End If
    GetMonthOrYear = sOutput
End Function

Function SortArray(ArrayToSort() As String, Optional iColumn As Integer = 1) As String()
    '==================================================================================
    ' This function will sort a one dimensional or two dimension (with specified column)
    ' array.  Should work with numbers and letters
    '==================================================================================
    Dim First As Integer, Last As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim Temp  As String
    Dim iUbound As Integer
     
    On Error Resume Next
    iUbound = -1
    iUbound = UBound(ArrayToSort, 2)
    On Error GoTo 0
    If iUbound > 0 Then

        First = LBound(ArrayToSort, 2)
        Last = UBound(ArrayToSort, 2)
        For i = First To Last - 1
            For j = i + 1 To Last
                If ArrayToSort(iColumn, i) > ArrayToSort(iColumn, j) Then
                    For k = LBound(ArrayToSort, 1) To UBound(ArrayToSort, 1)
                        Temp = ArrayToSort(k, j)
                        ArrayToSort(k, j) = ArrayToSort(k, i)
                        ArrayToSort(k, i) = Temp
                    Next k
                End If
            Next j
        Next i
        
    Else
        On Error GoTo 0
        First = LBound(ArrayToSort)
        Last = UBound(ArrayToSort)
        For i = First To Last - 1
            For j = i + 1 To Last
                If ArrayToSort(i) > ArrayToSort(j) Then
                    Temp = ArrayToSort(j)
                    ArrayToSort(j) = ArrayToSort(i)
                    ArrayToSort(i) = Temp
                End If
            Next j
        Next i
    End If
    
    SortArray = ArrayToSort
End Function

Function CurrencyCodes() As Variant
    Dim curCode(1 To 158) As Variant
    
    curCode(1) = Array("AED", "AE", "United Arab Emirates", "UAE dirham")
    curCode(2) = Array("AFN", "AF", "Afghanistan", "Afghan afghani")
    curCode(3) = Array("ALL", "AL", "Albania", "Albanian lek")
    curCode(4) = Array("AMD", "AM", "Armenia", "Armenian dram")
    curCode(5) = Array("AOA", "AO", "Angola", "Angolan kwanza")
    curCode(6) = Array("ARS", "AR", "Argentina", "Argentine peso")
    curCode(7) = Array("AUD", "AU", "Australia", "Australian dollar")
    curCode(8) = Array("AWG", "AW", "Aruba", "Aruban florin")
    curCode(9) = Array("AZN", "AZ", "Azerbaijan", "Azerbaijani manat")
    curCode(10) = Array("BAM", "BA", "Bosnia-Herzegovina", "Bosnia and Herzegovina konvertibilna marka")
    curCode(11) = Array("BBD", "BB", "Barbados", "Barbadian dollar")
    curCode(12) = Array("BDT", "BD", "Bangladesh", "Bangladeshi taka")
    curCode(13) = Array("BGN", "BG", "Bulgaria", "Bulgarian lev")
    curCode(14) = Array("BHD", "BH", "Bahrain", "Bahraini dinar")
    curCode(15) = Array("BIF", "BI", "Burundi", "Burundi franc")
    curCode(16) = Array("BMD", "BM", "Bermuda", "Bermudian dollar")
    curCode(17) = Array("BND", "BN", "Brunei Darussalam", "Brunei dollar")
    curCode(18) = Array("BOB", "BO", "Bolivia", "Bolivian boliviano")
    curCode(19) = Array("BRL", "BR", "Brazil", "Brazilian real")
    curCode(20) = Array("BSD", "BS", "Bahamas", "Bahamian dollar")
    curCode(21) = Array("BTN", "BT", "Bhutan", "Bhutanese ngultrum")
    curCode(22) = Array("BWP", "BW", "Botswana", "Botswana pula")
    curCode(23) = Array("BYR", "BY", "Belarus", "Belarusian ruble")
    curCode(24) = Array("BZD", "BZ", "Belize", "Belize dollar")
    curCode(25) = Array("CAD", "CA", "Canada", "Canadian dollar")
    curCode(26) = Array("CDF", "CD", "Congo, Democratic Republic", "Congolese franc")
    curCode(27) = Array("CHF", "CH", "Switzerland", "Swiss franc")
    curCode(28) = Array("CLP", "CL", "Chile", "Chilean peso")
    curCode(29) = Array("CNY", "CN", "China", "Chinese renminbi")
    curCode(30) = Array("COP", "CO", "Colombia", "Colombian peso")
    curCode(31) = Array("CRC", "CR", "Costa Rica", "Costa Rican colon")
    curCode(32) = Array("CUC", "CU", "Cuba", "Cuban peso")
    curCode(33) = Array("CVE", "CV", "Cape Verde", "Cape Verdean escudo")
    curCode(34) = Array("CZK", "CZ", "Czech Republic", "Czech koruna")
    curCode(35) = Array("DJF", "DJ", "Djibouti", "Djiboutian franc")
    curCode(36) = Array("DKK", "DK", "Denmark", "Danish krone")
    curCode(37) = Array("DOP", "DO", "Dominican Republic", "Dominican peso")
    curCode(38) = Array("DZD", "DZ", "Algeria", "Algerian dinar")
    curCode(39) = Array("EEK", "EE", "Estonia", "Estonian kroon")
    curCode(40) = Array("EGP", "EG", "Egypt", "Egyptian pound")
    curCode(41) = Array("ERN", "ER", "Eritrea", "Eritrean nakfa")
    curCode(42) = Array("ETB", "ET", "Ethiopia", "Ethiopian birr")
    curCode(43) = Array("EUR", "EU", "Europe", "European euro")
    curCode(44) = Array("FJD", "FJ", "Fiji", "Fijian dollar")
    curCode(45) = Array("FKP", "FK", "Falkland Islands", "Falkland Islands pound")
    curCode(46) = Array("GBP", "GB", "United Kingdom", "British pound")
    curCode(47) = Array("GEL", "GE", "Georgia", "Georgian lari")
    curCode(48) = Array("GHS", "GH", "Ghana", "Ghanaian cedi")
    curCode(49) = Array("GIP", "GI", "Gibraltar", "Gibraltar pound")
    curCode(50) = Array("GMD", "GM", "Gambia", "Gambian dalasi")
    curCode(51) = Array("GNF", "GN", "Guinea", "Guinean franc")
    curCode(52) = Array("GQE", "GQ", "Equatorial Guinea", "Central African CFA franc")
    curCode(53) = Array("GTQ", "GT", "Guatemala", "Guatemalan quetzal")
    curCode(54) = Array("GYD", "GY", "Guyana", "Guyanese dollar")
    curCode(55) = Array("HKD", "HK", "Hong Kong", "Hong Kong dollar")
    curCode(56) = Array("HNL", "HN", "Honduras", "Honduran lempira")
    curCode(57) = Array("HRK", "HR", "Croatia", "Croatian kuna")
    curCode(58) = Array("HTG", "HT", "Haiti", "Haitian gourde")
    curCode(59) = Array("HUF", "HU", "Hungary", "Hungarian forint")
    curCode(60) = Array("IDR", "ID", "Indonesia", "Indonesian rupiah")
    curCode(61) = Array("ILS", "IL", "Israel", "Israeli new sheqel")
    curCode(62) = Array("INR", "IN", "India", "Indian rupee")
    curCode(63) = Array("IQD", "IQ", "Iraq", "Iraqi dinar")
    curCode(64) = Array("IRR", "IR", "Iran", "Iranian rial")
    curCode(65) = Array("ISK", "IS", "Iceland", "Icelandic króna")
    curCode(66) = Array("JMD", "JM", "Jamaica", "Jamaican dollar")
    curCode(67) = Array("JOD", "JO", "Jordan", "Jordanian dinar")
    curCode(68) = Array("JPY", "JP", "Japan", "Japanese yen")
    curCode(69) = Array("KES", "KE", "Kenya", "Kenyan shilling")
    curCode(70) = Array("KGS", "KG", "Kyrgyzstan", "Kyrgyzstani som")
    curCode(71) = Array("KHR", "KH", "Cambodia", "Cambodian riel")
    curCode(72) = Array("KMF", "KM", "Comoros", "Comorian franc")
    curCode(73) = Array("KPW", "KP", "Korea, North", "North Korean won")
    curCode(74) = Array("KRW", "KR", "Korea, South", "South Korean won")
    curCode(75) = Array("KWD", "KW", "Kuwait", "Kuwaiti dinar")
    curCode(76) = Array("KYD", "KY", "Cayman Islands", "Cayman Islands dollar")
    curCode(77) = Array("KZT", "KZ", "Kazakhstan", "Kazakhstani tenge")
    curCode(78) = Array("LAK", "LA", "Laos", "Lao kip")
    curCode(79) = Array("LBP", "LB", "Lebanon", "Lebanese lira")
    curCode(80) = Array("LKR", "LK", "Sri Lanka", "Sri Lankan rupee")
    curCode(81) = Array("LRD", "LR", "Liberia", "Liberian dollar")
    curCode(82) = Array("LSL", "LS", "Lesotho", "Lesotho loti")
    curCode(83) = Array("LTL", "LT", "Lithuania", "Lithuanian litas")
    curCode(84) = Array("LVL", "LV", "Latvia", "Latvian lats")
    curCode(85) = Array("LYD", "LY", "Libya", "Libyan dinar")
    curCode(86) = Array("MAD", "MA", "Morocco", "Moroccan dirham")
    curCode(87) = Array("MDL", "MD", "Moldova", "Moldovan leu")
    curCode(88) = Array("MGA", "MG", "Madagascar", "Malagasy ariary")
    curCode(89) = Array("MKD", "MK", "Macedonia", "Macedonian denar")
    curCode(90) = Array("MMK", "MM", "Myanmar", "Myanma kyat")
    curCode(91) = Array("MNT", "MN", "Mongolia", "Mongolian tugrik")
    curCode(92) = Array("MOP", "MO", "Macau", "Macanese pataca")
    curCode(93) = Array("MRO", "MR", "Mauritania", "Mauritanian ouguiya")
    curCode(94) = Array("MUR", "MU", "Mauritius", "Mauritian rupee")
    curCode(95) = Array("MVR", "MV", "Maldives", "Maldivian rufiyaa")
    curCode(96) = Array("MWK", "MW", "Malawi", "Malawian kwacha")
    curCode(97) = Array("MXN", "MX", "Mexico", "Mexican peso")
    curCode(98) = Array("MYR", "MY", "Malaysia", "Malaysian ringgit")
    curCode(99) = Array("MZM", "MZ", "Mozambique", "Mozambican metical")
    curCode(100) = Array("NAD", "NA", "Namibia", "Namibian dollar")
    curCode(101) = Array("NGN", "NG", "Nigeria", "Nigerian naira")
    curCode(102) = Array("NIO", "NI", "Nicaragua", "Nicaraguan córdoba")
    curCode(103) = Array("NOK", "NO", "Norway", "Norwegian krone")
    curCode(104) = Array("NPR", "NP", "Nepal", "Nepalese rupee")
    curCode(105) = Array("NZD", "NZ", "New Zealand", "New Zealand dollar")
    curCode(106) = Array("OMR", "OM", "Oman", "Omani rial")
    curCode(107) = Array("PAB", "PA", "Panama", "Panamanian balboa")
    curCode(108) = Array("PEN", "PE", "Peru", "Peruvian nuevo sol")
    curCode(109) = Array("PGK", "PG", "Papua New Guinea", "Papua New Guinean kina")
    curCode(110) = Array("PHP", "PH", "Philippines", "Philippine peso")
    curCode(111) = Array("PKR", "PK", "Pakistan", "Pakistani rupee")
    curCode(112) = Array("PLN", "PL", "Poland", "Polish zloty")
    curCode(113) = Array("PYG", "PY", "Paraguay", "Paraguayan guarani")
    curCode(114) = Array("QAR", "QA", "Qatar", "Qatari riyal")
    curCode(115) = Array("RON", "RO", "Romania", "Romanian leu")
    curCode(116) = Array("RSD", "RS", "Serbia", "Serbian dinar")
    curCode(117) = Array("RUB", "RU", "Russia", "Russian ruble")
    curCode(118) = Array("RWF", "RW", "Rwanda", "Rwandan franc")
    curCode(119) = Array("SAR", "SA", "Saudi Arabia", "Saudi riyal")
    curCode(120) = Array("SBD", "SB", "Solomon Islands", "Solomon Islands dollar")
    curCode(121) = Array("SCR", "SC", "Seychelles", "Seychellois rupee")
    curCode(122) = Array("SDG", "SD", "Sudan", "Sudanese pound")
    curCode(123) = Array("SEK", "SE", "Sweden", "Swedish krona")
    curCode(124) = Array("SGD", "SG", "Singapore", "Singapore dollar")
    curCode(125) = Array("SHP", "SH", "St. Helena", "Saint Helena pound")
    curCode(126) = Array("SLL", "SL", "Sierra Leone", "Sierra Leonean leone")
    curCode(127) = Array("SOS", "SO", "Somalia", "Somali shilling")
    curCode(128) = Array("SRD", "SR", "Suriname", "Surinamese dollar")
    curCode(129) = Array("STD", "ST", "São Tomé and Príncipe", "São Tomé and Príncipe dobra")
    curCode(130) = Array("SYP", "SY", "Syria", "Syrian pound")
    curCode(131) = Array("SZL", "SZ", "Swaziland", "Swazi lilangeni")
    curCode(132) = Array("THB", "TH", "Thailand", "Thai baht")
    curCode(133) = Array("TJS", "TJ", "Tajikistan", "Tajikistani somoni")
    curCode(134) = Array("TMT", "TM", "Turkmenistan", "Turkmen manat")
    curCode(135) = Array("TND", "TN", "Tunisia", "Tunisian dinar")
    curCode(136) = Array("TRY", "TR", "Turkey", "Turkish new lira")
    curCode(137) = Array("TTD", "TT", "Trinidad and Tobago", "Trinidad and Tobago dollar")
    curCode(138) = Array("TWD", "TW", "Taiwan", "New Taiwan dollar")
    curCode(139) = Array("TZS", "TZ", "Tanzania", "Tanzanian shilling")
    curCode(140) = Array("UAH", "UA", "Ukraine", "Ukrainian hryvnia")
    curCode(141) = Array("UGX", "UG", "Uganda", "Ugandan shilling")
    curCode(142) = Array("USD", "US", "United States of America", "United States dollar")
    curCode(143) = Array("USD - RoW", "WW", "Rest of the World", "United States dollar")
    curCode(144) = Array("UYU", "UY", "Uruguay", "Uruguayan peso")
    curCode(145) = Array("UZS", "UZ", "Uzbekistan", "Uzbekistani som")
    curCode(146) = Array("VEB", "VE", "Venezuela", "Venezuelan bolivar")
    curCode(147) = Array("VND", "VN", "Vietnam", "Vietnamese dong")
    curCode(148) = Array("VUV", "VU", "Vanuatu", "Vanuatu vatu")
    curCode(149) = Array("WST", "WS", "Samoa", "Samoan tala")
    curCode(150) = Array("XAF", "XA", "Central Africa", "Central African CFA franc")
    curCode(151) = Array("XCD", "XC", "East Caribbean", "East Caribbean dollar")
    curCode(152) = Array("XDR", "XD", "International Monetary Fund", "Special Drawing Rights")
    curCode(153) = Array("XOF", "XO", "West Africa", "West African CFA franc")
    curCode(154) = Array("XPF", "PF", "French Polynesia", "CFP franc")
    curCode(155) = Array("YER", "YE", "Yemen", "Yemeni rial")
    curCode(156) = Array("ZAR", "ZA", "South Africa", "South African rand")
    curCode(157) = Array("ZMK", "ZM", "Zambia", "Zambian kwacha")
    curCode(158) = Array("ZWR", "ZW", "Zimbabwe", "Zimbabwean dollar")
    
    CurrencyCodes = curCode
End Function

Function CreateSampleExchangeRates() As cExchangeRates
    'This sub simply creates a collection of exchange rates
    'to provide an example of how they are required on the worksheet
    Dim cER As cExchangeRate, cers As cExchangeRates
    
    Set cers = New cExchangeRates
    Set cER = New cExchangeRate

    With cER
        .CurrencyCode = "AU"
        .Region = "Australia"
        .RegionCode = "AUS"
        .ExchangeRate = 1#
        .Month = "August"
        .Year = "2008"
    End With
    cers.Add cER
    
    Set cER = Nothing
    Set cER = New cExchangeRate
    With cER
        .CurrencyCode = "CAN"
        .Region = "Canada"
        .RegionCode = "CA"
        .ExchangeRate = 0.95168
        .Month = "August"
        .Year = "2008"
    End With
    cers.Add cER
    
    Set cER = Nothing
    Set cER = New cExchangeRate
    With cER
        .CurrencyCode = "EUR"
        .Region = "Europe"
        .RegionCode = "EU"
        .ExchangeRate = 1.12291
        .Month = "August"
        .Year = "2008"
    End With
    cers.Add cER
    
    Set cER = Nothing
    Set cER = New cExchangeRate
    With cER
        .CurrencyCode = "GBR"
        .Region = "Great Britain"
        .RegionCode = "GB"
        .ExchangeRate = 1.58013
        .Month = "August"
        .Year = "2008"
    End With
    cers.Add cER
    
    Set cER = Nothing
    Set cER = New cExchangeRate
    With cER
        .CurrencyCode = "JPY"
        .Region = "Japan"
        .RegionCode = "JP"
        .ExchangeRate = 1.25714285714286E-02
        .Month = "August"
        .Year = "2008"
    End With
    cers.Add cER
    
    Set cER = Nothing
    Set cER = New cExchangeRate
    With cER
        .CurrencyCode = "MXN"
        .Region = "Mexico"
        .RegionCode = "MX"
        .ExchangeRate = 7.50992063492064E-02
        .Month = "August"
        .Year = "2008"
    End With
    cers.Add cER
    
    Set cER = Nothing
    Set cER = New cExchangeRate
    With cER
        .CurrencyCode = "USA"
        .Region = "United States of America"
        .RegionCode = "US"
        .ExchangeRate = 1.02005
        .Month = "August"
        .Year = "2008"
    End With
    cers.Add cER
    
    Set cER = Nothing
    Set cER = New cExchangeRate
    With cER
        .CurrencyCode = "USD - RoW"
        .Region = "Rest of the World"
        .RegionCode = "WW"
        .ExchangeRate = 1.02005
        .Month = "August"
        .Year = "2008"
    End With
    cers.Add cER
    
    Set CreateSampleExchangeRates = cers
End Function
