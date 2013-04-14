Attribute VB_Name = "ReadInReports"
Option Explicit

Sub ReadFromExcelSheet()
    'Run this when we want to bring up a file picker
    Call ReadInReports.ReadiTunesReports
End Sub

Sub ReadiTunesReports(Optional vFileName As Variant = False)
    '========================================================================
    ' Author: Andrew Hammonds
    ' Date: November 2012 - April 2013
    ' Description: This sub will take a list of iTunes Connect Financial
    '              Reports (if none are found it'll present a picker to
    '              select) and read in the reports and output to a worksheet
    '========================================================================

    #If Win32 Or Win64 Then
        Dim fs As Object, fsFile As Object
    #End If
    
    Dim iRowPrev As Integer, wkSht As Worksheet, vTotalledData As Variant
    Dim sYear As String, sMonthNo As String, sFile As String, sCurrentWorksheetName As String
    Dim ii As Integer, jj As Integer, kk As Integer, iLine As Integer, vOut As Variant
    Dim vSummary As Variant, bFound As Boolean, sFormula As String, sWorksheets() As String
    Dim colReports As New cReports, clsReport As cReport, wsActive As Worksheet
    
    'Start by initialising
    If Not General.GeneralInitialise Then Exit Sub
    Application.ScreenUpdating = False

    'Get the activesheet
    Set wsActive = ThisWorkbook.ActiveSheet

    'Set the initial output array
    ReDim vOut(1 To 10, 1 To 1)
    
    #If Win32 Or Win64 Then
        'If we're on Windows, use the Scripting FileSystemObject
        Set fs = CreateObject("Scripting.FileSystemObject")
        If VBA.IsNumeric(vFileName) Then
            If bReadFiles Then
                vFileName = Application.GetOpenFilename(",*.txt", , "Select iTunes Report Files", , True)
                If Not IsArray(vFileName) Then Exit Sub
            Else
                vFileName = BrowsePath(OptReportFilesFolder.Value2, "Select Files") 'Application.GetOpenFilename(",*.txt", , "Select iTunes Report Files", , True)
                If vFileName = "False" Then Exit Sub
                vFileName = ListAllFiles(VBA.CStr(vFileName))
            End If
        End If
    #Else
        If VBA.IsNumeric(vFileName) Then
            'If we're on Mac, we'll use a custom function
            vFileName = Select_File_Or_Files_Mac
            If IsEmpty(vFileName) Then Exit Sub
        End If
    #End If
    
    'Loop through all the files specified
    For ii = LBound(vFileName) To UBound(vFileName)
        sFile = CutToCharacter(vFileName(ii), Application.PathSeparator)
        
        'Update the status bar
        On Error GoTo 0
        Application.StatusBar = sFile & " Opened"
                             
        #If Win32 Or Win64 Then
            'Use filescripting object to open the file on Windows
            Set fsFile = fs.OpenTextFile(vFileName(ii), 1, False)
        
            'Loop through the file
            Do While fsFile.AtEndOfStream <> True
                
                'Set up a new report class, read it and add it to the collection
                Set clsReport = ReadTextLineReport(fsFile.ReadLine, sFile)
                If Not clsReport Is Nothing Then colReports.Add clsReport
                Set clsReport = Nothing
            Loop
            
            'Fix up the worksheets array
            Call AddToWorksheetArray(sWorksheets, colReports)
            
            'Clean up
            fsFile.Close
            Set fsFile = Nothing
        #Else
            'This is how do it on the mac
            Dim inputFilePath, inputFile, vLineData As Variant

            inputFile = FreeFile
            inputFilePath = vFileName(ii)
            Open inputFilePath For Input As #inputFile

            While Not EOF(inputFile)
                'Set up a new report class, read it and add it to the collection
                iLine = iLine + 1
                Line Input #inputFile, vLineData

                Set clsReport = ReadTextLineReport(VBA.CStr(vLineData), sFile)
                If Not clsReport Is Nothing Then colReports.Add clsReport
                Set clsReport = Nothing
            Wend
            
            'Fix up the worksheets array
            Call AddToWorksheetArray(sWorksheets, colReports)
            Close #inputFile
        #End If
        
    Next ii
    
    'Now loop through all the worksheets as inputted into the worksheet array
    For ii = LBound(sWorksheets, 2) To UBound(sWorksheets, 2)
        sCurrentWorksheetName = sWorksheets(1, ii) & "-" & sWorksheets(2, ii)
        sYear = sWorksheets(2, ii)
        sMonthNo = ReturnAbbreviatedMonth(sWorksheets(1, ii))
        
        'Go to the sheet for that month
        On Error Resume Next
        Set wkSht = Nothing
        Set wkSht = ThisWorkbook.Worksheets(sCurrentWorksheetName)
        On Error GoTo 0
        
        'If the sheet doesn't exist then create it
        If wkSht Is Nothing Then
            Set wkSht = ThisWorkbook.Worksheets.Add
            wkSht.Name = sCurrentWorksheetName
            
            'Lucky last we'll move the sheet to the end
            wkSht.Move After:=wkSht.Parent.Worksheets(wkSht.Parent.Worksheets.Count - 1)
        Else
            wkSht.UsedRange.ClearContents
        End If
    
        'Output the current month to a variable for dumping into the worksheet
        vOut = OutputReportsForMonth(sWorksheets, ii, colReports)
    
        'Clear out whatever was in the worksheet and layout what we've read in
        wkSht.Cells.ClearContents
        wkSht.Range(wkSht.Cells(1, 1), wkSht.Cells(UBound(vOut, 2), UBound(vOut, 1))).Value = Application.WorksheetFunction.Transpose(vOut)
    
        'Set the column widths
        wkSht.Columns("A:A").ColumnWidth = 10.43
        wkSht.Columns("B:B").ColumnWidth = 10.43
        wkSht.Columns("C:C").ColumnWidth = 13.86
        wkSht.Columns("I:I").ColumnWidth = 9.86
        wkSht.Columns("J:J").ColumnWidth = 10.43
        wkSht.Cells.HorizontalAlignment = xlCenter
    
        'Sort the results
        With wkSht.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("C2:C" & wkSht.UsedRange.Rows.Count), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("D2:D" & wkSht.UsedRange.Rows.Count), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("H2:H" & wkSht.UsedRange.Rows.Count), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            
            .SetRange wkSht.UsedRange
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
        '==============================================================================
        'Now Sort through the output and summarise the results
        '==============================================================================
        ReDim vSummary(1 To 10, 1 To 1)
        
        'Set the headings for the summary
        vSummary(1, 1) = "App Name":        vSummary(2, 1) = "Region":          vSummary(3, 1) = "Currency"
        vSummary(4, 1) = "Units":           vSummary(5, 1) = "Local Price":     vSummary(6, 1) = "Local Total"
        vSummary(7, 1) = "Exchange":        vSummary(8, 1) = "AU$":             vSummary(9, 1) = "Tax"
        vSummary(10, 1) = "Payment"
        
        'Set the read in data from the files to a variable to process
        vOut = wkSht.UsedRange.Value
        
        'Loop through the report summary to add up for the summary
        For kk = LBound(vOut, 1) + 1 To UBound(vOut, 1)
            'Assume we haven't found something new
            bFound = False
            
            'First check to see if there is any info
            If VBA.Len(vOut(kk, 1)) = 0 Then Exit For
        
            'Sort through the summary array to see if there is something new
            For jj = LBound(vSummary, 2) To UBound(vSummary, 2)
                'Check if the app is new and the country is new
                If vOut(kk, 4) = vSummary(2, jj) And vOut(kk, 3) = vSummary(1, jj) And vOut(kk, 9) = vSummary(5, jj) Then
                    bFound = True       'We've found something that already exists so
                    Exit For            'we don't need to keep searching
                End If
            Next jj
            
            If vOut(kk, 8) <> "Promo" Then
                'If we've found something new then add it to the summary
                If Not bFound Then
                    ReDim Preserve vSummary(1 To 10, 1 To UBound(vSummary, 2) + 1)
                    
                    vSummary(1, UBound(vSummary, 2)) = vOut(kk, 3)        'App Name
                    vSummary(2, UBound(vSummary, 2)) = vOut(kk, 4)        'Region
                    vSummary(3, UBound(vSummary, 2)) = vOut(kk, 5)        'Currency
                    vSummary(4, UBound(vSummary, 2)) = vOut(kk, 7)        'Units
                    vSummary(5, UBound(vSummary, 2)) = vOut(kk, 9)        'Local Price
                    vSummary(6, UBound(vSummary, 2)) = vOut(kk, 10)       'Local Total
                Else
                    'Otherwise just add it to whatever is there
                    vSummary(4, jj) = vSummary(4, jj) + vOut(kk, 7)
                    vSummary(6, jj) = vSummary(6, jj) + vOut(kk, 10)
                End If
            End If
            
            'Lets also setup an array to record the ranges of the different apps to make the formulae
            'We've already sorted the Apps so we can just check if there is a different app name
            If kk = 2 Then
                ReDim vRange(1 To 3, 1 To 1)
                Dim iCount As Integer
                iCount = 1
                vRange(1, 1) = 2:           vRange(2, 1) = 2
                vRange(3, 1) = vOut(kk, 3)      'Record the app name
            Else
                If vRange(3, iCount) <> vOut(kk, 3) Then        'If the current app name doesn't match the new one
                    ReDim Preserve vRange(1 To 3, 1 To iCount + 1)
                    iCount = iCount + 1
                    vRange(2, iCount - 1) = kk - 1          'This has to be the end of the previous range
                    vRange(1, iCount) = kk:         vRange(2, iCount) = kk
                    vRange(3, iCount) = vOut(kk, 3)
                End If
            End If
            'End If
        Next kk
        vRange(2, iCount) = kk
        
        With wkSht
            'Output the summarised data
            .Range(wkSht.Cells(1, 12), wkSht.Cells(UBound(vSummary, 2), 11 + UBound(vSummary, 1))).Value = Application.WorksheetFunction.Transpose(vSummary)
            
            'Now get the totals
            vTotalledData = SortTotals(vSummary)
            .Range(wkSht.Cells(UBound(vSummary, 2) + 1, 12), wkSht.Cells(UBound(vSummary, 2) + UBound(vTotalledData, 2), 11 + UBound(vSummary, 1))).Value = Application.WorksheetFunction.Transpose(vTotalledData)
            
            'Sort the results
            With wkSht.Sort
                .SortFields.Clear
                .SortFields.Add Key:=Range("M" & UBound(vSummary, 2) + 1 & ":M" & UBound(vSummary, 2) + UBound(vTotalledData, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .SortFields.Add Key:=Range("N" & UBound(vSummary, 2) + 1 & ":N" & UBound(vSummary, 2) + UBound(vTotalledData, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            
                .SetRange wkSht.Range("L" & UBound(vSummary, 2) + 1 & ":V" & UBound(vSummary, 2) + UBound(vTotalledData, 2))
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
                    
            'Format the Summary
            .Columns("L:L").ColumnWidth = 13.86
            .Columns("P:P").ColumnWidth = 10.43
            .Columns("Q:Q").ColumnWidth = 10.57
            .Columns("P:Q").NumberFormat = "0.00"
            .Columns("R:R").NumberFormat = "0.00000"
            .Columns("S:U").NumberFormat = "0.00"
    
            'Fix up the formulae
            For kk = 2 To UBound(vSummary, 2) + UBound(vTotalledData, 2)
                'We recorded the..
                For jj = LBound(vRange, 2) To UBound(vRange, 2)
                    If vRange(3, jj) = .Cells(kk, "L").Value Then
                        .Cells(kk, "O").Value = "=SUMIFS($G$" & vRange(1, jj) & ":$G$" & vRange(2, jj) & ",$D$" & vRange(1, jj) & ":$D$" & vRange(2, jj) & ",M" & kk & ",$I$" & vRange(1, jj) & ":$I$" & vRange(2, jj) & ",P" & kk & ")"
                        Exit For
                    End If
                Next jj
                .Cells(kk, "Q").Value = "=O" & kk & "*P" & kk
                .Cells(kk, "R").Value = "=VLOOKUP(M" & kk & ",'Exchange Rates'!$A$2:$CC$210,MATCH(MID(CELL(""filename"",A1),FIND(""]"",CELL(""filename"",A1))+1,256),'Exchange Rates'!$B$1:$CC$1,0)+1,FALSE)"
                .Cells(kk, "S").Value = "=if(R" & kk & " = """" , """" , Q" & kk & "*R" & kk & ")"
                If .Cells(kk, "N").Value = "AUD" Or .Cells(kk, "N").Value = "NZD" Then .Cells(kk, "T").Value = "=Round(Q" & kk & " * 0.1 * R" & kk & ",2)"
                If .Cells(kk, "N").Value = "JPY" Then .Cells(kk, "T").Value = "=if(R" & kk & " = """" , """" , S" & kk & " * -0.2)"
                .Cells(kk, "U").Value = "=if(R" & kk & " = """" , """" , Round(S" & kk & " + T" & kk & ",2))"
                
                'Do a Sumif for the totals
                If kk > UBound(vSummary, 2) Then
                    .Cells(kk, "O").Value = "=SUMIF($M2:$M" & UBound(vSummary, 2) & ", M" & kk & ", $O2:$O" & UBound(vSummary, 2) & ")"
                    'Need to adjust the unit price to average out any prices changes
                    sFormula = "=("
                    For jj = 2 To UBound(vSummary, 2)
                        If vSummary(2, jj) = .Cells(kk, "M").Value Then
                            If VBA.Len(sFormula) = 2 Then sFormula = sFormula & "O" & jj & "*P" & jj Else sFormula = sFormula & "+ O" & jj & "*P" & jj
                        End If
                    Next jj
                    sFormula = sFormula & ")/O" & kk
                    .Cells(kk, "P") = sFormula
                End If
            Next kk
            
            'Add in the subtotals.  The range shouldn't have changed so that's easy.
            vSummary = .Range(wkSht.Cells(1, 12), wkSht.Cells(UBound(vSummary, 2) + UBound(vTotalledData, 2), 11 + UBound(vSummary, 1))).Value
            For kk = LBound(vSummary, 1) + 1 To UBound(vSummary, 1)
                If kk = 2 Then iRowPrev = 2
                If kk > 2 And vSummary(kk, 1) <> vSummary(kk - 1, 1) Then
                    .Cells(kk - 1, "V").Value = "=Sum(U" & iRowPrev & ":U" & kk - 1 & ")"
                    iRowPrev = kk
                End If
                If kk = UBound(vSummary, 1) Then .Cells(kk, "V").Value = "=Sum(U" & iRowPrev & ":U" & kk & ")"
            Next kk
        End With
    Next ii
    
    'Finally sort the worksheets in order
    Call SortWorksheets(ThisWorkbook, bLeftToRight)
    
    'Clean up
    #If Win32 Or Win64 Then
        Set fs = Nothing
    #End If
    wsActive.Activate
    Set wsActive = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
CancelLine:
End Sub

Function SortTotals(vDataArray As Variant)
    Dim ii As Integer, jj As Integer, iCount As Integer, sFirstName As String, bFound As Boolean
    Dim vOut() As Variant
    Dim iCorrection As Integer
    
    'Record the first app so we can just record the first set of countries
    sFirstName = vDataArray(LBound(vDataArray, 1), LBound(vDataArray, 2) + 1)
    
    'Count through first rather than keep ReDim'ing the array
    iCount = LBound(vDataArray, 1)
    While vDataArray(LBound(vDataArray, 2), iCount + 1) = sFirstName And iCount < UBound(vDataArray, 2) - 1
        iCount = iCount + 1
        'We need to look out for mid month price changes
        If vDataArray(2, iCount) = vDataArray(2, iCount + 1) Then
            iCorrection = iCorrection + 1
        Else
            ReDim Preserve vOut(1 To 10, 1 To iCount - iCorrection - 1)
            vOut(1, iCount - iCorrection - 1) = "Total"
            vOut(2, iCount - iCorrection - 1) = vDataArray(2, iCount)         'Region
            vOut(3, iCount - iCorrection - 1) = vDataArray(3, iCount)         'Currency
            vOut(5, iCount - iCorrection - 1) = vDataArray(5, iCount)         'Unit Price
        End If
    Wend
    iCount = iCount - 1 - iCorrection
    
    'Now loop through the rest of the entries to find other regions
    For ii = iCount + iCorrection + 2 To UBound(vDataArray, 2)
        bFound = False
        If IsEmpty(vDataArray(2, ii)) Then Exit For
        For jj = LBound(vOut, 2) To UBound(vOut, 2)
            If vOut(2, jj) = vDataArray(2, ii) Then
                bFound = True
                Exit For
            End If
        Next jj
        
        'If we haven't found the region ie it's new
        If Not bFound Then
            ReDim Preserve vOut(1 To 10, 1 To UBound(vOut, 2) + 1)
            vOut(1, UBound(vOut, 2)) = "Total"
            vOut(2, UBound(vOut, 2)) = vDataArray(2, ii)         'Region
            vOut(3, UBound(vOut, 2)) = vDataArray(3, ii)         'Currency
            vOut(5, UBound(vOut, 2)) = vDataArray(5, ii)         'Unit Price
        End If
    Next ii

    SortTotals = vOut
End Function

Function ReadTextLineReport(sTextLine As String, sFileName As String) As cReport
    Dim vTemp As Variant, sType As String
    Dim clsReport As New cReport
    
    If VBA.Left$(sTextLine, 5) <> "Total" And VBA.Left$(sTextLine, 5) <> "Start" Then
        
        vTemp = VBA.Split(sTextLine, vbTab)
        With clsReport
            .AppID = vTemp(10)                      'App ID
            .StartDate = VBA.Format(vTemp(0), "mm/dd/yyyy")                   'Start Date
            .EndDate = VBA.Format(vTemp(1), "mm/dd/yyyy")                     'End Date
            .AppName = vTemp(12)                    'App Name
            .CurrencyType = vTemp(8)                'Currency
            .Country = vTemp(17)                    'Country Code
            .Units = vTemp(5)                       'Units
            sType = ""
            If VBA.Len(vTemp(19)) > 0 Then sType = VBA.UCase(VBA.Left$(vTemp(19), 1)) & VBA.LCase(VBA.Right$(vTemp(19), VBA.Len(vTemp(19)) - 1))
            .SaleType = IIf(sType = "Cr", "Promo", sType)
            .LocalPrice = vTemp(6)                  'Local Price
            .LocalTotal = vTemp(7)                  'Local Total
            .ReadFileName sFileName
        End With
        
        Set ReadTextLineReport = clsReport
    End If
    
    Set clsReport = Nothing
End Function

Sub AddToWorksheetArray(ByRef sWorksheets() As String, ByVal colReports As cReports)
    '=====================================================================
    'This module checks the last report in the collection to see if it's in the list of worksheets that
    'require data.  If not it is added to the list
    '=====================================================================
    Dim sMonthYear As String, iUpper As Integer, bFound As Boolean, jj As Integer
    
    'Workout if we've got this in the list of worksheets
    sMonthYear = VBA.Left$(colReports(colReports.Count).Month, 3) & colReports(colReports.Count).Year
    
    'Loop through to see if we have it in the array
    On Error Resume Next
    iUpper = UBound(sWorksheets, 2)
    On Error GoTo 0
    
    'Assume it's not found
    bFound = False
    
    'If we've already got values then loop through
    If iUpper > 0 Then
        For jj = LBound(sWorksheets, 2) To UBound(sWorksheets, 2)
            If sWorksheets(1, jj) & sWorksheets(2, jj) = sMonthYear Then
                bFound = True
                Exit For
            End If
        Next jj
    Else
        'Otherwise, we don't have amy values so redim for the initial value
        ReDim sWorksheets(1 To 2, 1 To 1)
        sWorksheets(1, UBound(sWorksheets, 2)) = VBA.Left$(colReports(colReports.Count).Month, 3)
        sWorksheets(2, UBound(sWorksheets, 2)) = colReports(colReports.Count).Year
        bFound = True
    End If
    
    'Add it to the array if not found there
    If Not bFound Then
        ReDim Preserve sWorksheets(1 To 2, 1 To UBound(sWorksheets, 2) + 1)
        sWorksheets(1, UBound(sWorksheets, 2)) = VBA.Left$(colReports(colReports.Count).Month, 3)
        sWorksheets(2, UBound(sWorksheets, 2)) = colReports(colReports.Count).Year
    End If
End Sub

Function OutputReportsForMonth(sWorksheet() As String, ii As Integer, colReports As cReports) As Variant
    '=====================================================================
    'Output the report values from the Report Class
    '=====================================================================
    Dim vOut() As Variant, clsReport As cReport, jj As Integer
    
    'Hit up the column names
    ReDim vOut(1 To 10, 1 To 1)
    vOut(1, 1) = "Start Date":      vOut(2, 1) = "End Date":        vOut(3, 1) = "App Name"
    vOut(4, 1) = "Region":          vOut(5, 1) = "Currency":        vOut(6, 1) = "Country"
    vOut(7, 1) = "Units":           vOut(8, 1) = "Type":            vOut(9, 1) = "Local Price"
    vOut(10, 1) = "Local Total"
    
    'Loop through all the reports and if they match then add them to the output variable
    For jj = 1 To colReports.Count
        Set clsReport = colReports(jj)
        
        If clsReport.Month = sWorksheet(1, ii) And clsReport.Year = sWorksheet(2, ii) Then
            ReDim Preserve vOut(1 To 10, 1 To UBound(vOut, 2) + 1)
            
            vOut(1, UBound(vOut, 2)) = clsReport.StartDate
            vOut(2, UBound(vOut, 2)) = clsReport.EndDate
            vOut(3, UBound(vOut, 2)) = clsReport.AppName
            vOut(4, UBound(vOut, 2)) = clsReport.Region
            vOut(5, UBound(vOut, 2)) = clsReport.CurrencyType
            vOut(6, UBound(vOut, 2)) = clsReport.Country
            vOut(7, UBound(vOut, 2)) = clsReport.Units
            vOut(8, UBound(vOut, 2)) = clsReport.SaleType
            vOut(9, UBound(vOut, 2)) = clsReport.LocalPrice
            vOut(10, UBound(vOut, 2)) = clsReport.LocalTotal
        End If
    Next jj

    OutputReportsForMonth = vOut
End Function

Function ListAllFiles(sPath As String) As String()
    '=====================================================================
    'Function to return all the files in a particular folder
    'PC Only
    '=====================================================================
    Dim fs As Object, fsPath As Object, fsFile As Object, fsSubFolder As Object
    Dim sArray() As String, sTemp() As String
    Dim ii As Long, lUBound As Long, lUBound2 As Long

    Set fs = CreateObject("Scripting.FileSystemObject")
    With fs
        Set fsPath = .GetFolder(sPath)
        If fsPath.Files.Count > 0 Then
            ReDim sArray(1 To fsPath.Files.Count)
            For Each fsFile In fsPath.Files
                ii = ii + 1
                sArray(ii) = fsFile.Path
            Next fsFile
            lUBound = UBound(sArray)
        End If
        For Each fsSubFolder In fsPath.SubFolders
            sTemp = ListAllFiles(fsSubFolder.Path)
            lUBound2 = 0
            On Error Resume Next
            lUBound2 = UBound(sTemp)
            On Error GoTo 0
            If lUBound2 > 0 Then
                ReDim Preserve sArray(1 To lUBound + UBound(sTemp))
                For ii = LBound(sTemp) To UBound(sTemp)
                    sArray(lUBound + ii) = sTemp(ii)
                Next ii
                lUBound = UBound(sArray)
            End If
        Next fsSubFolder
    End With
    ListAllFiles = sArray
End Function

#If Win32 Or Win64 Then
Function BrowsePath(InitialFolder As String, Title As String, Optional InitialView As Office.MsoFileDialogView = msoFileDialogViewDetails) As String
    '=======================================================================
    ' Author:       Andrew Hammonds
    ' Date:         March 2013
    ' Description:  Function to return a Path string via the Application.FileDialog
    '               see www.cpearson.com/excel/browsefolder.aspx
    '=======================================================================
    Dim V As Variant
    Dim InitFolder As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title
        .InitialView = InitialView
        
        If Len(InitialFolder) > 0 Then
            If Dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        
        .Show
        
        On Error Resume Next
        Err.Clear
        V = .SelectedItems(1)
        If Err.Number <> 0 Then V = "False"
    End With
    BrowsePath = CStr(V)
End Function
#End If
