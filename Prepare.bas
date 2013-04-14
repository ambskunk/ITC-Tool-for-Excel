Attribute VB_Name = "Prepare"
Option Explicit

Sub PrepareWorkbook()
    Dim newWB As Workbook
    Dim wsOptions As Worksheet, wsExchangeRates As Worksheet
    Dim objDownload As Object, objReadRpts As Object, cers As cExchangeRates
    
    Set newWB = ThisWorkbook

    'Check to see if we're starting with a freshbook
    #If Win32 Or Win64 Then
        If newWB.Worksheets.Count <> 3 Then GoTo NotFreshWorkbook
        Application.DisplayAlerts = False
        newWB.Worksheets("Sheet3").Delete
        Application.DisplayAlerts = True
    #Else
        If newWB.Worksheets.Count <> 1 Then GoTo NotFreshWorkbook
        newWB.Worksheets.Add
    #End If
    On Error GoTo NotFreshWorkbook
    Set wsOptions = newWB.Worksheets("Sheet1")
    Set wsExchangeRates = newWB.Worksheets("Sheet2")
    On Error GoTo 0

    Application.ScreenUpdating = False

    'Rename the worksheets
    wsOptions.Name = "Options"
    wsExchangeRates.Name = "Exchange Rates"
    
    'Setup the Options worksheet.  This is where the main action starts
    With wsOptions
        'Set up the column widths and row heights
        .Columns("A:B").ColumnWidth = 1#
        .Columns("Q:Q").ColumnWidth = 1#
        .Rows("1:1").RowHeight = 8.25
        
        'Set up the text values
        .Cells(2, "C").Value = "iTunes Connect Financial Reporting Tool"
        .Cells(4, "C").Value = "Settings:"
        .Cells(5, "C").Value = "iTunes Connect Username"
        .Cells(6, "C").Value = "iTunes Connect Password"
        .Cells(7, "C").Value = "iTunes Connect Vendor ID"
        .Cells(9, "C").Value = "Financial Reports Download Folder:"
        
        .Cells(11, "C").Value = "General Options"
        .Cells(12, "C").Value = "        Order month worksheets Left to Right"
        
        .Cells(11, "H").Value = "Download Options"
        .Cells(12, "H").Value = "        Sort reports into sub folders by month"
        .Cells(13, "H").Value = "        Overwrite Existing Data"
        .Cells(14, "H").Value = "        Download Reports"
        .Cells(15, "H").Value = "        Download Exchange Rates"
        .Cells(16, "H").Value = "        Download Latest Month Only"

        .Cells(11, "M").Value = "Text File Read Options"
        .Cells(12, "M").Value = "        Select Text Files to Read"
        .Cells(13, "M").Value = "        Select Entire Folder to Read"
        .Cells(14, "M").Value = "               Include Sub Folders"
        
        .Cells(9, "P").Value = ThisWorkbook.Path
        .Cells(5, "P").HorizontalAlignment = xlRight
        .Cells(6, "P").HorizontalAlignment = xlRight
        .Cells(7, "P").HorizontalAlignment = xlRight
        .Cells(9, "P").HorizontalAlignment = xlRight

        'We want heading bold, right?
        .Cells(2, "C").Font.Bold = True
        .Cells(11, "C").Font.Bold = True
        .Cells(11, "H").Font.Bold = True
        .Cells(11, "M").Font.Bold = True
        .Cells(2, "C").Font.Size = 18
        
        'Make the whole worksheet white
        With .Cells.Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        'Set up the internal fill
        With .Range("H5:P7").Interior
            .Pattern = xlLightUp
            .PatternThemeColor = xlThemeColorAccent3
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0.599963377788629
        End With
        
        With .Range("H9:P9").Interior
            .Pattern = xlLightUp
            .PatternThemeColor = xlThemeColorAccent3
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0.599963377788629
        End With
        
        'Set the border.  Alot of code for a border!
        With .Range("B2:Q19")
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
        
        'Do the controls
        Set objDownload = .Buttons.Add(wsOptions.Cells(18, "H").Left + 2, wsOptions.Cells(18, "H").Top + 2, 80, 23)
        objDownload.OnAction = "LogintoiTunesConnect"
        objDownload.Caption = "Download"
        Set objDownload = Nothing
        
        Set objReadRpts = .Buttons.Add(wsOptions.Cells(18, "M").Left + 2, wsOptions.Cells(18, "M").Top + 2, 80, 23)
        objReadRpts.OnAction = "ReadFromExcelSheet"
        objReadRpts.Caption = "Read Reports"
        Set objReadRpts = Nothing

        Dim objCheckBox As Object
        Set objCheckBox = .CheckBoxes.Add(wsOptions.Cells(12, "C").Left + 1, wsOptions.Cells(12, "C").Top + 1, 14, 16)
        objCheckBox.Caption = ""
        objCheckBox.Name = "cboxLeftToRight"
        objCheckBox.Value = True
        Set objCheckBox = Nothing
        
        Set objCheckBox = .CheckBoxes.Add(wsOptions.Cells(12, "H").Left + 1, wsOptions.Cells(12, "H").Top + 1, 14, 16)
        objCheckBox.Caption = ""
        objCheckBox.Name = "cboxSubFolders"
        Set objCheckBox = Nothing
        
        Set objCheckBox = .CheckBoxes.Add(wsOptions.Cells(13, "H").Left + 1, wsOptions.Cells(13, "H").Top + 1, 14, 16)
        objCheckBox.Caption = ""
        objCheckBox.Name = "cboxOverWrite"
        Set objCheckBox = Nothing
        
        Set objCheckBox = .CheckBoxes.Add(wsOptions.Cells(14, "H").Left + 1, wsOptions.Cells(14, "H").Top + 1, 14, 16)
        objCheckBox.Caption = ""
        objCheckBox.Name = "cboxDownloadReports"
        Set objCheckBox = Nothing
        
        Set objCheckBox = .CheckBoxes.Add(wsOptions.Cells(15, "H").Left + 1, wsOptions.Cells(15, "H").Top + 1, 14, 16)
        objCheckBox.Caption = ""
        objCheckBox.Name = "cboxExchangeRates"
        Set objCheckBox = Nothing
        
        Set objCheckBox = .CheckBoxes.Add(wsOptions.Cells(16, "H").Left + 1, wsOptions.Cells(16, "H").Top + 1, 14, 16)
        objCheckBox.Caption = ""
        objCheckBox.Name = "cbxLatestReport"
        Set objCheckBox = Nothing
        
        Set objCheckBox = .CheckBoxes.Add(wsOptions.Cells(14, "M").Left + 17, wsOptions.Cells(14, "M").Top + 1, 14, 16)
        objCheckBox.Caption = ""
        objCheckBox.Name = "cboxReadInSubFolders"
        Set objCheckBox = Nothing
        
        'Option buttons for selecting individual files or whole folders
        Dim obRadio As Object
        Set obRadio = .OptionButtons.Add(wsOptions.Cells(12, "M").Left + 1, wsOptions.Cells(12, "M").Top + 1, 14, 16)
        obRadio.Name = "obIndividualFiles"
        obRadio.Value = True
        obRadio.Text = ""
        Set obRadio = Nothing
        
        Set obRadio = .OptionButtons.Add(wsOptions.Cells(13, "M").Left + 1, wsOptions.Cells(13, "M").Top + 1, 14, 16)
        obRadio.Name = "obEntireFolder"
        obRadio.Text = ""
        Set obRadio = Nothing
    End With
    
    'Make the whole Exchange Rates worksheet white
    With wsExchangeRates.Cells.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'On the Mac it appears OLE Automation isn't set by default which is required for the collection objects.  Fix that.
    Dim ref As Object
    Set ref = ThisWorkbook.VBProject.References.AddFromGuid("{00020430-0000-0000-C000-000000000046}", 0, 0)
    Set ref = Nothing
    
    'Put some sample exchange rates in the workbook
    Set cers = CreateSampleExchangeRates
    PutExchangeRatesInWorksheet wsExchangeRates, cers
    
    Set newWB = Nothing
    Set wsExchangeRates = Nothing:  Set wsOptions = Nothing
    
    MsgBox "You will need to save this workbook as a '.xlsm' file before some of the formula will work."
    Exit Sub
NotFreshWorkbook:
    MsgBox "This doesn't look like a new workbook." & vbCrLf & "Please start with a fresh workbook."

End Sub

