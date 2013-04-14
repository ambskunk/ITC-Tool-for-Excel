Attribute VB_Name = "General"
Option Explicit

Public OptReportFilesFolder As Range
Public OptUsername As Range
Public OptPassword As Range
Public OptVendorID As Range

'General Options
Public bLeftToRight As Boolean

'Download Options
Public PathExistsCancelled As Boolean
Public bOverWriteData As Boolean
Public bDownloadReports As Boolean
Public bDownloadExRates As Boolean
Public bDownloadLatestReport As Boolean
Public bSubFolder As Boolean

'ReportReadInOptions
Public bReadSubFolders As Boolean
Public bReadFiles As Boolean

Function GeneralInitialise() As Boolean
    'Sets up the controls on the worksheet
    Dim objFormCheckBox As Object, objDownloadReports As Object, objExchangeRates As Object
    Dim objSubFolders As Object, objLeftoRight As Object, objReadInSubFolders As Object
    Dim objLatestReport As Object, objOBReadFiles
    
    GeneralInitialise = True
    On Error GoTo BadInitialise
    Application.ScreenUpdating = False
    
    'Input variables
    With ThisWorkbook.Worksheets("Options")
        Set OptUsername = .Cells(5, "P")
        Set OptPassword = .Cells(6, "P")
        Set OptVendorID = .Cells(7, "P")
        Set OptReportFilesFolder = .Cells(9, "P")
        
        Set objFormCheckBox = .CheckBoxes("cboxOverWrite")
        Set objDownloadReports = .CheckBoxes("cboxDownloadReports")
        Set objExchangeRates = .CheckBoxes("cboxExchangeRates")
        Set objSubFolders = .CheckBoxes("cboxSubFolders")
        Set objLeftoRight = .CheckBoxes("cboxLeftToRight")
        Set objReadInSubFolders = .CheckBoxes("cboxReadInSubFolders")
        Set objLatestReport = .CheckBoxes("cbxLatestReport")
        Set objOBReadFiles = .OptionButtons("obIndividualFiles")
        
        If Not objFormCheckBox Is Nothing Then bOverWriteData = IIf(objFormCheckBox.Value = 1, True, False)
        If Not objDownloadReports Is Nothing Then bDownloadReports = IIf(objDownloadReports.Value = 1, True, False)
        If Not objExchangeRates Is Nothing Then bDownloadExRates = IIf(objExchangeRates.Value = 1, True, False)
        If Not objSubFolders Is Nothing Then bSubFolder = IIf(objSubFolders.Value = 1, True, False)
        If Not objLeftoRight Is Nothing Then bLeftToRight = IIf(objLeftoRight.Value = 1, True, False)
        If Not objReadInSubFolders Is Nothing Then bReadSubFolders = IIf(objReadInSubFolders.Value = 1, True, False)
        If Not objLatestReport Is Nothing Then bDownloadLatestReport = IIf(objLatestReport.Value = 1, True, False)
        If Not objOBReadFiles Is Nothing Then bReadFiles = IIf(objOBReadFiles.Value = 1, True, False)
    End With
    
    'Clean up all the objects
    Set objFormCheckBox = Nothing
    Set objDownloadReports = Nothing
    Set objExchangeRates = Nothing
    Set objSubFolders = Nothing
    Set objLeftoRight = Nothing
    Set objReadInSubFolders = Nothing
    Set objLatestReport = Nothing
    Set objOBReadFiles = Nothing
    
    Application.ScreenUpdating = False
    Exit Function
    
BadInitialise:
    GeneralInitialise = False
End Function

Function DirectoryExists(FilePath As String, Title As String, IncludeBaseFileName As Boolean) As String
    Dim fs As Object
    Dim PathLength As Integer, Slash1 As Integer, Slash2 As Integer, iResponse As Integer
    Dim FolderPath As String, BaseFileName As String
    Dim TempPath As String, FileFilter As String, ScreenUpdate As Boolean

    'Store the Screen Updating property state
    ScreenUpdate = Application.ScreenUpdating
    If Not ScreenUpdate Then Application.ScreenUpdating = True

    'Create a file system object
    Set fs = CreateObject("Scripting.FileSystemObject")

    'Set the cancelled state to false
    PathExistsCancelled = False

    'Separate the path into folder path and base filename
    FolderPath = fs.GetParentFolderName(FilePath)
    If Right(FolderPath, 1) = "\" Then FolderPath = Left(FolderPath, Len(FolderPath) - 1)
    BaseFileName = fs.GetFileName(FilePath)
    If BaseFileName = "" Then BaseFileName = "Set by RAMSAS"

    'Loop while the folder path can not be found
    While Not fs.FolderExists(FolderPath)
        PathLength = Len(FolderPath)
        Slash1 = InStr(10, FolderPath, "\")
        Slash2 = InStrRev(FolderPath, "\")

        'Truncate display of file path if necessary
        If PathLength > 60 Then
            TempPath = Left(FolderPath, Slash1) & "..." & Right(FolderPath, PathLength - Slash2 + 1)
        Else
            TempPath = FolderPath
        End If

        'Display error message
        iResponse = MsgBox("Cannot find the following folder:" & vbCr & vbCr & _
                "'" & TempPath & "'" & vbCr & vbCr & _
                "Would you like to search for this folder yourself?", vbExclamation + vbOKCancel, Title)
        If iResponse = 1 Then
            TempPath = Application.GetSaveAsFilename(BaseFileName, FileFilter, , "Select Folder")
            If TempPath <> "False" Then
                FolderPath = fs.GetParentFolderName(TempPath)
                If Right(FolderPath, 1) = "\" Then FolderPath = Left(FolderPath, Len(FolderPath) - 1)
                BaseFileName = fs.GetFileName(TempPath)
            End If
        Else
            PathExistsCancelled = True
            GoTo CancelLine
        End If
    Wend

CancelLine:
    Set fs = Nothing
    DirectoryExists = FolderPath & IIf(IncludeBaseFileName, "\" & BaseFileName, "")
    If Not ScreenUpdate Then Application.ScreenUpdating = False
End Function

Function CutToCharacter(ByVal strPath As String, sChr As String) As String
    If Right$(strPath, 1) <> sChr And Len(strPath) > 0 Then
        CutToCharacter = CutToCharacter(Left$(strPath, Len(strPath) - 1), sChr) + Right$(strPath, 1)
    End If
End Function

Function CutFromCharacter(ByVal strPath As String, sChr As String) As String
    Dim sFile As String
    
    If VBA.InStr(1, strPath, sChr) = 0 Then CutFromCharacter = strPath: Exit Function
    sFile = VBA.Trim(CutToCharacter(strPath, sChr))
    CutFromCharacter = VBA.Trim(VBA.Left$(VBA.Trim(strPath), VBA.Len(VBA.Trim(strPath)) - VBA.Len(sFile) - 2))
End Function

Function Select_File_Or_Files_Mac() As Variant
    'Uses AppleScript to select files on a Mac
    Dim MyPath As String, MyScript As String, MyFiles As String, MySplit As Variant

    On Error Resume Next
    MyPath = MacScript("return (path to documents folder) as String")

    MyScript = "set applescript's text item delimiters to "","" " & vbNewLine & _
            "set theFiles to (choose file of type " & _
          " {""public.TEXT""} " & _
            "with prompt ""Please select a file or files"" default location alias """ & _
            MyPath & """ multiple selections allowed true) as string" & vbNewLine & _
            "set applescript's text item delimiters to """" " & vbNewLine & _
            "return theFiles"

    MyFiles = MacScript(MyScript)
    On Error GoTo 0

    If MyFiles <> "" Then
        MySplit = Split(MyFiles, ",")
        Select_File_Or_Files_Mac = MySplit
    End If
End Function


Function ReturnReportRegionAbbreviated(sReportName As String) As String
    If (sReportName = "Americas") Then ReturnReportRegionAbbreviated = "US": Exit Function
    If (sReportName = "Australia") Then ReturnReportRegionAbbreviated = "AU": Exit Function
    If (sReportName = "Canada") Then ReturnReportRegionAbbreviated = "CA": Exit Function
    If (sReportName = "China") Then ReturnReportRegionAbbreviated = "CN": Exit Function
    If (sReportName = "Denmark") Then ReturnReportRegionAbbreviated = "DK": Exit Function
    If (sReportName = "Euro-Zone") Then ReturnReportRegionAbbreviated = "EU": Exit Function
    If (sReportName = "Hong Kong") Then ReturnReportRegionAbbreviated = "HK": Exit Function
    If (sReportName = "Indonesia") Then ReturnReportRegionAbbreviated = "ID": Exit Function
    If (sReportName = "Japan") Then ReturnReportRegionAbbreviated = "JP": Exit Function
    If (sReportName = "Mexico") Then ReturnReportRegionAbbreviated = "MX": Exit Function
    If (sReportName = "Norway") Then ReturnReportRegionAbbreviated = "NO": Exit Function
    If (sReportName = "New Zealand") Then ReturnReportRegionAbbreviated = "NZ": Exit Function
    If (sReportName = "Russia") Then ReturnReportRegionAbbreviated = "RU": Exit Function
    If (sReportName = "Singapore") Then ReturnReportRegionAbbreviated = "SG": Exit Function
    If (sReportName = "Saudi Arabia") Then ReturnReportRegionAbbreviated = "SA": Exit Function
    If (sReportName = "South Africa") Then ReturnReportRegionAbbreviated = "ZA": Exit Function
    If (sReportName = "Sweden") Then ReturnReportRegionAbbreviated = "SE": Exit Function
    If (sReportName = "Switzerland") Then ReturnReportRegionAbbreviated = "CH": Exit Function
    If (sReportName = "Taiwan") Then ReturnReportRegionAbbreviated = "TW": Exit Function
    If (sReportName = "Turkey") Then ReturnReportRegionAbbreviated = "TR": Exit Function
    If (sReportName = "Rest of World") Then ReturnReportRegionAbbreviated = "WW": Exit Function
    If (sReportName = "United Kingdom") Then ReturnReportRegionAbbreviated = "GB": Exit Function
    If (sReportName = "United Arab Emirates") Then ReturnReportRegionAbbreviated = "AE": Exit Function
    Stop
End Function

Function ReturnAbbreviatedMonth(sTheMonth As String) As String
    If VBA.InStr(1, sTheMonth, "Jan") > 0 Then ReturnAbbreviatedMonth = "01": Exit Function
    If VBA.InStr(1, sTheMonth, "Feb") > 0 Then ReturnAbbreviatedMonth = "02": Exit Function
    If VBA.InStr(1, sTheMonth, "Mar") > 0 Then ReturnAbbreviatedMonth = "03": Exit Function
    If VBA.InStr(1, sTheMonth, "Apr") > 0 Then ReturnAbbreviatedMonth = "04": Exit Function
    If VBA.InStr(1, sTheMonth, "May") > 0 Then ReturnAbbreviatedMonth = "05": Exit Function
    If VBA.InStr(1, sTheMonth, "Jun") > 0 Then ReturnAbbreviatedMonth = "06": Exit Function
    If VBA.InStr(1, sTheMonth, "Jul") > 0 Then ReturnAbbreviatedMonth = "07": Exit Function
    If VBA.InStr(1, sTheMonth, "Aug") > 0 Then ReturnAbbreviatedMonth = "08": Exit Function
    If VBA.InStr(1, sTheMonth, "Sep") > 0 Then ReturnAbbreviatedMonth = "09": Exit Function
    If VBA.InStr(1, sTheMonth, "Oct") > 0 Then ReturnAbbreviatedMonth = "10": Exit Function
    If VBA.InStr(1, sTheMonth, "Nov") > 0 Then ReturnAbbreviatedMonth = "11": Exit Function
    If VBA.InStr(1, sTheMonth, "Dec") > 0 Then ReturnAbbreviatedMonth = "12": Exit Function
End Function

Function ReturnFullMonth(sTheMonth As String) As String
    If VBA.InStr(1, sTheMonth, "Jan") > 0 Then ReturnFullMonth = "January": Exit Function
    If VBA.InStr(1, sTheMonth, "Feb") > 0 Then ReturnFullMonth = "February": Exit Function
    If VBA.InStr(1, sTheMonth, "Mar") > 0 Then ReturnFullMonth = "March": Exit Function
    If VBA.InStr(1, sTheMonth, "Apr") > 0 Then ReturnFullMonth = "April": Exit Function
    If VBA.InStr(1, sTheMonth, "May") > 0 Then ReturnFullMonth = "May": Exit Function
    If VBA.InStr(1, sTheMonth, "Jun") > 0 Then ReturnFullMonth = "June": Exit Function
    If VBA.InStr(1, sTheMonth, "Jul") > 0 Then ReturnFullMonth = "July": Exit Function
    If VBA.InStr(1, sTheMonth, "Aug") > 0 Then ReturnFullMonth = "August": Exit Function
    If VBA.InStr(1, sTheMonth, "Sep") > 0 Then ReturnFullMonth = "September": Exit Function
    If VBA.InStr(1, sTheMonth, "Oct") > 0 Then ReturnFullMonth = "October": Exit Function
    If VBA.InStr(1, sTheMonth, "Nov") > 0 Then ReturnFullMonth = "November": Exit Function
    If VBA.InStr(1, sTheMonth, "Dec") > 0 Then ReturnFullMonth = "December": Exit Function
End Function

Sub SortWorksheets(wbSort As Workbook, bOrderLeftotRight As Boolean)
    'Procedure to sort the month-year worksheets
    Dim ii As Integer, sShtArray() As String
    
    'Get the relevant worksheets from the workbook and then sort them
    sShtArray = GetWorksheets(wbSort)
    sShtArray = SortArray(sShtArray)
    
    'Loop through the sorted array and move the sheets so they are in order
    Application.ScreenUpdating = False
    If bOrderLeftotRight Then
        For ii = LBound(sShtArray, 2) To UBound(sShtArray, 2)
            wbSort.Sheets(sShtArray(2, ii)).Move After:=wbSort.Sheets(wbSort.Sheets.Count)
        Next ii
    Else
        For ii = UBound(sShtArray, 2) To LBound(sShtArray, 2) Step -1
            wbSort.Sheets(sShtArray(2, ii)).Move After:=wbSort.Sheets(wbSort.Sheets.Count)
        Next ii
    End If
End Sub

Function GetWorksheets(wbFromWorkbook As Workbook) As String()
    Dim sMonthNo As String, sYearNo As String
    Dim wkSht As Worksheet, sShtArray() As String
    Dim iCount As Integer
    
    ReDim sShtArray(1 To 2, 1 To wbFromWorkbook.Worksheets.Count)
    For Each wkSht In wbFromWorkbook.Worksheets
        If VBA.Len(wkSht.Name) = 6 Then
            sYearNo = VBA.Right$(wkSht.Name, 2)
            sMonthNo = ReturnAbbreviatedMonth(VBA.Left$(wkSht.Name, 3))
            'If we don't have a number in the last two digits or the month
            If VBA.IsNumeric(sYearNo) And VBA.Len(sMonthNo) Then
                iCount = iCount + 1
                sShtArray(1, iCount) = sYearNo & sMonthNo
                sShtArray(2, iCount) = wkSht.Name
            End If
        End If
    Next wkSht
    
    If iCount > 0 Then
    ReDim Preserve sShtArray(1 To 2, 1 To iCount)
    GetWorksheets = sShtArray
    End If
End Function
