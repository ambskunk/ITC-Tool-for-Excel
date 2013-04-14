Attribute VB_Name = "DownloadManager"
Option Explicit

Public Const MULTIPART_BOUNDARY As String = "----WebKitFormBoundaryo4y3bJWBcfhN2pyb"
Public Const COL_ONE_PRED As String = "col-1 first"
Public Const COL_TWO_PRED As String = "col-2"
Public Const COL_SHOW_PRED As String = "/itc/images/btn-white-show.png"


Sub LogintoiTunesConnect()
    '==================================================================================================
    ' This code logs into iTunes with the users credentials and downloads the financial reports.
    ' It then unzips then and passes an array of names to ReadInReports.ReadiTunesReports for parsing.
    ' Unfortunately, it only works on the PC since I don't know of an equivalent to the MSXML2 object
    ' If you know, please let me know and even answer at http://stackoverflow.com/q/14986015/1733206
    '==================================================================================================
    Dim urlITCBase As String, urlITC As String, objHTTP As Object
    Dim sPostInformation As String, sLoginURL As String, sPaymentsURL As String
    Dim sLoginPageHTML As String, sLoginResponse As String, sPaymentsResponse As String, sSalesAndTrendsURL As String, sSalesAndTrendsResponse As String
    Dim iStart As Integer, iEnd As Integer, bProceedWithDownload As Boolean, iMonthYearIndex As Integer
    Dim sDivider As String, sWorksheets() As String, sMonthFirst As String
    ReDim sDownloadedAlready(1 To 12) As String
    
    'TODO:
    '1. Get the vendor ID from the Sales and Trends page
    '2. Support multiple vendors
    
    'Start by initialising
    If Not General.GeneralInitialise Then Exit Sub
    
    'Get the current list of months as worksheets
    sWorksheets = GetWorksheets(ThisWorkbook)
    
    'Check for important values
    If VBA.Len(OptUsername.Value) = 0 Or VBA.Len(OptPassword.Value) = 0 Then Exit Sub
    
    'Check the reports folder:
    OptReportFilesFolder.Value = DirectoryExists(OptReportFilesFolder.Value & "\Set Report Files Folder", "Select the reports folder", False)
    If PathExistsCancelled Then Exit Sub
    
    'Set up the URLs to query.
    urlITCBase = "https://itunesconnect.apple.com"
    urlITC = urlITCBase & "/WebObjects/iTunesConnect.woa"
    
    'HTML form strings
    sDivider = "--" & MULTIPART_BOUNDARY & vbCr & vbLf
    
    Application.StatusBar = "Attempting Login"
    
    ''Open up an HTTP to begin downloading the reports
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "POST", urlITC, False
    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHTTP.send ("")
    sLoginPageHTML = objHTTP.responseText
    
    ''Find the string WebObjects/iTunesConnect so that we can grab the URL for the form to post the username and password to
    iStart = VBA.InStr(1, sLoginPageHTML, "/WebObjects/iTunesConnect")
    If iStart = 0 Then
        Application.StatusBar = False
        Call MsgBox("Cannot find the iTunes Connect Login Page", vbOK, "Connection Error")
        Exit Sub
    End If
    
    iEnd = VBA.InStr(iStart, sLoginPageHTML, ">")
    sLoginURL = VBA.Mid$(sLoginPageHTML, iStart, iEnd - iStart - 1)
    
    ''Set up the post information to log into the form
    sPostInformation = "theAccountName=" & OptUsername.Value & "&theAccountPW=" & OptPassword.Value
    sPostInformation = sPostInformation & "&1.Continue.x=39&1.Continue.y=7"
    
    'Reset the MSXML2 object and reload
    Set objHTTP = Nothing
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "POST", urlITCBase & sLoginURL, False
    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHTTP.send (sPostInformation)
    sLoginResponse = objHTTP.responseText
    
    'Check to see if there is a "Sign out" on the page
    If VBA.InStr(1, sLoginResponse, "Sign Out") = 0 Then
        Call MsgBox("Cannot Login to iTunes Connect, please check your username/password", vbOK, "Login Error")
        Exit Sub
    End If
        
    'Look for the sales and trends to get the vendor ID
    iStart = VBA.InStr(1, sLoginResponse, "alt=""Sales and Trends")
    iStart = VBA.InStr(iStart - 170, sLoginResponse, "/WebObjects/iTunesConnect.woa")
    iEnd = VBA.InStr(iStart, sLoginResponse, ">")
    sSalesAndTrendsURL = VBA.Mid$(sLoginResponse, iStart, iEnd - iStart - 1)
    
    'Now navigate to the Sales and Trends page
    objHTTP.Open "POST", urlITCBase & sSalesAndTrendsURL, False
    objHTTP.send ("")
    sSalesAndTrendsResponse = objHTTP.responseText
        
    'Navigate back to the original page
    objHTTP.Open "POST", urlITCBase, False
    objHTTP.send ("")
    sLoginResponse = objHTTP.responseText
    
    'Now look for the Financial Reports page
    iStart = VBA.InStr(1, sLoginResponse, "Payments and Financial Reports")
    If iStart = 0 Then MsgBox "Cannont find 'Payments and Financial Reports'.": GoTo Cleanup
    iStart = VBA.InStr(iStart - 64, sLoginResponse, "/WebObjects/iTunesConnect.woa")            'Need an adjustment because the text is after the link
    iEnd = VBA.InStr(iStart, sLoginResponse, ">")
    sPaymentsURL = VBA.Mid$(sLoginResponse, iStart, iEnd - iStart - 1)
    
    'Now navigate to the Payments and Financial Reports page
    objHTTP.Open "POST", urlITCBase & sPaymentsURL, False
    objHTTP.send ("")
    sPaymentsResponse = objHTTP.responseText
    
    If VBA.Len(sPaymentsResponse) = 0 Then
        Call MsgBox("Cannot access Financial Reports and Payments Page. Please log into iTunes Connect via a browser and ensure no messages require acceptance.", vbOKCancel, "Cannot Find Reports")
        Exit Sub
    End If
    
    'Download the reports
    Dim sParse As String, sEarningsURL As String, sEarningsRepsonse As String
    Dim iStartTemp As Integer, iEndTemp As Integer, ii As Integer, jj As Integer
    Dim sEarningsFormId As String, sEarningsReportPostData As String, sFinancialReportFormURL As String
    Dim sFinancialReportSelects() As String, sTableRowArray() As String, bFinancialsDownloaded As Boolean, sBtnWhiteShow() As String, sBtnShowId As String
    Dim sRegionSelectedValue As String, sMonthSelectedValue As String, sYearSelectedValue As String

    Dim sBodyString As String, sInputsArray() As String, sFinancialReportName As String
    Dim sReportSelectStatement As String, sRegionArray() As String, sRegionSelect As String
    Dim sMonthSelectStatement As String, sMonthArray() As String, sMonthSelect As String
    Dim sYearSelectStatement As String, sYearArray() As String, sYearSelect As String
    Dim sTableRowString As String, sMonthYear() As String, sRegion() As String, sSubmitName As String
    Dim sMonthYearSelected As String, sRegionSelected As String, iReportMatch As Integer
    Dim sPaymentsFormID As String, sPaymentsReportPostData As String
    Dim byteResponse() As Byte
    
    Dim iMonthIndex As Integer, iYearIndex As Integer, iCERCount As Integer

    Dim sDownloadedFiles() As String, iCounter As Integer
    ReDim sDownloadedFiles(1 To 100)

    '================
    ' Exchange Rates
    '================
       
    If bDownloadExRates Then
       
        Application.StatusBar = "Downloading Exchange Rates"
       
        'Find the URL for the form for which the Earnings are displayed
        sParse = "<form name=""mainForm"" enctype=""multipart/form-data"" method=""post"" action="""
        iStart = VBA.InStr(1, sPaymentsResponse, sParse) + VBA.Len(sParse)
        iEnd = VBA.InStr(iStart, sPaymentsResponse, """>")
        sEarningsURL = VBA.Mid$(sPaymentsResponse, iStart, iEnd - iStart)
        
        sParse = "value=""Payments"" name="""
        iStartTemp = VBA.InStr(1, sPaymentsResponse, sParse)
        If iStartTemp = 0 Then
            MsgBox "Cannot find earnings tab."
            GoTo Cleanup
        End If
        iEndTemp = VBA.InStr(iStartTemp, sPaymentsResponse, """ />")
        sPaymentsFormID = VBA.Mid$(sPaymentsResponse, iStartTemp + VBA.Len(sParse), iEndTemp - iStartTemp - VBA.Len(sParse))
        
        'Build up the string for posting the form data to the form URL
        sPaymentsReportPostData = sDivider & "Content-Disposition: form-data; name=""" & sPaymentsFormID & """" & vbCr & vbLf & vbCr & vbLf & "Earnings" & vbCr & vbLf
        sPaymentsReportPostData = sPaymentsReportPostData & "--" & MULTIPART_BOUNDARY & "--" & vbCr & vbLf
        
        'Navigate on over to the Earnings page
        objHTTP.Open "POST", urlITCBase & sEarningsURL, False
        objHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & MULTIPART_BOUNDARY
        objHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_1; en-us) AppleWebKit/531.9 (KHTML, like Gecko) Version/4.0.3 Safari/531.9"
        objHTTP.send (sPaymentsReportPostData)
        sEarningsRepsonse = objHTTP.responseText
        
        Dim cers As New cExchangeRates, bExchangeRate As Boolean
        Dim sEarningsFormURL As String
        
        If Not bOverWriteData Then Set cers = GetExchangeRatesFromWorksheet(ThisWorkbook.Worksheets("Exchange Rates"))
        
        Do
        
            bExchangeRate = False
            
            iCERCount = cers.Count
            Set cers = GetExchangeRatesFromHTML(sEarningsRepsonse, cers)
            If cers.Count > iCERCount Then bExchangeRate = True
            
            sParse = "<form name=""mainForm"" enctype=""multipart/form-data"" method=""post"" action="""
            iStart = VBA.InStr(1, sEarningsRepsonse, sParse) + VBA.Len(sParse)
            iEnd = VBA.InStr(iStart, sEarningsRepsonse, """>")
            sEarningsFormURL = VBA.Mid$(sEarningsRepsonse, iStart, iEnd - iStart)
            
            sFinancialReportSelects = GetArrayofInstancesFromHTML(sEarningsRepsonse, "select", "")
            sMonthSelectStatement = sFinancialReportSelects(1)
            sMonthSelect = GetValueForVariable(sMonthSelectStatement, "name")
            sMonthArray = GetArrayofInstancesFromHTML(sMonthSelectStatement, "option", "")
            sMonthSelectedValue = GetValueForVariable(ReturnSelectedString(sMonthArray, "selected"), "value", True)
    
            sYearSelectStatement = sFinancialReportSelects(2)
            sYearSelect = GetValueForVariable(sYearSelectStatement, "name")
            sYearArray = GetArrayofInstancesFromHTML(sYearSelectStatement, "option", "")
            sYearSelectedValue = GetValueForVariable(ReturnSelectedString(sYearArray, "selected"), "value", True)
            
            'Get the ID of the 'show' button
            sBtnWhiteShow = GetArrayofInstancesFromHTML(sEarningsRepsonse, "input", COL_SHOW_PRED)
            sBtnShowId = GetValueForVariable(sBtnWhiteShow(UBound(sBtnWhiteShow)), "name")  'VBA.Mid$(sPaymentsResponse, iStartTemp + VBA.Len(sParse), iEndTemp - iStartTemp - VBA.Len(sParse))
            sBtnShowId = VBA.Mid$(sBtnShowId, 2, VBA.Len(sBtnShowId) - 2)
            
            iMonthIndex = VBA.Int(sMonthSelectedValue)
            iYearIndex = VBA.Int(sYearSelectedValue)
            
            'Drop the indexes back one to get to the older forms
            If iMonthIndex > 0 Then
                iMonthIndex = iMonthIndex - 1
            ElseIf iYearIndex > 0 Then
                iMonthIndex = 11
                iYearIndex = iYearIndex - 1
            Else
                bExchangeRate = False
            End If
        
            If bDownloadLatestReport Then bExchangeRate = False
        
            'Only proceed if we're still within the form bounds
            If bExchangeRate Then
                sMonthSelectedValue = VBA.CStr(iMonthIndex)
                sYearSelectedValue = VBA.CStr(iYearIndex)
                
                'Reset the form data
                ReDim ReportFormArray(1 To 2, 1 To 4) As String
                ReportFormArray(1, 1) = sMonthSelectedValue
                ReportFormArray(2, 1) = sMonthSelect
                ReportFormArray(1, 2) = sYearSelectedValue
                ReportFormArray(2, 2) = sYearSelect
                ReportFormArray(1, 3) = "10"
                ReportFormArray(2, 3) = """" & sBtnShowId & ".x"""
                ReportFormArray(1, 4) = "5"
                ReportFormArray(2, 4) = """" & sBtnShowId & ".y"""
        
                sBodyString = BuildFormString(ReportFormArray)
                
                objHTTP.Open "POST", urlITCBase & sEarningsFormURL, False
                            objHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & MULTIPART_BOUNDARY
                            objHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_1; en-us) AppleWebKit/531.9 (KHTML, like Gecko) Version/4.0.3 Safari/531.9"
                            objHTTP.send (sBodyString)
        
                sEarningsRepsonse = objHTTP.responseText
            End If
            
        Loop Until bExchangeRate = False
        
        
        
        If Not cers Is Nothing Then Call PutExchangeRatesInWorksheet(ThisWorkbook.Worksheets("Exchange Rates"), cers)

    End If

    '================
    ' Exchange Rates
    '================

    '================
    ' Reports
    '================
    If bDownloadReports Then
        Application.StatusBar = "Downloading Reports"
        DoEvents
        
        'Find the URL for the form for which the Earnings are displayed
        sParse = "<form name=""mainForm"" enctype=""multipart/form-data"" method=""post"" action="""
        iStart = VBA.InStr(1, sPaymentsResponse, sParse) + VBA.Len(sParse)
        iEnd = VBA.InStr(iStart, sPaymentsResponse, """>")
        sEarningsURL = VBA.Mid$(sPaymentsResponse, iStart, iEnd - iStart)
    
        'Find the ID of the control to display all the earnings
        sParse = "value=""Earnings"" name="""
        iStartTemp = VBA.InStr(1, sPaymentsResponse, sParse)
        If iStartTemp = 0 Then
            MsgBox "Cannot find earnings tab."
        End If
        iEndTemp = VBA.InStr(iStartTemp, sPaymentsResponse, """ />")
        sEarningsFormId = VBA.Mid$(sPaymentsResponse, iStartTemp + VBA.Len(sParse), iEndTemp - iStartTemp - VBA.Len(sParse))
    
        'Build up the string for posting the form data to the form URL
        sEarningsReportPostData = sDivider & "Content-Disposition: form-data; name=""" & sEarningsFormId & """" & vbCr & vbLf & vbCr & vbLf & "Earnings" & vbCr & vbLf
        sEarningsReportPostData = sEarningsReportPostData & "--" & MULTIPART_BOUNDARY & "--" & vbCr & vbLf
    
        'Navigate on over to the Earnings page
        objHTTP.Open "POST", urlITCBase & sEarningsURL, False
        objHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & MULTIPART_BOUNDARY
        objHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_1; en-us) AppleWebKit/531.9 (KHTML, like Gecko) Version/4.0.3 Safari/531.9"
        objHTTP.send (sEarningsReportPostData)
        sEarningsRepsonse = objHTTP.responseText
            
        Do
            'Set the flag for the while statement.  We'll keep going as long as there is a financials report
            bFinancialsDownloaded = False
    
            'Find the URL for the form for which the earns are displayed
            'This is the form to select the dates and regions for which to download the reports.
            sParse = "<form name=""mainForm"" enctype=""multipart/form-data"" method=""post"" action="""
            iStart = VBA.InStr(1, sEarningsRepsonse, sParse) + VBA.Len(sParse)
            iEnd = VBA.InStr(iStart, sEarningsRepsonse, """>")
            sFinancialReportFormURL = VBA.Mid$(sEarningsRepsonse, iStart, iEnd - iStart)
    
            'We need to find all the options for the forms.
            sFinancialReportSelects = GetArrayofInstancesFromHTML(sEarningsRepsonse, "select", "")
    
            'And all the table rows so we can get the download buttons
            sTableRowArray = GetArrayofInstancesFromHTML(sEarningsRepsonse, "tr", "class=""values""")
    
            'Split up the selects into regions, months and years
            sReportSelectStatement = sFinancialReportSelects(1)
            sRegionSelect = GetValueForVariable(sReportSelectStatement, "name")
            sRegionArray = GetArrayofInstancesFromHTML(sReportSelectStatement, "option", "")
            sRegionSelectedValue = GetValueForVariable(sRegionArray(1), "value", True)
    
            sMonthSelectStatement = sFinancialReportSelects(2)
            sMonthSelect = GetValueForVariable(sMonthSelectStatement, "name")
            sMonthArray = GetArrayofInstancesFromHTML(sMonthSelectStatement, "option", "")
            sMonthSelectedValue = GetValueForVariable(ReturnSelectedString(sMonthArray, "selected"), "value", True)
    
            sYearSelectStatement = sFinancialReportSelects(3)
            sYearSelect = GetValueForVariable(sYearSelectStatement, "name")
            sYearArray = GetArrayofInstancesFromHTML(sYearSelectStatement, "option", "")
            sYearSelectedValue = GetValueForVariable(ReturnSelectedString(sYearArray, "selected"), "value", True)
    
            'Get the ID of the 'show' button
            sBtnWhiteShow = GetArrayofInstancesFromHTML(sEarningsRepsonse, "input", COL_SHOW_PRED)
            sBtnShowId = GetValueForVariable(sBtnWhiteShow(UBound(sBtnWhiteShow)), "name")  'VBA.Mid$(sPaymentsResponse, iStartTemp + VBA.Len(sParse), iEndTemp - iStartTemp - VBA.Len(sParse))
            sBtnShowId = VBA.Mid$(sBtnShowId, 2, VBA.Len(sBtnShowId) - 2)
            
            'Loop through all the table rows
            For ii = LBound(sTableRowArray) To UBound(sTableRowArray)
                
                sTableRowString = sTableRowArray(ii)
                sMonthYear = GetArrayofInstancesFromHTML(sTableRowString, "td", COL_ONE_PRED)
                sRegion = GetArrayofInstancesFromHTML(sTableRowString, "td", COL_TWO_PRED)
    
                sMonthYearSelected = GetInnerText(sMonthYear(UBound(sMonthYear)))
                
                'We want to check first loop in case we only want the latest month
                If ii = LBound(sTableRowArray) Then sMonthFirst = sMonthYearSelected
                
                'Check if we have a worksheet for the month already
                bProceedWithDownload = True
                If Not bOverWriteData Then
                    On Error Resume Next
                    iMonthYearIndex = 0
                    iMonthYearIndex = Application.WorksheetFunction.Match(VBA.Left$(sMonthYearSelected, 3) & "-" & VBA.Right$(sMonthYearSelected, 2), Application.WorksheetFunction.Index(sWorksheets, 2, 0), 0)
                    On Error GoTo 0
                    
                    If iMonthYearIndex > 0 Then bProceedWithDownload = False
                End If
                
                'Check if we only wanted to download the latest month
                If bDownloadLatestReport Then
                    If sMonthYearSelected <> sMonthFirst Then
                        bProceedWithDownload = False
                        bFinancialsDownloaded = False
                    End If
                End If
                
                If Len(Join(sRegion, "")) = 0 Then bProceedWithDownload = False
                
                If bProceedWithDownload Then
                    'Only bother getting the region if we are to proceed with the download
                    sRegionSelected = GetInnerText(sRegion(UBound(sRegion)))
            
                    'Find the id of the button required to download the report
                    sInputsArray = GetArrayOfAnInput(sTableRowString)
                    For jj = LBound(sInputsArray, 2) To UBound(sInputsArray, 2)
                        If VBA.Trim$(sInputsArray(1, jj)) = "name" Then sSubmitName = sInputsArray(2, jj): Exit For
                    Next jj
        
                    If (VBA.Len(sRegionSelected) > 0 And VBA.Len(sSubmitName) > 0) Then
                        sFinancialReportName = OptVendorID.Value & "_" & ReturnAbbreviatedMonth(sMonthYearSelected) & VBA.Right$(sMonthYearSelected, 2) & "_" & ReturnReportRegionAbbreviated(sRegionSelected)
        
                        'Check to make sure we haven't downloaded the report already
                        On Error Resume Next
                        iReportMatch = 0
                        iReportMatch = Application.WorksheetFunction.Match(OptReportFilesFolder.Value & IIf(VBA.Right$(OptReportFilesFolder.Value, 1) <> "\", "\", "") & sFinancialReportName & ".txt", sDownloadedFiles, False)
                        On Error GoTo 0
        
                        'If we don't find a match then it's already been downloaded
                        If iReportMatch = 0 Then
        
                            'Check if we are actually able to download the report
                            If VBA.Len(sFinancialReportName) > 0 Then
                                ReDim ReportFormArray(1 To 2, 1 To 5) As String
            
                                ReportFormArray(1, 1) = sRegionSelectedValue
                                ReportFormArray(2, 1) = sRegionSelect
                                ReportFormArray(1, 2) = sMonthSelectedValue
                                ReportFormArray(2, 2) = sMonthSelect
                                ReportFormArray(1, 3) = sYearSelectedValue
                                ReportFormArray(2, 3) = sYearSelect
                                ReportFormArray(1, 4) = "10"
                                ReportFormArray(2, 4) = """" & sSubmitName & ".x"""
                                ReportFormArray(1, 5) = "5"
                                ReportFormArray(2, 5) = """" & sSubmitName & ".y"""
            
                                sBodyString = BuildFormString(ReportFormArray)
            
                                objHTTP.Open "POST", urlITCBase & sFinancialReportFormURL, False
                                objHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & MULTIPART_BOUNDARY
                                objHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_1; en-us) AppleWebKit/531.9 (KHTML, like Gecko) Version/4.0.3 Safari/531.9"
                                objHTTP.send (sBodyString)
            
                                byteResponse = objHTTP.responseBody
                                If Application.WorksheetFunction.Sum(byteResponse) > 0 Then
                                    bFinancialsDownloaded = True
                                    iCounter = iCounter + 1
                                    
                                    'We need to add a check if the number is over 100 to redim preserve
                                    If iCounter Mod 100 = 0 Then ReDim Preserve sDownloadedFiles(1 To UBound(sDownloadedFiles) + 100)
                                    sDownloadedFiles(iCounter) = SaveBytesToFile(byteResponse, OptReportFilesFolder.Value, sFinancialReportName & ".txt.gz", True, sMonthYearSelected, bSubFolder)
                                End If
                            End If
                        End If
                    End If
                End If
            Next ii
                        
            'Check to see if we only wanted the first month
            If bDownloadLatestReport And Not bFinancialsDownloaded Then
                'Do nothing
            Else
                'We now need loop back through the forms to find the old reports
                iMonthIndex = VBA.Int(sMonthSelectedValue)
                iYearIndex = VBA.Int(sYearSelectedValue)
        
                'Drop the indexes back one to get to the older forms
                If iMonthIndex < 3 And iYearIndex = 0 Then
                    iMonthIndex = 2
                ElseIf iMonthIndex > 2 Then
                    iMonthIndex = iMonthIndex - 3
                ElseIf iYearIndex > 0 Then
                    iMonthIndex = iMonthIndex + 9
                    iYearIndex = iYearIndex - 1
                Else
                    bFinancialsDownloaded = False
                End If
                
                'If we are not overwriting data then we'll keep checking all the reports
                If Not bOverWriteData Then
                    If iMonthIndex > 0 Or iYearIndex > 0 Then
                        bFinancialsDownloaded = True
                    Else
                        bFinancialsDownloaded = False
                    End If
                End If
            End If
            
            'Only proceed if we're still within the form bounds
            If bFinancialsDownloaded Then
                sMonthSelectedValue = VBA.CStr(iMonthIndex)
                sYearSelectedValue = VBA.CStr(iYearIndex)
                
                'Reset the form data
                ReDim ReportFormArray(1 To 2, 1 To 5) As String
                ReportFormArray(1, 1) = sRegionSelectedValue
                ReportFormArray(2, 1) = sRegionSelect
                ReportFormArray(1, 2) = sMonthSelectedValue
                ReportFormArray(2, 2) = sMonthSelect
                ReportFormArray(1, 3) = sYearSelectedValue
                ReportFormArray(2, 3) = sYearSelect
                ReportFormArray(1, 4) = "10"
                ReportFormArray(2, 4) = """" & sBtnShowId & ".x"""
                ReportFormArray(1, 5) = "5"
                ReportFormArray(2, 5) = """" & sBtnShowId & ".y"""
        
                sBodyString = BuildFormString(ReportFormArray)
                
                objHTTP.Open "POST", urlITCBase & sFinancialReportFormURL, False
                            objHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & MULTIPART_BOUNDARY
                            objHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_1; en-us) AppleWebKit/531.9 (KHTML, like Gecko) Version/4.0.3 Safari/531.9"
                            objHTTP.send (sBodyString)
        
                sEarningsRepsonse = objHTTP.responseText
            End If
            
        Loop Until bFinancialsDownloaded = False
    
        'Finally read in the reports
        If iCounter > 0 Then
            ReDim Preserve sDownloadedFiles(1 To iCounter)
            If Len(Join(sDownloadedFiles)) > 0 Then Call ReadiTunesReports(sDownloadedFiles)
        End If
    End If

Cleanup:
    Set objHTTP = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Function SaveBytesToFile(Bytes() As Byte, sPath As String, sFile As String, bUnZip As Boolean, sMonthYear As String, bSaveIntoSubFolder As Boolean) As String
    Dim cFile As New clsFileManager, sLocation As String, sYYMM As String

    'Check the path ending
    If VBA.Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    'Do a bit more work if we're saving into subfolders
    If bSaveIntoSubFolder Then
        sYYMM = VBA.Right$(sMonthYear, 2) & ReturnAbbreviatedMonth(VBA.Left$(sMonthYear, 3)) & " " & ReturnFullMonth(VBA.Left$(sMonthYear, 3)) & " " & VBA.Right$(sMonthYear, 4) & "\"
        sPath = sPath & sYYMM
        
        'Check if the folder exists, if not create it
        If Dir(sPath, vbDirectory) = vbNullString Then
            'Folder will be in the format YYMM Mmm yyyy
            MkDir sPath
        End If
    End If
    
    ' Get the final file location
    sLocation = sPath & sFile

    ' Delete any existing file.
    On Error Resume Next
    Kill sLocation
    On Error GoTo 0
    
    'Do all the writing
    cFile.OpenFile sLocation
    cFile.WriteBytes Bytes
    cFile.CloseFile
    
    If bUnZip Then Call UnZipFile(sLocation, sPath)
    
    Set cFile = Nothing
    Kill sLocation
    
    SaveBytesToFile = VBA.Left$(sLocation, VBA.Len(sLocation) - 3)
End Function

'========================================================================
' TESTING FUNCTIONS
'========================================================================
Sub SaveStringToFile(sString As String, sFile As String)
    Dim hFile As Long
    hFile = FreeFile
    Open sFile For Output As #hFile
    Print #hFile, sString
    Close #hFile
End Sub

Function OpenTextFileToString2(ByVal strFile As String) As String
    ' RB Smissaert - Author
    Dim hFile As Long
    hFile = FreeFile
    Open strFile For Input As #hFile
    OpenTextFileToString2 = Input$(LOF(hFile), hFile)
    Close #hFile
End Function
'========================================================================
' END TESTING FUNCTIONS
'========================================================================
