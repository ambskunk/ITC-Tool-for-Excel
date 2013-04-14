This Excel VBA code will download iTunes Connect Monthly Financial Reports and breakdown the sales per app.  This is useful for profit sharing where proceeds are required on an app basis.

Specifically, this app will (with appropriate iTunes Connect Credentials):
-Log in to iTunes Connect
-Download Monthly Financial Reports and/or available exchange rates for monthly reports
-Read in reports and sum by region and month
-Tabulate exchange rates
-Present calculated, converted proceeds

It can also read in monthly financial reports already downloaded from the iTunes Connect Website

Limitations:
-There is no way (that I'm aware of) to POST to a URL within the Mac version of Excel (the PC version uses MSXML2.ServerXMLHTTP).  If you know a way (any way from within Excel on Mac then post here http://stackoverflow.com/q/14986015/1733206 and let me know).
-You need to have 7zip installed in order to extract the downloaded zip files.  Need to put some error handling in for that I think....
-The program only downloads the reports for the default vendor ID.  I only have a single vendor ID so haven't implemented downloading from multiple vendors.
-The vendor ID is only used to save the files in the same manner as downloaded from the iTunes Connect website.  Any string can be used here.
-If any particular exchange rate is low (ie Japanese Yen to, well, pretty much anything) there are minor rounding errors, normally only a cent or two.
-There is not much in the way of notifying the user of progress.  I'll do a progress bar one day.

How to use:
-Open Excel and with a new workbook, go to the Visual Basic editor (Alt+F11).
-Import all the modules and class modules into the workbook
-Run the 'PrepareWorkbook' sub in the 'Prepare' module.  This will create two worksheets, 'Options' and 'Exchange Rates'.

-From the Options worksheet, input the iTunes Connect credentials into cells P5 to P7.
-Then either download the reports or read in reports using the 'Download' and 'Read Reports' buttons respectively.

Note: 
-When the workbook is prepared, sample data is entered on the 'Exchange Rates' worksheet.  Not that in order to be correctly used with the report worksheets the columns *must* have the heading "'Feb" (note the single apostrophe).
