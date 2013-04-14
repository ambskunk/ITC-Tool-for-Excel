Attribute VB_Name = "ZipExtraction"
Option Explicit

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103

'==================================================================================================
' This code is from http://www.rondebruin.nl/7zipwithexcelunzip.htm.  Thanks Ron!
'==================================================================================================

Public Sub ShellAndWait(ByVal PathName As String, Optional WindowState)
    Dim hProg As Long
    Dim hProcess As Long, ExitCode As Long
    'fill in the missing parameter and execute the program
    If IsMissing(WindowState) Then WindowState = 1
    hProg = Shell(PathName, WindowState)
    'hProg is a "process ID under Win32. To get the process handle:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    Do
        'populate Exitcode variable
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
End Sub

Sub UnZipFile(sFile As Variant, sNameUnZipFolder As String)
    Dim PathZipProgram As String
    Dim FileNameZip As Variant, ShellStr As String

    'Path of the Zip program
    PathZipProgram = "C:\Program Files\7-Zip\"

    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        MsgBox "Please find your copy of 7z.exe and try again"
        Exit Sub
    End If

    FileNameZip = sFile

    'Unzip the files/folders from the zip file in the NameUnZipFolder folder
    If FileNameZip = False Then
        'do nothing
    Else
        'There are a few commands/Switches that you can change in the ShellStr
        'We use x command now to keep the folder stucture, replace it with e if you want only the files
        '-aoa Overwrite All existing files without prompt.
        '-aos Skip extracting of existing files.
        '-aou aUto rename extracting file (for example, name.txt will be renamed to name_1.txt).
        '-aot auto rename existing file (for example, name.txt will be renamed to name_1.txt).
        'Use -r if you also want to unzip the subfolders from the zip file
        'You can add -ppassword if you want to unzip a zip file with password (only 7zip files)
        'Change "*.*" to for example "*.txt" if you only want to unzip the txt files
        'Use "*.xl*" for all Excel files: xls, xlsx, xlsm, xlsb
        ShellStr = PathZipProgram & "7z.exe x -aoa -r" & " " & Chr(34) & FileNameZip & Chr(34) & " -o" & Chr(34) & sNameUnZipFolder & Chr(34) & " " & "*.*"

        ShellAndWait ShellStr, vbHide
    End If
End Sub
