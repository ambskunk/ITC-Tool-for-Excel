VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Taken from http://support.microsoft.com/kb/189981

Public Enum W32F_Errors
  W32F_UNKNOWN_ERROR = 45600
  W32F_FILE_ALREADY_OPEN
  W32F_PROBLEM_OPENING_FILE
  W32F_FILE_ALREADY_CLOSED
  W32F_Problem_seeking
End Enum

Private Const W32F_SOURCE = "Win32File Object"

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const OPEN_ALWAYS = 4
Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_BEGIN = 0, FILE_CURRENT = 1, FILE_END = 2

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Declare Function FormatMessage Lib "kernel32" _
        Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                lpSource As Long, _
                                ByVal dwMessageId As Long, _
                                ByVal dwLanguageId As Long, _
                                ByVal lpBuffer As String, _
                                ByVal nSize As Long, _
                                Arguments As Any) As Long

Private Declare Function ReadFile Lib "kernel32" _
                        (ByVal hFile As Long, _
                         lpBuffer As Any, _
                         ByVal nNumberOfBytesToRead As Long, _
                         lpNumberOfBytesRead As Long, _
                         ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
                        (ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" _
                        (ByVal hFile As Long, _
                         lpBuffer As Any, _
                         ByVal nNumberOfBytesToWrite As Long, _
                         lpNumberOfBytesWritten As Long, _
                         ByVal lpOverlapped As Long) As Long

Private Declare Function CreateFile Lib "kernel32" _
        Alias "CreateFileA" (ByVal lpFileName As String, _
                             ByVal dwDesiredAccess As Long, _
                             ByVal dwShareMode As Long, _
                             ByVal lpSecurityAttributes As Long, _
                             ByVal dwCreationDisposition As Long, _
                             ByVal dwFlagsAndAttributes As Long, _
                             ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" _
                        (ByVal hFile As Long, _
                         ByVal lDistanceToMove As Long, _
                         lpDistanceToMoveHigh As Long, _
                         ByVal dwMoveMethod As Long) As Long

Private Declare Function FlushFileBuffers Lib "kernel32" _
                        (ByVal hFile As Long) As Long

Private hFile As Long, sFName As String, fAutoFlush As Boolean

Public Property Get FileHandle() As Long
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  FileHandle = hFile
End Property

Public Property Get FileName() As String
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  FileName = sFName
End Property

Public Property Get IsOpen() As Boolean
  IsOpen = hFile <> INVALID_HANDLE_VALUE
End Property

Public Property Get AutoFlush() As Boolean
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  AutoFlush = fAutoFlush
End Property

Public Property Let AutoFlush(ByVal NewVal As Boolean)
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  fAutoFlush = NewVal
End Property

Public Sub OpenFile(ByVal sFileName As String)
  If hFile <> INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_OPEN, sFName
  End If
  hFile = CreateFile(sFileName, GENERIC_WRITE Or GENERIC_READ, 0, _
                     0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_PROBLEM_OPENING_FILE, sFileName
  End If
  sFName = sFileName
End Sub

Public Sub CloseFile()
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  CloseHandle hFile
  sFName = ""
  fAutoFlush = False
  hFile = INVALID_HANDLE_VALUE
End Sub

Public Function ReadBytes(ByVal ByteCount As Long) As Variant
Dim BytesRead As Long, Bytes() As Byte
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  ReDim Bytes(0 To ByteCount - 1) As Byte
  ReadFile hFile, Bytes(0), ByteCount, BytesRead, 0
  ReadBytes = Bytes
End Function

Public Sub WriteBytes(DataBytes() As Byte)
Dim fSuccess As Long, BytesToWrite As Long, BytesWritten As Long
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  BytesToWrite = UBound(DataBytes) - LBound(DataBytes) + 1
  fSuccess = WriteFile(hFile, DataBytes(LBound(DataBytes)), _
                       BytesToWrite, BytesWritten, 0)
  If fAutoFlush Then Flush
End Sub

Public Sub Flush()
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  FlushFileBuffers hFile
End Sub

Public Sub SeekAbsolute(ByVal HighPos As Long, ByVal LowPos As Long)
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  LowPos = SetFilePointer(hFile, LowPos, HighPos, FILE_BEGIN)
End Sub

Public Sub SeekRelative(ByVal Offset As Long)
Dim TempLow As Long, TempErr As Long
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  TempLow = SetFilePointer(hFile, Offset, ByVal 0&, FILE_CURRENT)
  If TempLow = -1 Then
    TempErr = Err.LastDllError
    If TempErr Then
      RaiseError W32F_Problem_seeking, "Error " & TempErr & "." & _
                                       vbCrLf & CStr(TempErr)
    End If
  End If
End Sub

Private Sub Class_Initialize()
  hFile = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
  If hFile <> INVALID_HANDLE_VALUE Then CloseHandle hFile
End Sub

Private Sub RaiseError(ByVal ErrorCode As W32F_Errors, _
                       Optional sExtra)
Dim Win32Err As Long, Win32Text As String
  Win32Err = Err.LastDllError
  If Win32Err Then
    Win32Text = vbCrLf & "Error " & Win32Err & vbCrLf & _
                DecodeAPIErrors(Win32Err)
  End If
  Select Case ErrorCode
    Case W32F_FILE_ALREADY_OPEN
      Err.Raise W32F_FILE_ALREADY_OPEN, W32F_SOURCE, _
          "The file '" & sExtra & "' is already open." & Win32Text
    Case W32F_PROBLEM_OPENING_FILE
      Err.Raise W32F_PROBLEM_OPENING_FILE, W32F_SOURCE, _
          "Error opening '" & sExtra & "'." & Win32Text
    Case W32F_FILE_ALREADY_CLOSED
      Err.Raise W32F_FILE_ALREADY_CLOSED, W32F_SOURCE, _
          "There is no open file."
    Case W32F_Problem_seeking
      Err.Raise W32F_Problem_seeking, W32F_SOURCE, _
          "Seek Error." & vbCrLf & sExtra
    Case Else
      Err.Raise W32F_UNKNOWN_ERROR, W32F_SOURCE, _
         "Unknown error." & Win32Text
  End Select
End Sub

Private Function DecodeAPIErrors(ByVal ErrorCode As Long) As String
Dim sMessage As String, MessageLength As Long
  sMessage = Space$(256)
  MessageLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
                                ErrorCode, 0&, sMessage, 256&, 0&)
  If MessageLength > 0 Then
    DecodeAPIErrors = Left(sMessage, MessageLength)
  Else
    DecodeAPIErrors = "Unknown Error."
  End If
End Function

