VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolReports As Collection

Property Get Count() As Integer
    Count = mcolReports.Count
End Property

Property Get Item(ByVal vID As Variant) As cReport
Attribute Item.VB_UserMemId = 0
    Set Item = mcolReports(vID)
End Property

Public Sub Add(ByRef inReport As cReport)
    Dim clsReport As cReport, sKey As String, clsTempReport As cReport
    
    Set clsReport = New cReport
    Set clsReport = inReport
    
    sKey = clsReport.Country & clsReport.StartDate & clsReport.EndDate & clsReport.Month & clsReport.Year & clsReport.AppID & clsReport.LocalTotal & clsReport.SaleType
    
    'Sometimes a report will get downloaded twice (dunno why) so check if the report has already been added.
    If mcolReports.Count > 0 Then
        On Error Resume Next
        Set clsTempReport = mcolReports(sKey)
        On Error GoTo 0
        If Not clsTempReport Is Nothing Then Exit Sub
    End If
        
    mcolReports.Add Item:=clsReport, Key:=sKey
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mcolReports.[_NewEnum]
End Function

Public Sub Terminate()
    Dim clsReport As cReport

    For Each clsReport In mcolReports
        Set clsReport = Nothing
    Next clsReport

    Set mcolReports = Nothing
End Sub

Private Sub Class_Initialize()
    Set mcolReports = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolReports = Nothing
End Sub

