VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_dtStartDate As Date
Private m_dtEndDate As Date
Private m_sAppName As String
Private m_sRegion As String
Private m_sCurrencyType As String
Private m_sCountry As String
Private m_lUnits As Long
Private m_sSaleType As String
Private m_dLocalPrice As Double
Private m_dLocalTotal As Double
Private m_sAppID As String
Private m_sMonth As String
Private m_sYear As String

Public Property Get Month() As String
    Month = m_sMonth
End Property

Public Property Let Month(ByVal sMonth As String)
    m_sMonth = sMonth
End Property

Public Property Get StartDate() As Date
    StartDate = m_dtStartDate
End Property

Public Property Let StartDate(ByVal dtStartDate As Date)
    m_dtStartDate = dtStartDate
End Property

Public Property Get EndDate() As Date
    EndDate = m_dtEndDate
End Property

Public Property Let EndDate(ByVal dtEndDate As Date)
    m_dtEndDate = dtEndDate
End Property

Public Property Get AppName() As String
    AppName = m_sAppName
End Property

Public Property Let AppName(ByVal sAppName As String)
    m_sAppName = sAppName
End Property

Public Property Get Region() As String
    Region = m_sRegion
End Property

Public Property Let Region(ByVal sRegion As String)
    m_sRegion = sRegion
End Property

Public Property Get CurrencyType() As String
    CurrencyType = m_sCurrencyType
End Property

Public Property Let CurrencyType(ByVal sCurrencyType As String)
    m_sCurrencyType = sCurrencyType
End Property

Public Property Get Country() As String
    Country = m_sCountry
End Property

Public Property Let Country(ByVal sCountry As String)
    m_sCountry = sCountry
End Property

Public Property Get Units() As Long
    Units = m_lUnits
End Property

Public Property Let Units(ByVal lUnits As Long)
    m_lUnits = lUnits
End Property

Public Property Get SaleType() As String
    SaleType = m_sSaleType
End Property

Public Property Let SaleType(ByVal sSaleType As String)
    m_sSaleType = sSaleType
End Property

Public Property Get LocalPrice() As Double
    LocalPrice = m_dLocalPrice
End Property

Public Property Let LocalPrice(ByVal dLocalPrice As Double)
    m_dLocalPrice = dLocalPrice
End Property

Public Property Get LocalTotal() As Double
    LocalTotal = m_dLocalTotal
End Property

Public Property Let LocalTotal(ByVal dLocalTotal As Double)
    m_dLocalTotal = dLocalTotal
End Property

Public Property Get AppID() As String
    AppID = m_sAppID
End Property

Public Property Let AppID(ByVal sAppID As String)
    m_sAppID = sAppID
End Property

Public Property Get Year() As String
    Year = m_sYear
End Property

Public Property Let Year(ByVal sYear As String)
    m_sYear = sYear
End Property

Sub ReadFileName(sFile As String)
    Dim sMonth As String, sMonthNo As String

    sMonth = VBA.Mid$(sFile, 10, 4)
    sMonthNo = VBA.Left$(sMonth, 2)
    Me.Year = VBA.Right$(sMonth, 2)
    Me.Month = VBA.MonthName(sMonthNo, True)
    Me.Region = VBA.Mid$(sFile, 15, 2)
End Sub
