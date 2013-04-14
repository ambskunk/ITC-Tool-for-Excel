VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExchangeRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_sMonth As String
Private m_sRegionCode As String
Private m_dExchangeRate As Double
Private m_sYear As String
Private m_sRegion As String
Private m_sCurrencyCode As String

Public Property Get Month() As String
    Month = m_sMonth
End Property

Public Property Let Month(ByVal sMonth As String)
    m_sMonth = sMonth
End Property

Public Property Get RegionCode() As String
    RegionCode = m_sRegionCode
End Property

Public Property Let RegionCode(ByVal sRegionCode As String)
    m_sRegionCode = sRegionCode
End Property

Public Property Get ExchangeRate() As Double
    ExchangeRate = m_dExchangeRate
End Property

Public Property Let ExchangeRate(ByVal dExchangeRate As Double)
    m_dExchangeRate = dExchangeRate
End Property

Public Property Get Year() As String
    Year = m_sYear
End Property

Public Property Let Year(ByVal sYear As String)
    m_sYear = sYear
End Property

Public Property Get Region() As String
    Region = m_sRegion
End Property

Public Property Let Region(ByVal sRegion As String)
    m_sRegion = sRegion
End Property

Public Property Get CurrencyCode() As String
    CurrencyCode = m_sCurrencyCode
End Property

Public Property Let CurrencyCode(ByVal sCurrencyCode As String)
    m_sCurrencyCode = sCurrencyCode
End Property
