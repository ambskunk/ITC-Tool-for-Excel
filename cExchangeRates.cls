VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExchangeRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolExchangeRates As Collection
Private m_bOverwriteNew As Boolean

Public Property Get OverwriteNew() As Boolean
    OverwriteNew = m_bOverwriteNew
End Property

Public Property Let OverwriteNew(ByVal bOverwriteNew As Boolean)
    m_bOverwriteNew = bOverwriteNew
End Property

Property Get Count() As Integer
    Count = mcolExchangeRates.Count
End Property

Property Get Item(ByVal vID As Variant) As cExchangeRate
Attribute Item.VB_UserMemId = 0
    Set Item = mcolExchangeRates(vID)
End Property

Public Sub Add(ByRef inExchangeRate As cExchangeRate)
    Dim clsExchangeRate As cExchangeRate, clsTempER As cExchangeRate
    
    Set clsExchangeRate = New cExchangeRate
    Set clsExchangeRate = inExchangeRate
    
    'Check if the exchange rate already exists in the collection
    On Error Resume Next
    Set clsTempER = mcolExchangeRates(clsExchangeRate.RegionCode & clsExchangeRate.Month & clsExchangeRate.Year)
    On Error GoTo 0
    
    'If there is clsTempER is Nothing then we can just add the new one
    If Not clsTempER Is Nothing Then
        'Check if we are overwriting or not and check the value is different
        If clsTempER.ExchangeRate <> clsExchangeRate.ExchangeRate Then
            If Me.OverwriteNew Then
                mcolExchangeRates(clsExchangeRate.RegionCode & clsExchangeRate.Month & clsExchangeRate.Year).ExchangeRate = clsExchangeRate.ExchangeRate
            End If
        End If
    Else
        mcolExchangeRates.Add Item:=clsExchangeRate, Key:=clsExchangeRate.RegionCode & clsExchangeRate.Month & clsExchangeRate.Year
    End If
        
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mcolExchangeRates.[_NewEnum]
End Function

Public Sub Terminate()
    Dim clsExchangeRate As cExchangeRate

    For Each clsExchangeRate In mcolExchangeRates
        Set clsExchangeRate = Nothing
    Next clsExchangeRate

    Set mcolExchangeRates = Nothing
End Sub

Private Sub Class_Initialize()
    Set mcolExchangeRates = New Collection
    m_bOverwriteNew = True
End Sub

Private Sub Class_Terminate()
    Set mcolExchangeRates = Nothing
End Sub

