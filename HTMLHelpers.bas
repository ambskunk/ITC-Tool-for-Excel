Attribute VB_Name = "HTMLHelpers"
Option Explicit

'=========================================================================================================================
' Functions used for HTML scrapping.  Ugly Business
'=========================================================================================================================
Function GetArrayofInstancesFromHTML(sHTML As String, sSearchTag As String, sSearchPredicate As String) As String()
    Dim sTagStart As String, sTagEnd As String, sFoundText As String
    Dim iStart As Long, iEnd As Long, iCounter As Long, sOutputArray() As String

    sTagStart = "<" & sSearchTag & " "
    sTagEnd = "/" & sSearchTag & ">"
    If sSearchTag = "input" Then sTagEnd = " />"
    
    iStart = 1:    iCounter = 0
    While iStart > 0
        iStart = VBA.InStr(iStart + 1, sHTML, sTagStart)
        If iStart > 0 Then
            iEnd = VBA.InStr(iStart, sHTML, sTagEnd)
            sFoundText = VBA.Mid$(sHTML, iStart + VBA.Len(sTagStart) - 1, iEnd - (iStart + VBA.Len(sTagStart) - 1))
            
            'If we have set a predicate, then make sure it matches
            If VBA.Len(sSearchPredicate) > 0 Then
                If VBA.InStr(1, sFoundText, sSearchPredicate) = 0 Then sFoundText = ""
            End If
        End If
        
        'If we've found something then chuck it in the array
        If VBA.Len(sFoundText) > 0 Then
            iCounter = iCounter + 1
            ReDim Preserve sOutputArray(1 To iCounter)
            sOutputArray(iCounter) = sFoundText
        End If
    Wend

    GetArrayofInstancesFromHTML = sOutputArray
End Function

Function GetValueForVariable(sHTML As String, sValue As String, Optional bRemoveQuotes As Boolean) As String
    Dim iStart As Integer, iEnd As Integer, sResponse As String
    
    iStart = VBA.InStr(1, sHTML, sValue & "=") + VBA.Len(sValue & "=")
    iEnd = VBA.InStr(iStart + 1, sHTML, """")
    sResponse = VBA.Mid$(sHTML, iStart, iEnd - iStart + 1)
    
    If bRemoveQuotes Then
        If VBA.Left$(sResponse, 1) = """" Then sResponse = VBA.Right$(sResponse, VBA.Len(sResponse) - 1)
        If VBA.Right$(sResponse, 1) = """" Then sResponse = VBA.Left$(sResponse, VBA.Len(sResponse) - 1)
    End If
    
    GetValueForVariable = sResponse
End Function

Function GetInnerText(sString As String) As String
    Dim iStart As Integer, iEnd As Integer, sResponse As String

    iStart = VBA.InStr(1, sString, ">")
    iEnd = VBA.InStr(iStart, sString, "<")
    sResponse = VBA.Mid$(sString, iStart + 1, iEnd - iStart - 1)
    
    GetInnerText = sResponse
End Function

Function GetArrayOfAnInput(sHTML As String) As String()
    ''Gets all the variables for all the inputs in the sent string

    Dim sInputsArray() As String, sTemp As String
    Dim iStart As Integer, iStart2 As Integer, iEnd As Integer, iEnd2 As Integer, iEnd2Old As Integer, iCounter As Integer
    
    
    iStart = 1
    While iStart > 0
        iStart = VBA.InStr(iStart + 1, sHTML, "<input ")
        If iStart > 0 Then
            iEnd = VBA.InStr(iStart, sHTML, """ />")
            If iEnd > 0 Then sTemp = VBA.Mid$(sHTML, iStart + VBA.Len("<input "), iEnd - (iStart + VBA.Len(""" />")) - 2)
            
            'We've found an input so work out all the individual values
            If VBA.Len(sTemp) > 0 Then
                iCounter = 0
                iStart2 = 0
                iEnd2Old = 0
                Do
                    'Loop while we keep finding a =" string
                    iStart2 = VBA.InStr(iStart2 + 1, sTemp, "=""")
                    If iStart2 > 0 Then
                        'Find the quote at the end
                        iEnd2 = VBA.InStr(iStart2 + 2, sTemp, """")
                        If iEnd2 > 0 Then
                            'Add it to the output array
                            iCounter = iCounter + 1
                            ReDim Preserve sInputsArray(1 To 2, 1 To iCounter)
                            sInputsArray(1, iCounter) = VBA.Mid$(sTemp, iEnd2Old + 1, iStart2 - iEnd2Old - 1)
                            sInputsArray(2, iCounter) = VBA.Mid$(sTemp, iStart2 + 2, iEnd2 - iStart2 - 2)
                            iEnd2Old = iEnd2
                        End If
                    End If
                Loop Until iStart2 = 0
            End If
        End If
    Wend

    GetArrayOfAnInput = sInputsArray
End Function

Function ReturnSelectedString(sArray() As String, sWithString As String) As String
    Dim ii As Integer

    For ii = LBound(sArray) To UBound(sArray)
        If VBA.InStr(1, sArray(ii), sWithString) Then
            ReturnSelectedString = sArray(ii)
            Exit Function
        End If
    Next ii
End Function

Function BuildFormString(sArray() As String) As String
    'This function builds a standard HTML web form string from an array of input values
    Dim ii As Integer, sReturnedString As String, sDivider As String
    
    sDivider = "--" & MULTIPART_BOUNDARY

    For ii = LBound(sArray, 2) To UBound(sArray, 2)
        sReturnedString = sReturnedString & sDivider & vbCr & vbLf
        sReturnedString = sReturnedString & "Content-Disposition: form-data; name=" & sArray(2, ii) & vbCr & vbLf & vbCr & vbLf & sArray(1, ii) & vbCr & vbLf
    Next ii
    
    sReturnedString = sReturnedString & sDivider & "--"
    BuildFormString = sReturnedString
End Function

'Function GetParametersFromAJAXString(sHTML As String) As String()
'    Dim lStart As Long, lEnd As Long
'    Dim sMid As String
'    Dim sArray() As String
'
'    lStart = VBA.InStr(1, sHTML, "A4J.AJAX.Submit")
'
'
'    If lStart > 0 Then
'        lStart = VBA.InStr(lStart, sHTML, "(")
'        lEnd = VBA.InStr(lStart, sHTML, ")")
'        sMid = VBA.Mid$(sHTML, lStart + 1, lEnd - lStart - 1)
'        sArray = VBA.Split(sMid, ",")
'
'        GetParametersFromAJAXString = sArray
'    End If
'End Function
'
'Function GetAJAXViewState(sHTML As String) As String
'    Dim lStart As Long, lEnd As Long
'    Dim sMid As String
'
'    lStart = VBA.InStr(1, sHTML, "javax.faces.ViewState")
'    lStart = VBA.InStr(lStart, sHTML, "value=""")
'
'    If lStart > 0 Then
'        lEnd = VBA.InStr(lStart, sHTML, """ />")
'        sMid = VBA.Mid$(sHTML, lStart + VBA.Len("value="""), lEnd - lStart - VBA.Len("value="""))
'        GetAJAXViewState = sMid
'    End If
'
'End Function
