Attribute VB_Name = "modFunctionsLegacy"
Option Compare Database
Option Explicit

Public Function PopulateQuantityFromSale(details As Variant) As Long
    PopulateQuantityFromSale = ExtractQuantityFromDetails(details, "Sale of", -1)
End Function

Public Function PopulateQuantityFromPurchase(details As Variant) As Long
    PopulateQuantityFromPurchase = ExtractQuantityFromDetails(details, "Purchase of", 1)
End Function

Private Function ExtractQuantityFromDetails(details As Variant, keyword As String, multiplier As Long) As Long
    Dim quantity As Long
    Dim keywordPos As Long
    Dim spacePos As Long
    Dim nextSpacePos As Long
    Dim qtyStr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    quantity = 0
    qtyStr = ""

    If IsNull(details) Or details = "" Then GoTo LogAndExit
    details = Trim(details)

    keywordPos = InStr(1, details, keyword & " ", vbTextCompare)
    If keywordPos > 0 Then
        spacePos = keywordPos + Len(keyword & " ")
        If spacePos <= Len(details) Then
            nextSpacePos = InStr(spacePos, details, " ")
            If nextSpacePos = 0 Then nextSpacePos = Len(details) + 1
            qtyStr = Trim(Mid(details, spacePos, nextSpacePos - spacePos))
            If IsNumeric(qtyStr) Then
                quantity = multiplier * CLng(qtyStr)
            End If
        End If
    End If

LogAndExit:
    On Error Resume Next
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DebugQuantityLog", dbOpenDynaset)
    rs.AddNew
    rs!details = Left(details & "", 255)
    rs!keywordFound = keyword
    rs!QuantityString = qtyStr
    rs!ExtractedQuantity = quantity
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0

    ExtractQuantityFromDetails = quantity
End Function

Public Function StripAfterDividendLegacy(details As Variant) As String
    Dim result As String
    Dim dividendFound As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    result = ExtractSecurityNameDiv(details, "Dividend on ", " shares at ", dividendFound)

    On Error Resume Next
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DebugStripAfterDividend", dbOpenDynaset)
    rs.AddNew
    rs!details = Left(details & "", 255)
    rs!dividendFound = dividendFound
    rs!ModifiedDetails = result
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0

    StripAfterDividendLegacy = result
End Function

Private Function ExtractSecurityNameDiv(details As Variant, actionPrefix As String, suffix As String, foundFlag As String) As String
    Dim result As String
    Dim actionPos As Long
    Dim spacePos As Long
    Dim nextSpacePos As Long
    Dim suffixPos As Long

    result = ""
    foundFlag = "No"

    If IsNull(details) Or details = "" Then
        ExtractSecurityNameDiv = result
        Exit Function
    End If

    result = Trim(details)
    actionPos = InStr(1, result, actionPrefix, vbTextCompare)

    If actionPos = 1 Then
        foundFlag = "Yes"
        spacePos = actionPos + Len(actionPrefix)
        If spacePos <= Len(result) Then
            nextSpacePos = InStr(spacePos, result, " ")
            If nextSpacePos > spacePos Then
                suffixPos = InStr(nextSpacePos, result, suffix, vbTextCompare)
                If suffixPos > nextSpacePos Then
                    result = Trim(Mid(result, nextSpacePos + 1, suffixPos - nextSpacePos - 1))
                Else
                    result = Trim(Mid(result, nextSpacePos + 1))
                End If
            End If
        End If
    End If

    ExtractSecurityNameDiv = result
End Function

Public Function StripAfterPurchaseLegacy(details As Variant) As String
    Dim result As String
    Dim purchaseFound As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    result = ExtractSecurityName(details, "Purchase of ", purchaseFound)

    On Error Resume Next
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DebugStripAfterPurchase", dbOpenDynaset)
    rs.AddNew
    rs!details = Left(details & "", 255)
    rs!purchaseFound = purchaseFound
    rs!ModifiedDetails = result
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0

    StripAfterPurchaseLegacy = result
End Function

Public Function StripAfterSaleLegacy(details As Variant) As String
    Dim result As String
    Dim saleFound As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    result = ExtractSecurityName(details, "Sale of ", saleFound)

    On Error Resume Next
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DebugStripAfterSale", dbOpenDynaset)
    rs.AddNew
    rs!details = Left(details & "", 255)
    rs!saleFound = saleFound
    rs!ModifiedDetails = result
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0

    StripAfterSaleLegacy = result
End Function

Private Function ExtractSecurityName(details As Variant, actionPrefix As String, foundFlag As String) As String
    Dim result As String
    Dim actionPos As Long
    Dim spacePos As Long
    Dim nextSpacePos As Long
    Dim sharesPos As Long

    result = ""
    foundFlag = "No"

    If IsNull(details) Or details = "" Then
        ExtractSecurityName = result
        Exit Function
    End If

    result = Trim(details)
    actionPos = InStr(1, result, actionPrefix, vbTextCompare)

    If actionPos = 1 Then
        foundFlag = "Yes"
        spacePos = actionPos + Len(actionPrefix)
        If spacePos <= Len(result) Then
            nextSpacePos = InStr(spacePos, result, " ")
            If nextSpacePos > spacePos Then
                sharesPos = InStr(nextSpacePos, result, " shares", vbTextCompare)
                If sharesPos > nextSpacePos Then
                    result = Trim(Mid(result, nextSpacePos + 1, sharesPos - nextSpacePos - 1))
                Else
                    result = Trim(Mid(result, nextSpacePos + 1))
                End If
            End If
        End If
    End If

    ExtractSecurityName = result
End Function

Public Function StripAfterAutomaticLegacy(details As Variant) As String
    Const AUTO_PREFIX As String = "Automatic dividend reinvest - purchase "
    Dim result As String
    Dim automaticFound As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ' Initialize
    result = ""
    automaticFound = "No"

    ' Handle null or empty input
    If IsNull(details) Or details = "" Then
        GoTo LogAndExit
    End If

    ' Clean input
    result = Trim(details)

    ' Check if string begins with expected prefix
    If InStr(1, result, AUTO_PREFIX, vbTextCompare) = 1 Then
        automaticFound = "Yes"
        result = ExtractSecurityNameAuto(result, AUTO_PREFIX)
    End If

LogAndExit:
    ' Log debugging information
    On Error Resume Next
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DebugStripAfterAutomatic", dbOpenDynaset)
    rs.AddNew
    rs!details = Left(details & "", 255)
    rs!automaticFound = automaticFound
    rs!ModifiedDetails = result
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0

    StripAfterAutomaticLegacy = result
End Function

Private Function ExtractSecurityNameAuto(fullText As String, prefix As String) As String
    Dim startPos As Long
    Dim quantityEndPos As Long
    Dim sharesPos As Long
    Dim securityName As String

    ' Start after the prefix
    startPos = Len(prefix) + 1

    ' Find end of quantity (first space after prefix)
    quantityEndPos = InStr(startPos, fullText, " ")
    If quantityEndPos = 0 Then
        ExtractSecurityNameAuto = ""
        Exit Function
    End If

    ' Find position of " shares"
    sharesPos = InStr(quantityEndPos + 1, fullText, " shares", vbTextCompare)

    If sharesPos > quantityEndPos Then
        ' Extract string between quantity and " shares"
        securityName = Mid(fullText, quantityEndPos + 1, sharesPos - quantityEndPos - 1)
    Else
        ' If " shares" not found, return string after quantity
        securityName = Mid(fullText, quantityEndPos + 1)
    End If

    ExtractSecurityNameAuto = Trim(securityName)
End Function

