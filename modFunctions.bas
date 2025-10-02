Attribute VB_Name = "modFunctions"
Option Compare Database
Option Explicit

Public Function ExtractQuantityFromBuy(details As Variant) As Long
    ExtractQuantityFromBuy = ExtractQuantityByKeyword(details, "Bought", 1)
End Function

Public Function ExtractQuantityFromSell(details As Variant) As Long
    ExtractQuantityFromSell = ExtractQuantityByKeyword(details, "Sold", -1)
End Function

Private Function ExtractQuantityByKeyword(details As Variant, keyword As String, polarity As Long) As Long
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

    keywordPos = InStr(1, details, keyword, vbTextCompare)
    If keywordPos > 0 Then
        spacePos = keywordPos + Len(keyword)
        If spacePos <= Len(details) And Mid(details, spacePos, 1) = " " Then
            spacePos = spacePos + 1
            nextSpacePos = InStr(spacePos, details, " ")
            If nextSpacePos = 0 Then nextSpacePos = Len(details) + 1
            qtyStr = Trim(Mid(details, spacePos, nextSpacePos - spacePos))
            If IsNumeric(qtyStr) Then quantity = polarity * CLng(qtyStr)
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

    ExtractQuantityByKeyword = quantity
End Function

Public Function ExtractQuantityDividendReinvestment(details As Variant) As Long
    ' Initialize variables
    Dim quantity As Long
    Dim keywordPos As Long
    Dim spacePos As Long
    Dim nextSpacePos As Long
    Dim qtyStr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    ' Initialize quantity and logging variables
    quantity = 0
    qtyStr = ""
    
    ' Handle null or empty input
    If IsNull(details) Or details = "" Then
        GoTo LogAndExit
    End If
    
    ' Trim the input to remove leading/trailing spaces
    details = Trim(details)
    
    ' Find "Automatic dividend reinvest - purchase"
    keywordPos = InStr(1, details, "Automatic dividend reinvest - purchase", vbTextCompare)
    If keywordPos > 0 Then
        ' Find the space after the keyword
        spacePos = keywordPos + Len("Automatic dividend reinvest - purchase")
        If spacePos <= Len(details) Then
            If Mid(details, spacePos, 1) = " " Then
                spacePos = spacePos + 1
                ' Find the next space or end of string
                nextSpacePos = InStr(spacePos, details, " ")
                If nextSpacePos = 0 Then nextSpacePos = Len(details) + 1
                ' Extract the quantity string
                qtyStr = Trim(Mid(details, spacePos, nextSpacePos - spacePos))
                ' Convert to number, ensure it's numeric
                If IsNumeric(qtyStr) Then
                    quantity = CLng(qtyStr) ' Positive for purchase
                End If
            End If
        End If
    End If

LogAndExit:
    ' Log debugging information
    On Error Resume Next
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DebugQuantityLog", dbOpenDynaset)
    rs.AddNew
    rs!details = Left(details & "", 255) ' Truncate to avoid Long Text issues
    rs!keywordFound = "Automatic dividend reinvest - purchase"
    rs!QuantityString = qtyStr
    rs!ExtractedQuantity = quantity
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    
    ' Return the extracted quantity
    ExtractQuantityDividendReinvestment = quantity
End Function

Public Function StripTransactionFee(details As Variant) As String
    Const FEE_LABEL As String = "Online transaction fee "
    Const SUFFIX_BUY As String = " Buy"
    Const SUFFIX_SELL As String = " Sell"

    Dim result As String
    Dim feePos As Long
    Dim suffixPos As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ' Handle null or empty input
    If IsNull(details) Or details = "" Then
        StripTransactionFee = ""
        Exit Function
    End If

    ' Clean input
    result = Trim(details)

    ' Check for "Online transaction fee "
    feePos = InStr(1, result, FEE_LABEL, vbTextCompare)
    If feePos > 0 Then
        result = Trim(Mid(result, feePos + Len(FEE_LABEL)))
    End If

    ' Remove " Buy" or " Sell" from the end (case-insensitive)
    suffixPos = InStr(1, result, SUFFIX_BUY, vbTextCompare)
    If suffixPos > 0 Then
        result = Left(result, suffixPos - 1)
    Else
        suffixPos = InStr(1, result, SUFFIX_SELL, vbTextCompare)
        If suffixPos > 0 Then
            result = Left(result, suffixPos - 1)
        End If
    End If

LogAndExit:
    ' Log debugging information
    On Error Resume Next
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DebugStripTransactionFee", dbOpenDynaset)
    rs.AddNew
    rs!details = Left(details, 255) ' Truncate for Long Text
    rs!ModifiedDetails = result
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0

    StripTransactionFee = result
End Function

Public Function StripAfterQuantity(details As Variant) As String
    Dim result As String
    Dim keyword As String

    If IsNull(details) Or details = "" Then
        StripAfterQuantity = ""
        Exit Function
    End If

    If InStr(1, details, "Bought", vbTextCompare) > 0 Then
        keyword = "Bought"
    ElseIf InStr(1, details, "Sold", vbTextCompare) > 0 Then
        keyword = "Sold"
    Else
        StripAfterQuantity = Trim(details)
        Exit Function
    End If

    result = ExtractAfterKeyword(details, keyword, " @")
    StripAfterQuantity = result
End Function

Public Function StripAfterDividend(details As Variant) As String
    Dim result As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dividendFound As String

    dividendFound = "No"
    result = ""

    If IsNull(details) Or details = "" Then GoTo LogAndExit

    If InStr(1, details, "Dividend: ", vbTextCompare) = 1 Then
        dividendFound = "Yes"
        result = ExtractAfterKeyword(details, "Dividend: ", " (")
    End If

LogAndExit:
    On Error Resume Next
    Set db = CurrentDb
    Set rs = db.OpenRecordset("DebugStripAfterDividend", dbOpenDynaset)
    rs.AddNew
    rs!details = Left(details, 255)
    rs!dividendFound = dividendFound
    rs!ModifiedDetails = result
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0

    StripAfterDividend = result
End Function

Private Function ExtractAfterKeyword(details As Variant, keyword As String, Optional trimAfter As String = " @") As String
    Dim result As String
    Dim keywordPos As Long
    Dim spacePos As Long
    Dim nextSpacePos As Long
    Dim atPos As Long

    If IsNull(details) Or details = "" Then
        ExtractAfterKeyword = ""
        Exit Function
    End If

    result = Trim(details)
    keywordPos = InStr(1, result, keyword, vbTextCompare)

    If keywordPos > 0 Then
        spacePos = keywordPos + Len(keyword)
        If spacePos <= Len(result) Then
            If Mid(result, spacePos, 1) = " " Then
                spacePos = spacePos + 1
                nextSpacePos = InStr(spacePos, result, " ")
                If nextSpacePos = 0 Then nextSpacePos = Len(result) + 1

                If IsNumeric(Trim(Mid(result, spacePos, nextSpacePos - spacePos))) Then
                    result = Trim(Mid(result, nextSpacePos + 1))
                    If trimAfter <> "" Then
                        atPos = InStr(1, result, trimAfter)
                        If atPos > 0 Then
                            result = Trim(Left(result, atPos - 1))
                        End If
                    End If
                End If
            End If
        End If
    End If

    ExtractAfterKeyword = result
End Function
