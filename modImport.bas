Attribute VB_Name = "modImport"
Option Compare Database
Option Explicit

Public Sub ImportExcelToAccessDynamicDAO(tableName As String, filePath As String)
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rsAccess As DAO.Recordset
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlSheet As Object
    Dim errorCount As Long
    Dim successCount As Long
    Dim startRow As Long
    Dim startCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim i As Long
    Dim dataRange As Object
    
    ' Initialize counters
    errorCount = 0
    successCount = 0
    
    ' Set reference to current database
    Set db = CurrentDb
    
    ' Validate file path
    If Len(Dir(filePath)) = 0 Then
        LogImport tableName, "Error", "Excel file not found at: " & filePath
        MsgBox "Excel file not found at: " & filePath, vbExclamation
        Exit Sub
    End If
    
    ' Create Excel application object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWorkbook = xlApp.Workbooks.Open(filePath)
    
    ' Use the first sheet
    Set xlSheet = xlWorkbook.Sheets(1)
    
    ' Find "Date" cell to determine top-left corner
    Dim dateCell As Object
    Set dateCell = xlSheet.Cells.Find(What:="Date", LookIn:=-4163, LookAt:=1, MatchCase:=True) ' xlValues, xlWhole
    If dateCell Is Nothing Then
        LogImport tableName, "Error", "Header 'Date' not found in sheet '" & xlSheet.Name & "'"
        GoTo CleanupExcel
    End If
    startRow = dateCell.Row
    startCol = dateCell.Column
    
    ' Validate headers
    Dim expectedHeaders As Variant
    expectedHeaders = Array("Date", "Details", "Account", "Paid In", "Withdrawn")
    For i = 0 To 4
        If Trim(xlSheet.Cells(startRow, startCol + i).Value) <> expectedHeaders(i) Then
            LogImport tableName, "Error", "Header mismatch in column " & _
                Split(xlSheet.Cells(1, startCol + i).Address, "$")(1) & _
                ": Expected '" & expectedHeaders(i) & "', found '" & xlSheet.Cells(startRow, startCol + i).Value & "'"
            errorCount = errorCount + 1
            GoTo CleanupExcel
        End If
    Next i
    
    ' Get the continuous data block
    Set dataRange = xlSheet.Cells(startRow, startCol).CurrentRegion
    lastRow = dataRange.Rows(dataRange.Rows.Count).Row
    lastCol = dataRange.Columns(dataRange.Columns.Count).Column
    
    ' Validate column count
    If lastCol - startCol + 1 < 5 Then
        LogImport tableName, "Error", "Data block has fewer than 5 columns (found " & (lastCol - startCol + 1) & ")"
        errorCount = errorCount + 1
        GoTo CleanupExcel
    End If
    
    ' Validate data rows
    If lastRow <= startRow Then
        LogImport tableName, "Error", "No data found below headers in sheet '" & xlSheet.Name & "'"
        errorCount = errorCount + 1
        GoTo CleanupExcel
    End If
    
    ' Open Access table
    Set rsAccess = db.OpenRecordset(tableName, dbOpenTable)
    
    ' Loop through data rows
    For rowIndex = startRow + 1 To lastRow
    On Error Resume Next
    If AssignFieldsFromExcel(rsAccess, xlSheet, rowIndex, startCol) Then
        successCount = successCount + 1
    Else
        LogImport tableName, "Error", "Error importing row " & rowIndex & ": " & Err.Description
        errorCount = errorCount + 1
        Err.Clear
        rsAccess.CancelUpdate
    End If
    On Error GoTo ErrHandler
Next rowIndex
    
Cleanup:
    If Not rsAccess Is Nothing Then
        rsAccess.Close
        Set rsAccess = Nothing
    End If
    
CleanupExcel:
    If Not xlWorkbook Is Nothing Then
        xlWorkbook.Close False
        Set xlWorkbook = Nothing
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    Set xlSheet = Nothing
    Set dataRange = Nothing
    Set db = Nothing
    
    ' Display summary
    If errorCount > 0 Then
        LogImport tableName, "Error", "Import completed with " & errorCount & " errors and " & successCount & " successful rows."
        MsgBox "Import completed with errors." & vbCrLf & _
               "Successful rows: " & successCount & vbCrLf & _
               "Failed rows: " & errorCount, vbExclamation
    Else
        LogImport tableName, "Success", successCount & " rows imported successfully."
        ' MsgBox not needed here; handled by form
    End If
    Exit Sub
    
ErrHandler:
    LogImport tableName, "Error", "Error in ImportExcelToAccessDynamicDAO: " & Err.Description
    Resume Cleanup
End Sub

Private Function AssignFieldsFromExcel(rs As DAO.Recordset, xlSheet As Object, rowIndex As Long, startCol As Long) As Boolean
    On Error GoTo AssignErr

    rs.AddNew
    rs.Fields("Date") = IIf(IsDate(xlSheet.Cells(rowIndex, startCol).Value), CDate(xlSheet.Cells(rowIndex, startCol).Value), Null)
    rs.Fields("Details") = Nz(xlSheet.Cells(rowIndex, startCol + 1).Value, "")
    rs.Fields("Account") = Nz(xlSheet.Cells(rowIndex, startCol + 2).Value, "")
    rs.Fields("Paid In") = Nz(xlSheet.Cells(rowIndex, startCol + 3).Value, 0)
    rs.Fields("Withdrawn") = Nz(xlSheet.Cells(rowIndex, startCol + 4).Value, 0)
    rs.Update

    AssignFieldsFromExcel = True
    Exit Function

AssignErr:
    rs.CancelUpdate
    AssignFieldsFromExcel = False
End Function

Public Sub RunRefreshTransactions()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim queries() As Variant
    Dim i As Integer
    
    ' Define array of query names
    queries = Array("qapp_TransactionsModified", _
                   "qupp_010_TransMapQuantityBuy", _
                   "qupp_011_TransMapQuantitySale", _
                   "qupp_012_TransMapQuantityPurchase", _
                   "qupp_013_TransMapQuantitySale", _
                   "qupp_014_TransMapQuantityDividend", _
                   "qupp_020_TransMapPrice", _
                   "qupp_031_StripAfterDividend", _
                   "qupp_031_StripAfterDividendLegacy", _
                   "qupp_032_StripAfterQuantity", _
                   "qupp_033_StripBuySell", _
                   "qupp_034_StripAfterPurchaseLegacy", _
                   "qupp_035_StripAfterSaleLegacy", _
                   "qupp_036_StripAfterAutomaticLegacy", _
                   "qupp_040_Interest", _
                   "qupp_050_AdminFee", _
                   "qupp_090_Misc", _
                   "qupp_091_ConsolidatedDetails")
    
    ' Set the database object
    Set db = CurrentDb
    
    ' Loop through queries
    DoCmd.SetWarnings False
    For i = LBound(queries) To UBound(queries)
        DoCmd.OpenQuery queries(i)
    Next i
    DoCmd.SetWarnings True
    
    LogImport "tblTransactions", "Success", "Refresh queries executed successfully."
    ' MsgBox not needed here; handled by form
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    LogImport "tblTransactions", "Error", "Error in RunRefreshTransactions: " & queries(i) & " - " & Err.Description
    DoCmd.SetWarnings True
    MsgBox "Error running refresh: " & queries(i) & " - " & Err.Description, vbCritical, "Error"
    Set db = Nothing
    
End Sub

Public Sub ImportExcelToAccessInvestmentsDAO(tableName As String, filePath As String)
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rsAccess As DAO.Recordset
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlSheet As Object
    Dim errorCount As Long
    Dim successCount As Long
    Dim startRow As Long
    Dim startCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim i As Long
    Dim dataRange As Object
    
    ' Initialize counters
    errorCount = 0
    successCount = 0
    
    ' Set reference to current database
    Set db = CurrentDb
    
    ' Validate file path
    If Len(Dir(filePath)) = 0 Then
        LogImport tableName, "Error", "Excel file not found at: " & filePath
        MsgBox "Excel file not found at: " & filePath, vbExclamation
        Exit Sub
    End If
    
    ' Create Excel application object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWorkbook = xlApp.Workbooks.Open(filePath)
    
    ' Use the first sheet
    Set xlSheet = xlWorkbook.Sheets(1)
    
    ' Find "Investment" cell to determine top-left corner
    Dim investmentCell As Object
    Set investmentCell = xlSheet.Cells.Find(What:="Investment", LookIn:=-4163, LookAt:=1, MatchCase:=True) ' xlValues, xlWhole
    If investmentCell Is Nothing Then
        LogImport tableName, "Error", "Header 'Investment' not found in sheet '" & xlSheet.Name & "'"
        GoTo CleanupExcel
    End If
    startRow = investmentCell.Row
    startCol = investmentCell.Column
    
    ' Validate headers
    Dim expectedHeaders As Variant
    expectedHeaders = Array("Investment", "Identifier", "Quantity Held", "Last Price", "Last Price CCY", _
                           "Value", "Value CCY", "FX Rate", "Last Price (p)", "Value (£)", _
                           "Book Cost", "Book Cost CCY", "Average FX Rate", "Book Cost (£)", "% Change")
    For i = 0 To 14
        If Trim(xlSheet.Cells(startRow, startCol + i).Value) <> expectedHeaders(i) Then
            LogImport tableName, "Error", "Header mismatch in column " & _
                Split(xlSheet.Cells(1, startCol + i).Address, "$")(1) & _
                ": Expected '" & expectedHeaders(i) & "', found '" & xlSheet.Cells(startRow, startCol + i).Value & "'"
            errorCount = errorCount + 1
            GoTo CleanupExcel
        End If
    Next i
    
    ' Get the continuous data block
    Set dataRange = xlSheet.Cells(startRow, startCol).CurrentRegion
    lastRow = dataRange.Rows(dataRange.Rows.Count).Row
    lastCol = dataRange.Columns(dataRange.Columns.Count).Column
    
    ' Validate column count
    If lastCol - startCol + 1 < 15 Then
        LogImport tableName, "Error", "Data block has fewer than 15 columns (found " & (lastCol - startCol + 1) & ")"
        errorCount = errorCount + 1
        GoTo CleanupExcel
    End If
    
    ' Validate data rows
    If lastRow <= startRow Then
        LogImport tableName, "Error", "No data found below headers in sheet '" & xlSheet.Name & "'"
        errorCount = errorCount + 1
        GoTo CleanupExcel
    End If
    
    ' Open Access table
    Set rsAccess = db.OpenRecordset(tableName, dbOpenDynaset)
    
    ' Loop through data rows
    For rowIndex = startRow + 1 To lastRow
        On Error Resume Next
        rsAccess.AddNew
        AssignInvestmentFields rsAccess, xlSheet, rowIndex, startCol
        rsAccess.Update
        If Err.Number <> 0 Then
            LogImport tableName, "Error", "Error importing row " & rowIndex & ": " & Err.Description
            errorCount = errorCount + 1
            Err.Clear
            rsAccess.CancelUpdate
        Else
            successCount = successCount + 1
        End If
        On Error GoTo ErrHandler
    Next rowIndex
    
Cleanup:
    If Not rsAccess Is Nothing Then
        rsAccess.Close
        Set rsAccess = Nothing
    End If
    
CleanupExcel:
    If Not xlWorkbook Is Nothing Then
        xlWorkbook.Close False
        Set xlWorkbook = Nothing
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    Set xlSheet = Nothing
    Set dataRange = Nothing
    Set db = Nothing
    
    ' Display summary
    If errorCount > 0 Then
        LogImport tableName, "Error", "Import completed with " & errorCount & " errors and " & successCount & " successful rows."
        MsgBox "Import completed with errors." & vbCrLf & _
               "Successful rows: " & successCount & vbCrLf & _
               "Failed rows: " & errorCount, vbExclamation
    Else
        LogImport tableName, "Success", successCount & " rows imported successfully."
        ' MsgBox not needed here; handled by form
    End If
    Exit Sub
    
ErrHandler:
    LogImport tableName, "Error", "Error in ImportExcelToAccessInvestmentsDAO: " & Err.Description
    Resume Cleanup
End Sub

Private Sub AssignInvestmentFields(rs As DAO.Recordset, xlSheet As Object, rowIndex As Long, startCol As Long)
    Dim fieldMap As Variant
    Dim i As Long
    Dim fieldName As String
    Dim cellValue As Variant

    fieldMap = Array( _
        "Investment", "Identifier", "Quantity Held", "Last Price", "Last Price CCY", _
        "Value", "Value CCY", "FX Rate", "Last Price (p)", "Value (£)", _
        "Book Cost", "Book Cost CCY", "Average FX Rate", "Book Cost (£)", "% Change" _
    )

    For i = 0 To UBound(fieldMap)
        fieldName = fieldMap(i)
        cellValue = xlSheet.Cells(rowIndex, startCol + i).Value
        If IsEmpty(cellValue) Then
            If VarType(rs.Fields(fieldName).Value) = vbString Then
                rs.Fields(fieldName).Value = ""
            Else
                rs.Fields(fieldName).Value = 0
            End If
        Else
            rs.Fields(fieldName).Value = cellValue
        End If
    Next i
End Sub

Private Sub LogImport(tableName As String, status As String, details As String)
    CurrentDb.Execute _
        "INSERT INTO tblImportLog (TableName, Status, Details, ImportDate) " & _
        "VALUES ('" & tableName & "', '" & status & "', '" & details & "', #" & Now & "#)", dbFailOnError
End Sub

Sub RunRefreshInvestments()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim queries() As Variant
    Dim i As Integer
    
    ' Define array of query names
    queries = Array("qapp_InvestmentsConsolidated") ' Add more queries as needed
        
    ' Set the database object
    Set db = CurrentDb
    
    ' Loop through queries
    DoCmd.SetWarnings False ' Suppress confirmation prompts
    For i = LBound(queries) To UBound(queries)
        DoCmd.OpenQuery queries(i)
    Next i
    DoCmd.SetWarnings True ' Re-enable warnings
    
Cleanup:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    DoCmd.SetWarnings True
    MsgBox "Error running query: " & queries(i) & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume Cleanup
End Sub

