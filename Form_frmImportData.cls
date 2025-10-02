VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
    ' Populate the table dropdown with valid tables from tblImportConfig
    Me.cboTable.RowSource = "SELECT TableName FROM tblImportConfig"
    Me.cboTable.Requery
End Sub

Private Sub cmdBrowse_Click()
    ' Open file picker for Excel files
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    fd.Filters.Add "Excel Files", "*.xlsx;*.xls"
    If fd.Show Then
        Me.txtFilePath = fd.SelectedItems(1)
    End If
    Set fd = Nothing
End Sub

Private Sub cmdImport_Click()
    On Error GoTo ErrHandler
    
    Dim tableName As String
    Dim filePath As String
    Dim importSubName As String
    Dim refreshSubName As String
    Dim deleteQueryName As String
    Dim secondaryDeleteQuery As String
    Dim rs As DAO.Recordset
    
    ' Validate inputs
    tableName = Nz(Me.cboTable, "")
    filePath = Nz(Me.txtFilePath, "")
    
    If tableName = "" Then
        MsgBox "Please select a target table.", vbExclamation
        Exit Sub
    End If
    If filePath = "" Or Dir(filePath) = "" Then
        MsgBox "Please provide a valid Excel file path.", vbExclamation
        Exit Sub
    End If
    
    ' Look up the import, refresh, and delete query names from tblImportConfig
    Set rs = CurrentDb.OpenRecordset("SELECT VBASubName, RefreshSubName, DeleteQueryName, SecondaryDeleteQuery FROM tblImportConfig WHERE TableName = '" & tableName & "'")
    If Not rs.EOF Then
        importSubName = rs!VBASubName
        refreshSubName = Nz(rs!refreshSubName, "")
        deleteQueryName = Nz(rs!deleteQueryName, "")
        secondaryDeleteQuery = Nz(rs!secondaryDeleteQuery, "")
    Else
        MsgBox "No import procedure defined for table: " & tableName, vbExclamation
        rs.Close
        Exit Sub
    End If
    rs.Close
    Set rs = Nothing
    
    ' Execute the primary delete query if defined
    If deleteQueryName <> "" Then
        On Error Resume Next
        DoCmd.SetWarnings False
        DoCmd.OpenQuery deleteQueryName
        DoCmd.SetWarnings True
        If Err.Number <> 0 Then
            LogImport tableName, "Error", "Failed to execute primary delete query " & deleteQueryName & ": " & Err.Description
            MsgBox "Failed to purge table " & tableName & ": " & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo ErrHandler
        LogImport tableName, "Success", "Table " & tableName & " purged using " & deleteQueryName
    End If
    
    ' Execute the secondary delete query if defined
    If secondaryDeleteQuery <> "" Then
        On Error Resume Next
        DoCmd.SetWarnings False
        DoCmd.OpenQuery secondaryDeleteQuery
        DoCmd.SetWarnings True
        If Err.Number <> 0 Then
            LogImport tableName, "Error", "Failed to execute secondary delete query " & secondaryDeleteQuery & ": " & Err.Description
            MsgBox "Failed to purge related table for " & tableName & ": " & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo ErrHandler
        LogImport tableName, "Success", "Related table purged using " & secondaryDeleteQuery
    End If
    
    ' Call the import VBA sub dynamically
    Application.Run importSubName, tableName, filePath
    
    ' Call the refresh VBA sub if defined
    If refreshSubName <> "" Then
        Application.Run refreshSubName
    End If
    
    ' Log success
    LogImport tableName, "Success", "Imported data from " & filePath & " and ran refresh."
    MsgBox "Tables purged, data imported, and refreshed successfully for " & tableName, vbInformation
    
    Exit Sub
    
ErrHandler:
    LogImport tableName, "Error", "Error in import/refresh process: " & Err.Description
    MsgBox "Error in import/refresh process: " & Err.Description, vbCritical
End Sub

Private Sub LogImport(tableName As String, status As String, details As String)
    CurrentDb.Execute "INSERT INTO tblImportLog (TableName, Status, Details, ImportDate) " & _
                      "VALUES ('" & tableName & "', '" & status & "', '" & details & "', #" & Now & "#)"
End Sub

