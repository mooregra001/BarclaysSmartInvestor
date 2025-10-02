Attribute VB_Name = "modFormCreation"
Option Compare Database
Option Explicit

Sub CreateImportForm()
    'On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctlText As Control
    Dim ctlLabelText As Control
    Dim ctlButtonBrowse As Control
    Dim ctlCombo As Control
    Dim ctlLabelCombo As Control
    Dim ctlButtonImport As Control
    Dim mdl As Module
    
    ' Create a new form
    Set frm = CreateForm
    frm.Caption = "Import Data"
    frm.RecordSource = ""
       
    
    ' Add a text box for file path
    Set ctlText = CreateControl(frm.Name, acTextBox, , , , 1000, 1000, 4000, 300)
    ctlText.Name = "txtFilePath"
    
    ' Add a label for the text box
    Set ctlLabelText = CreateControl(frm.Name, acLabel, , "txtFilePath", , 1000, 600, 4000, 300)
    ctlLabelText.Name = "lblFilePath"
    ctlLabelText.Caption = "Excel File Path:"
    
    ' Add a browse button
    Set ctlButtonBrowse = CreateControl(frm.Name, acCommandButton, , , , 5100, 1000, 1500, 400)
    ctlButtonBrowse.Name = "cmdBrowse"
    ctlButtonBrowse.Caption = "Browse"
    
    ' Add a combo box for table selection (default to tblInvestments)
    Set ctlCombo = CreateControl(frm.Name, acComboBox, , , , 1000, 2000, 4000, 300)
    ctlCombo.Name = "cboTable"
    ctlCombo.RowSource = "SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0"
    ctlCombo.DefaultValue = """tblInvestments"""
    
    ' Add a label for the combo box
    Set ctlLabelCombo = CreateControl(frm.Name, acLabel, , "cboTable", , 1000, 1600, 4000, 300)
    ctlLabelCombo.Name = "lblTable"
    ctlLabelCombo.Caption = "Target Table:"
    
    ' Add an import button
    Set ctlButtonImport = CreateControl(frm.Name, acCommandButton, , , , 5100, 2000, 1500, 400)
    ctlButtonImport.Name = "cmdImport"
    ctlButtonImport.Caption = "Import"
    
    ' Add VBA code for the browse button
    Set mdl = frm.Module
    mdl.InsertLines 1, "Private Sub cmdBrowse_Click()"
    mdl.InsertLines 2, "    On Error GoTo ErrHandler"
    mdl.InsertLines 3, "    Dim fDialog As Object"
    mdl.InsertLines 4, "    Set fDialog = Application.FileDialog(3) ' File picker"
    mdl.InsertLines 5, "    fDialog.Title = ""Select Excel File"""
    mdl.InsertLines 6, "    fDialog.Filters.Clear"
    mdl.InsertLines 7, "    fDialog.Filters.Add ""Excel Files"", ""*.xlsx;*.xls"""
    mdl.InsertLines 8, "    If fDialog.Show = True Then"
    mdl.InsertLines 9, "        Me.txtFilePath.Value = fDialog.SelectedItems(1)"
    mdl.InsertLines 10, "    End If"
    mdl.InsertLines 11, "    Exit Sub"
    mdl.InsertLines 12, "ErrHandler:"
    mdl.InsertLines 13, "    MsgBox ""Error selecting file: "" & Err.Description, vbCritical"
    mdl.InsertLines 14, "End Sub"
    
    ' Add VBA code for the import button
    mdl.InsertLines 15, "Private Sub cmdImport_Click()"
    mdl.InsertLines 16, "    On Error GoTo ErrHandler"
    mdl.InsertLines 17, "    Dim filePath As String"
    mdl.InsertLines 18, "    Dim targetTable As String"
    mdl.InsertLines 19, "    filePath = Me.txtFilePath.Value"
    mdl.InsertLines 20, "    targetTable = Me.cboTable.Value"
    mdl.InsertLines 21, "    If Len(filePath) > 0 And Len(targetTable) > 0 Then"
    mdl.InsertLines 22, "        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, targetTable, filePath, True"
    mdl.InsertLines 23, "        MsgBox ""Data imported successfully to "" & targetTable & ""!"", vbInformation"
    mdl.InsertLines 24, "    Else"
    mdl.InsertLines 25, "        MsgBox ""Please select a file and target table."", vbExclamation"
    mdl.InsertLines 26, "    End If"
    mdl.InsertLines 27, "    Exit Sub"
    mdl.InsertLines 28, "ErrHandler:"
    mdl.InsertLines 29, "    MsgBox ""Error importing data: "" & Err.Description, vbCritical"
    mdl.InsertLines 30, "End Sub"
    
    ' Save and close the form
    DoCmd.Close acForm, frm.Name, acSaveYes
    MsgBox "Form 'Import Excel Data to tblInvestments' created successfully!", vbInformation
    
    Exit Sub
ErrHandler:
    MsgBox "Error creating form: " & Err.Description, vbCritical
End Sub


