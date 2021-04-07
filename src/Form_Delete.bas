Attribute VB_Name = "Form_Delete"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: Form_Delete ******************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Const passWordLock As String = "fnextxx" 'Excel Application Password
Dim ctrl As MSForms.Control
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Delete /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Delete(UForm As Object, selectedRow As Long)

    Dim DebtorFile As Workbook, WorkBookForm As String, SheetFrom As Worksheet, SheetTo As Worksheet, NxtFreeRow As Long
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    If ThisWorkbook.ReadOnly = True Then Exit Sub
    If selectedRow = 0 Then
        Unload UForm
    Else
        WorkBookForm = ThisWorkbook.Name
        Set DebtorFile = Workbooks(WorkBookForm)
        Set SheetFrom = DebtorFile.Sheets(8)
'        aa_valData.Unprotect Password:=passWordLock ' unprotect "Form" Worksheet
            SheetFrom.Rows(selectedRow).Cut 'Select DataRow to transfer
            Set SheetTo = DebtorFile.Sheets(10) ' Set Deleted Sheet
            NxtFreeRow = SheetTo.Range("A1048576").End(xlUp).Row + 1 ' Search next empty row
            SheetTo.Rows(NxtFreeRow).Insert 'Transfer DataRow to Deleted Sheet
            SheetFrom.Rows(selectedRow).EntireRow.Delete 'Delete empty DataRow
'        aa_valData.Protect Password:=passWordLock, AllowFiltering:=True ' protect "Form" Worksheet
        Unload UForm 'Close the userform
    End If
End Sub

