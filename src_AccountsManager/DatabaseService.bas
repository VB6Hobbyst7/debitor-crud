Attribute VB_Name = "DatabaseService"
Attribute VB_Description = "Common utility methods that extend the functionality of the API"
'@Folder("AccountsManager.Infrastructure.Services")
'@ModuleDescription("Common utility methods that extend the functionality of the API")
Option Explicit
Option Private Module

Public Function QuickExecuteWorkSheetQuery(ByVal workBookFilePath As String, ByVal workSheetName As String, _
Optional FieldNames As String = "*", _
Optional ByVal joinClause As String = vbNullString, _
Optional ByVal predicateExpression As String = vbNullString, _
Optional ByVal OrderByExpression As String = vbNullString) As ADODB.Recordset

    Dim workBookConnectionString As String
    workBookConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & workBookFilePath & _
                               ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1'"
                               
    predicateExpression = IIf((predicateExpression <> vbNullString And InStr(UCase$(predicateExpression), "WHERE") = 0), "WHERE " & predicateExpression, predicateExpression)
    
    OrderByExpression = IIf((OrderByExpression <> vbNullString And InStr(UCase$(OrderByExpression), "ORDER BY") = 0), "ORDER BY " & OrderByExpression, OrderByExpression)
        
    Dim queryString As String
    queryString = "SELECT " & SanitizeDelimitedFieldNames(FieldNames) & vbNewLine & _
                  "FROM [" & workSheetName & "$] " & vbNewLine & _
                  joinClause & vbNewLine & _
                  predicateExpression & vbNewLine & _
                  OrderByExpression
    
    Dim dataBase As ADODB.Connection
    Set dataBase = New ADODB.Connection
    Set dataBase = CreateAdoConnection(workBookConnectionString, adUseClient)
    
    Dim adoCommand As ADODB.command
    Set adoCommand = New ADODB.command
    With adoCommand
    Set .ActiveConnection = dataBase
        .CommandText = queryString
        .commandType = adCmdText
    End With
        
    Dim results As ADODB.Recordset
    Set results = New ADODB.Recordset
    results.CursorType = adOpenKeyset
    results.LockType = adLockOptimistic
    
    Set results = adoCommand.Execute()
    
    Set QuickExecuteWorkSheetQuery = results
End Function

Private Function CreateAdoConnection(ByVal connString As String, ByVal cursorLocationValue As ADODB.CursorLocationEnum) As ADODB.Connection
    Dim result As ADODB.Connection
    Set result = New ADODB.Connection

    result.CursorLocation = cursorLocationValue  'must set before opening
    result.Open connString

    Set CreateAdoConnection = result
End Function

Private Function SanitizeDelimitedFieldNames(ByVal delimitedFieldNames As String) As String
    SanitizeDelimitedFieldNames = Replace(Replace(Replace(delimitedFieldNames, ",[]", vbNullString), ", []", vbNullString), "[]", vbNullString)
End Function


