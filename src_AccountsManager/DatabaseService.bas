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

    Dim sanitizeworkBookFilePath
    sanitizeworkBookFilePath = ParseResource(workBookFilePath)

    Dim workBookConnectionString As String
    workBookConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sanitizeworkBookFilePath & _
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


Private Function ParseResource(URL As String)
'Uncomment the below line to test locally without calling the function & remove argument above

Dim SplitURL() As String
Dim i As Integer
Dim WebDAVURI As String


'Check for a double forward slash in the resource path. This will indicate a URL
If Not InStr(1, URL, "//", vbBinaryCompare) = 0 Then

    'Split the URL into an array so it can be analyzed & reused
    SplitURL = Split(URL, "/", , vbBinaryCompare)

    'URL has been found so prep the WebDAVURI string
    WebDAVURI = "\\"

    'Check if the URL is secure
    If SplitURL(0) = "https:" Then
        'The code iterates through the array excluding unneeded components of the URL
        For i = 0 To UBound(SplitURL)
            If Not SplitURL(i) = "" Then
                Select Case i
                    Case 0
                        'Do nothing because we do not need the HTTPS element
                    Case 1
                        'Do nothing because this array slot is empty
                    Case 2
                    'This should be the root URL of the site. Add @ssl to the WebDAVURI
                        WebDAVURI = WebDAVURI & SplitURL(i) & "@ssl"
                    Case Else
                        'Append URI components and build string
                        WebDAVURI = WebDAVURI & "\" & SplitURL(i)
                End Select
            End If
        Next i

    Else
    'URL is not secure
        For i = 0 To UBound(SplitURL)

           'The code iterates through the array excluding unneeded components of the URL
            If Not SplitURL(i) = "" Then
                Select Case i
                    Case 0
                        'Do nothing because we do not need the HTTPS element
                    Case 1
                        'Do nothing because this array slot is empty
                        Case 2
                    'This should be the root URL of the site. Does not require an additional slash
                        WebDAVURI = WebDAVURI & SplitURL(i)
                    Case Else
                        'Append URI components and build string
                        WebDAVURI = WebDAVURI & "\" & SplitURL(i)
                End Select
            End If
        Next i
    End If
 'Set the Parse_Resource value to WebDAVURI
 ParseResource = WebDAVURI
Else
'There was no double forward slash so return system path as is
    ParseResource = URL
End If

End Function

