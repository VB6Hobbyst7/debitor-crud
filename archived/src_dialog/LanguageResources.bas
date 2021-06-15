Attribute VB_Name = "LanguageResources"
'@Folder("ValidateUserInput.Shared.Resources")
Option Explicit
Option Private Module

Public Enum Culture
    EN_US = 1033    'The English US language
    EN_UK = 2057    'The English UK language
    EN_CA = 4105    'The English Canadian language
    DE_DE = 1031    'The German language
    DE_AT = 3079    'The German Austria language
    DE_LI = 5127    'The German Liechtenstein language
    DE_LU = 4103    'The German Luxembourg language
End Enum

Private CaptionSourceSheet As Worksheet
Private RowSourceSheet As Worksheet

Private Sub Initialize(ByVal Source As String, Optional ByVal Switch As Boolean)
    
    Dim languageCode As String
    
    Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
        
        Case Culture.EN_CA, Culture.EN_UK, Culture.EN_US:
            languageCode = "EN"
        
        Case Culture.DE_DE, Culture.DE_AT, Culture.DE_LI, Culture.DE_LU:
            languageCode = "DE"
        
        Case Else:
            languageCode = "EN"
            
    End Select
    
    If Not Switch Then
        Set CaptionSourceSheet = ThisWorkbook.Worksheets(Source & languageCode)
    Else
        Set RowSourceSheet = ThisWorkbook.Worksheets(Source & languageCode)
    End If
    
End Sub

Public Function GetResourceString(ByVal resourceName As String, ByVal ColumnIndex As Integer) As String
    
    Dim TableList As ListObject
    If CaptionSourceSheet Is Nothing Then Initialize "CaptionSource."
    Set TableList = CaptionSourceSheet.ListObjects(1)
    
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = 1 To TableList.ListRows.Count
        Dim lookup As String
        lookup = TableList.Range(i + 1, 1)
        If lookup = resourceName Then
            GetResourceString = TableList.Range(i + 1, ColumnIndex)
            Exit Function
        End If
    Next
    
End Function

Public Function GetRowSourceList(ByVal resourceName As String) As Variant

    If RowSourceSheet Is Nothing Then Initialize "RowSources.", True
    
    Dim HeadersCount As Long
    HeadersCount = RowSourceSheet.UsedRange.Columns.Count
    
    Dim Headers As Range
    Set Headers = RowSourceSheet.Range("A1").Resize(, HeadersCount)
    
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = 1 To HeadersCount
        Dim lookupRange As Range
        Set lookupRange = Headers.Cells(1, i)
        Dim lookup As String
        lookup = lookupRange.Value2
        If lookup = resourceName Then
            Dim resultArray As Range
            Set resultArray = RowSourceSheet.Range(lookupRange.End(xlDown), lookupRange.End(xlUp).Offset(1))
            Dim result As Variant
            result = resultArray.value
            If IsArray(result) Then
                GetRowSourceList = result
            Else
                GetRowSourceList = Array(resultArray.Item(1).Value2)
            End If
            Exit Function
        End If
    Next

End Function
