Attribute VB_Name = "LanguageResources"
'@Folder AccountsManager.Common.Resources
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

Private captionSourceSheet As Worksheet
Private rowSourceSheet As Worksheet

Private Sub Initialize(ByVal Source As String, Optional ByVal switch As Boolean)
    GuardClauses.GuardEmptyString Source

    Dim languageCode As String
    
    Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
        
        Case Culture.EN_CA, Culture.EN_UK, Culture.EN_US:
            languageCode = "EN"
        
        Case Culture.DE_DE, Culture.DE_AT, Culture.DE_LI, Culture.DE_LU:
            languageCode = "DE"
        
        Case Else:
            languageCode = "EN"
            
    End Select
    
    If Not switch Then
        Set captionSourceSheet = ThisWorkbook.Worksheets(Source & languageCode)
    Else
        Set rowSourceSheet = ThisWorkbook.Worksheets(Source & languageCode)
    End If
    
End Sub

Public Function GetResourceString(ByVal resourceName As String, ByVal columnIndex As Integer, ByVal Source As String) As String
    GuardClauses.GuardEmptyString resourceName
    GuardClauses.GuardEmptyString Source
    
    Dim TableList As ListObject
    If captionSourceSheet Is Nothing Then Initialize Source & "."
    
    Set TableList = captionSourceSheet.ListObjects(1)
    
    Dim i As Long
    For i = 1 To TableList.ListRows.Count
        Dim lookup As String
        lookup = TableList.Range(i + 1, 1)
        If lookup = resourceName Then
            GetResourceString = TableList.Range(i + 1, columnIndex)
            Exit Function
        End If
    Next

End Function

Public Function GetRowSourceList(ByVal resourceName As String, ByVal resourceDescription As String, ByVal Source As String) As Variant
    GuardClauses.GuardEmptyString resourceName
    GuardClauses.GuardEmptyString resourceDescription
    GuardClauses.GuardEmptyString Source

    If rowSourceSheet Is Nothing Then Initialize Source & ".", True
    
    Dim HeadersCount As Long
    HeadersCount = rowSourceSheet.UsedRange.Columns.Count
    
    Dim Headers As Range
    Set Headers = rowSourceSheet.Range("A1").Resize(, HeadersCount)
    
    Dim i As Long
    For i = 1 To HeadersCount
    
        Dim lookupRange As Range
        Set lookupRange = Headers.Cells(1, i)
        
        Dim lookup As String
        lookup = lookupRange.Value2
        If lookup = resourceName Then
            Dim ArrayName As Range
            Set ArrayName = rowSourceSheet.Range(lookupRange.End(xlDown), lookupRange.End(xlUp).Offset(1))
            
            Dim resultName As Variant
            resultName = ArrayName.value
        End If
        
        If lookup = resourceDescription Then
            Dim ArrayDescription As Range
            Set ArrayDescription = rowSourceSheet.Range(lookupRange.End(xlDown), lookupRange.End(xlUp).Offset(1))
            
            Dim resultDescription As Variant
            resultDescription = ArrayDescription.value
        End If
        
        If (Not ArrayName Is Nothing) And (Not ArrayDescription Is Nothing) Then
            Exit For
        End If
    Next
    
    Dim result As Variant
    If IsArray(resultName) And IsArray(resultDescription) Then
        ReDim result(1 To 2, 1 To UBound(resultName))
        For i = 1 To UBound(result, 2)
            result(1, i) = resultName(i, 1)
            result(2, i) = resultDescription(i, 1)
        Next
        GetRowSourceList = Application.WorksheetFunction.Transpose(result)

    Else
        Dim nameList As Variant
        nameList = Array(ArrayName.Item(1).Value2)
        Dim descriptionList As Variant
        descriptionList = Array(ArrayDescription.Item(1).Value2)
        ReDim result(0 To 0, LBound(nameList) To UBound(nameList) + 1)
        For i = 0 To UBound(nameList)
            result(0, i) = nameList(0)
            result(0, i + 1) = descriptionList(0)
        Next i
        GetRowSourceList = result
        
    End If

End Function

