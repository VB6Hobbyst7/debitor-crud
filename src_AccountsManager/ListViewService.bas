Attribute VB_Name = "ListViewService"
'@Folder("AccountsManager.Infrastructure.Services")
Option Explicit
Option Private Module

#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Private Const LVM_FIRST = &H1000

Public Sub InitializeListView(ByRef ListViewObject As MSComctlLib.ListView, ByRef workBookFilePath As String, ByRef workSheetName As String, _
Optional ByRef FieldNames As String = "*", _
Optional ByRef joinClause As String = vbNullString, _
Optional ByRef predicateExpression As String = vbNullString, _
Optional ByRef OrderByExpression As String = vbNullString)
    
    If ListViewObject Is Nothing Then Exit Sub

    Dim DataRecordset As ADODB.Recordset
    Set DataRecordset = QuickExecuteWorkSheetQuery(workBookFilePath, workSheetName, FieldNames, joinClause, predicateExpression, OrderByExpression)
    Dim ColumnNames As ADODB.Fields
    Set ColumnNames = DataRecordset.Fields

    With ListViewObject
        .ColumnHeaders.Clear
        .ListItems.Clear
        Dim header As Variant
        For Each header In ColumnNames
            .ColumnHeaders.add text:=header.Name
        Next
        
        With DataRecordset
            Dim Count As Long
            Do Until .EOF
                Count = 1 'Initalize a count, this will help to determine whether to add a new row vice a new column
                Dim fieldItem As Variant
                For Each fieldItem In .Fields 'Loop through all the fields in the recordset
                    Dim viewListItem As MSComctlLib.ListItem
                    If Count = 1 Then 'If it's the first field of the recordset, that means we have the first column of a new row
                        'If it's a new row, then we will add a new ListItems (ROW) object
                        Set viewListItem = ListViewObject.ListItems.add(text:=IIf(IsNull(fieldItem.Value), vbNullString, fieldItem.Value))
                    Else
                        'If it's not a new row, then add a ListSubItem (ELEMENT) instead
                        viewListItem.ListSubItems.add text:=IIf(IsNull(fieldItem.Value), vbNullString, fieldItem.Value)
                    End If
                    Count = Count + 1 'Make sure to increment the count, or else EVERYONE will be a "New Row"
                Next
                .MoveNext
            Loop
        End With
        .LabelEdit = lvwManual

    End With
    ListViewAutoSizeColumn ListViewObject
End Sub

Public Sub ListViewColumnSort(ByRef ListViewObject As MSComctlLib.ListView, ByRef ColumnHeader As MSComctlLib.ColumnHeader)
    With ListViewObject
        If .SortKey <> ColumnHeader.Index - 1 Then
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        Else
            If .SortOrder = lvwAscending Then
             .SortOrder = lvwDescending
            Else
             .SortOrder = lvwAscending
            End If
        End If
        .Sorted = True
        If ColumnHeader.Index = 1 Then
            SortDataWithNumbers ListViewObject
        Else
            .Sorted = True
        End If
    End With
End Sub

Public Sub SortDataWithNumbers(ByRef ListViewObject As MSComctlLib.ListView)
    Dim stringTemp As String * 10
    Dim lvCount As Long
    '@Ignore UseMeaningfulName
    Dim i As Long
    
    With ListViewObject
        lvCount = .ListItems.Count
        
        For i = 1 To lvCount
            stringTemp = vbNullString
            
            If .SortKey Then
                'RSet - right align a string within a string variable.
                RSet stringTemp = .ListItems(i).SubItems(.SortKey)
                .ListItems(i).SubItems(.SortKey) = stringTemp
            Else
                RSet stringTemp = .ListItems(i)
                .ListItems(i).text = stringTemp
            End If
        Next
        
        .Sorted = True
        
        For i = 1 To lvCount
            If .SortKey Then
                .ListItems(i).SubItems(.SortKey) = _
                LTrim$(.ListItems(i).SubItems(.SortKey))
            Else
                .ListItems(i).text = LTrim$(.ListItems(i))
            End If
        Next
    End With
End Sub

Public Sub ListViewAutoSizeColumn(ByRef ListViewObject As ListView, Optional ByRef column As ColumnHeader = Nothing)

    If column Is Nothing Then
        Dim columnItem As ColumnHeader
        For Each columnItem In ListViewObject.ColumnHeaders
            '@Ignore ValueRequired
            SendMessage ListViewObject.hWnd, LVM_FIRST + 30, columnItem.Index - 1, ByVal -2
        Next
    Else
        '@Ignore ValueRequired
        SendMessage ListViewObject.hWnd, LVM_FIRST + 30, column.Index - 1, ByVal -2
    End If
    ListViewObject.Refresh

End Sub


