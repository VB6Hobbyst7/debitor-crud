VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "This class is essentially a Collection<T>, where all items are of type T... with lots of added functionality"
'@Folder AccountsManager.Common.List
'@ModuleDescription "This class is essentially a Collection<T>, where all items are of type T... with lots of added functionality"
'@PredeclaredId
'@Exposed
Private Type tList
    Encapsulated As Collection
    ItemTypeName As String
End Type

Private this As tList
Option Explicit

Public Function Create(ParamArray Values() As Variant) As list
                
    Dim valuesArray() As Variant
    valuesArray = Values

    Dim result As New list
    result.AddArray valuesArray
    
    Set Create = result

End Function

'' ***************************************************************************************
'' Private functions
'' ***************************************************************************************

Private Function IsReferenceType() As Boolean
    If this.ItemTypeName = vbNullString Then Exit Function
    IsReferenceType = IsObject(this.Encapsulated(1))
End Function

Private Function IsComparable() As Boolean
    If IsReferenceType Then
        IsComparable = TypeOf First Is IComparable
    End If
End Function

Private Function CompareReferenceTypes(Value As Variant, other As Variant) As Integer

    Dim comparable As IComparable

    If IsComparable Then

        Set comparable = Value
        CompareReferenceTypes = comparable.CompareTo(other)

    Else

        RaiseErrorMustImplementIComparable "CompareReferenceTypes()"

    End If

End Function

Private Function CompareValueTypes(Value As Variant, other As Variant) As Integer

    If Value < other Then

        CompareValueTypes = -1

    ElseIf Value > other Then

        CompareValueTypes = 1

    End If

End Function

Private Function IsEquatable() As Boolean
    If IsReferenceType Then
        IsEquatable = TypeOf First Is IEquatable
    End If
End Function

Private Function EquateReferenceTypes(Value As Variant, other As Variant) As Boolean

    Dim equatable As IEquatable
    If IsEquatable Then

        Set equatable = Value
        EquateReferenceTypes = equatable.Equals(other)

    Else
        Logging.logWARN "WARNING: Reference type doesn't implement IEquatable, using reference equality."
'        Debug.Print "WARNING: Reference type doesn't implement IEquatable, using reference equality."
        EquateReferenceTypes = (ObjPtr(Value) = ObjPtr(other))

    End If

End Function

Private Function EquateValueTypes(Value As Variant, other As Variant) As Boolean

    EquateValueTypes = (Value = other)

End Function

Private Function ValidateItemType(Value As Variant) As Variant

    If this.ItemTypeName = vbNullString Then this.ItemTypeName = TypeName(Value)
    ValidateItemType = IsTypeSafe(Value)

End Function

Private Sub RaiseErrorUnsafeType(member As String, suppliedType As String)
    Err.Raise 13, StringType.Format("{0}.{1}", ToString, member), _
                  StringType.Format("Type Mismatch. Expected: '{0}', '{1}' was supplied.", this.ItemTypeName, suppliedType)
End Sub

Private Sub RaiseErrorMustImplementIComparable(member As String)
    Err.Raise 5, StringType.Format("{0}.{1}", ToString, member), "Invalid operation: method requires a list of numeric, date or string values, or a list of objects implementing the IComparable interface."
End Sub

Private Sub Class_Initialize()
    Set this.Encapsulated = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.Encapsulated = Nothing
End Sub

'' ***************************************************************************************
'' Properties
'' ***************************************************************************************

'@Description "Gets or sets the element at the specified index."
'@DefaultMember
Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_Description = "Gets or sets the element at the specified index."
Attribute Item.VB_UserMemId = 0

    If IsReferenceType Then
        Set Item = this.Encapsulated(Index)
    Else
        Item = this.Encapsulated(Index)
    End If

End Property

Public Property Let Item(ByVal Index As Long, ByVal Value As Variant)
'Sets the element at the specified index.

    If Not IsTypeSafe(Value) Then RaiseErrorUnsafeType "Item(Let)", TypeName(Value)

    RemoveAt Index
    If Index = Count Then
        add Value
    Else
        Insert Index, Value
    End If

End Property

Public Property Set Item(ByVal Index As Long, ByVal Value As Variant)
'Sets the element at the specified index.

    If Not IsTypeSafe(Value) Then RaiseErrorUnsafeType "Item(Set)", TypeName(Value)

    RemoveAt Index
    If Index = Count Then
        add Value
    Else
        Insert Index, Value
    End If

End Property

'@Enumerator
'@MemberAttribute VB_MemberFlags, "40"
'@Description "Gets an enumerator that iterates through the List."
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Gets an enumerator that iterates through the List."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = this.Encapsulated.[_NewEnum]

End Property

'@Description "Gets the number of elements contained in the List."
Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the number of elements contained in the List."

    Count = this.Encapsulated.Count

End Property

'' ***************************************************************************************
'' Public functions
'' ***************************************************************************************

'@Description "Adds the specified element(s) to the end of the List."
Public Sub add(ParamArray Values())
Attribute add.VB_Description = "Adds the specified element(s) to the end of the List."

    Dim valuesArray() As Variant
    valuesArray = Values

    AddArray valuesArray

End Sub

'@Description "Adds the specified elements to the end of the List."
Public Sub AddRange(Values As list)
Attribute AddRange.VB_Description = "Adds the specified elements to the end of the List."

    AddArray Values.ToArray

End Sub

'@Description "Adds the specified elements to the end of the List."
Public Sub AddArray(Values() As Variant)
Attribute AddArray.VB_Description = "Adds the specified elements to the end of the List."

    '@Ignore UseMeaningfulName, MultipleDeclarations
    Dim Value As Variant, i As Long
    For i = LBound(Values) To UBound(Values)
        If ValidateItemType(Value) Then
            this.Encapsulated.add Values(i)
        Else
            RaiseErrorUnsafeType "AddArray()", TypeName(Value)
        End If
    Next

End Sub

'@Description "Removes all elements from the List."
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all elements from the List."

    Do Until Count = 0
        this.Encapsulated.Remove 1
    Loop

End Sub

'@Description "Determines whether an element is in the List."
Public Function Contains(Value As Variant) As Boolean
Attribute Contains.VB_Description = "Determines whether an element is in the List."
'Determines whether an element is in the List.

    Contains = (IndexOf(Value) <> -1)

End Function

'@Description "Returns the first element of the List."
Public Function First() As Variant
Attribute First.VB_Description = "Returns the first element of the List."

    If Count = 0 Then Exit Function
    If IsReferenceType Then
        Set First = Item(1)
    Else
        First = Item(1)
    End If

End Function

'@Description "Creates a copy of a range of elements in the source List."
Public Function GetRange(ByVal Index As Long, ByVal valuesCount As Long) As list
Attribute GetRange.VB_Description = "Creates a copy of a range of elements in the source List."

    Dim result As list
    If Index > Count Then Err.Raise 9 'index out of range

    Dim lastIndex As Long
    lastIndex = IIf(Index + valuesCount > Count, Count, Index + valuesCount)

    Set result = New list

    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = Index To lastIndex
        result.add Item(i)
    Next

    Set GetRange = result

End Function

'@Description "Searches for the specified object and returns the 1-based index of the first occurrence within the entire List."
Public Function IndexOf(Value As Variant) As Long
Attribute IndexOf.VB_Description = "Searches for the specified object and returns the 1-based index of the first occurrence within the entire List."

    Dim found As Boolean
    Dim isRef As Boolean
    isRef = IsReferenceType

    '@Ignore UseMeaningfulName
    Dim i As Long

    If Count = 0 Then IndexOf = -1: Exit Function
    For i = 1 To Count

        If isRef Then

            found = EquateReferenceTypes(Value, Item(i))

        Else

            found = EquateValueTypes(Value, Item(i))

        End If

        If found Then IndexOf = i: Exit Function

    Next

    IndexOf = -1

End Function

'@Description "Inserts the specified elements into the List at the specified index."
Public Sub InsertArray(ByVal Index As Long, Values() As Variant)
Attribute InsertArray.VB_Description = "Inserts the specified elements into the List at the specified index."

    Dim temp As list
    Set temp = GetRange(Index, Count)

    RemoveRange Index, Count

    AddArray Values
    AddRange temp

End Sub

'@Description "Inserts the specified elements into the List at the specified index."
Public Sub InsertRange(ByVal Index As Long, Values As list)
Attribute InsertRange.VB_Description = "Inserts the specified elements into the List at the specified index."

    Dim temp As list
    Set temp = GetRange(Index, Count)

    RemoveRange Index, Count

    AddRange Values
    AddRange temp

End Sub

'@Description "Inserts the specified element(s) into the List at the specified index."
Public Sub Insert(ByVal Index As Long, ParamArray Values())
Attribute Insert.VB_Description = "Inserts the specified element(s) into the List at the specified index."

    Dim valuesArray() As Variant
    valuesArray = Values

    InsertArray Index, valuesArray

End Sub

'@Description "Determines whether the List can be sorted."
Public Function IsSortable() As Boolean
Attribute IsSortable.VB_Description = "Determines whether the List can be sorted."

    If Count = 0 Then Exit Function

    Dim firstItem As Variant
    If IsReferenceType Then
        Set firstItem = First
    Else
        firstItem = First
    End If

    IsSortable = IsNumeric(firstItem) _
                Or IsDate(firstItem) _
                Or this.ItemTypeName = "String" _
                Or IsComparable

End Function

'@Description "Determines whether a value can be safely added to the List."
Public Function IsTypeSafe(Value As Variant) As Boolean
Attribute IsTypeSafe.VB_Description = "Determines whether a value can be safely added to the List."

'Returns true if the type of specified value matches the type of items already in the list,
'or it the type of specified value is a numeric type smaller than the type of items already in the list.
'This means a List<Long> can contain Integer values, but a List<Integer> cannot contain Long values.

    Dim result As Boolean

    'most common cases: this.ItemTypeName isn't yet defined, or matches TypeName(value):
    result = this.ItemTypeName = vbNullString Or this.ItemTypeName = TypeName(Value)
    If result Then IsTypeSafe = result: Exit Function

    'all other cases demand more processing:
    IsTypeSafe = result _
        Or this.ItemTypeName = "Integer" And StringType.MatchesAny(TypeName(Value), "Byte") _
        Or this.ItemTypeName = "Long" And StringType.MatchesAny(TypeName(Value), "Integer", "Byte") _
        Or this.ItemTypeName = "Single" And StringType.MatchesAny(TypeName(Value), "Long", "Integer", "Byte") _
        Or this.ItemTypeName = "Double" And StringType.MatchesAny(TypeName(Value), "Long", "Integer", "Byte", "Single") _
        Or this.ItemTypeName = "Currency" And StringType.MatchesAny(TypeName(Value), "Long", "Integer", "Byte", "Single", "Double")

End Function

'@Description "Returns the last element of the List."
Public Function Last() As Variant
Attribute Last.VB_Description = "Returns the last element of the List."

    If Count = 0 Then Exit Function
    If IsReferenceType Then
        Set Last = Item(Count)
    Else
        Last = Item(Count)
    End If

End Function

'@Description "Searches for the specified object and returns the 1-based index of the last occurrence within the entire List."
Public Function LastIndexOf(Value As Variant) As Long
Attribute LastIndexOf.VB_Description = "Searches for the specified object and returns the 1-based index of the last occurrence within the entire List."

    Dim found As Boolean
    Dim isRef As Boolean
    isRef = IsReferenceType

    LastIndexOf = -1
    If Count = 0 Then Exit Function

    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = 1 To Count

        If isRef Then

            found = EquateReferenceTypes(Value, Item(i))

        Else

            found = EquateValueTypes(Value, Item(i))

        End If

        If found Then LastIndexOf = i

    Next

End Function

'@Description "Returns the maximum value in the List."
Public Function Max() As Variant
Attribute Max.VB_Description = "Returns the maximum value in the List."

    Dim isRef As Boolean
    isRef = IsReferenceType

    Dim largest As Variant
    Dim isLarger As Boolean

    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = 1 To Count

        If isRef Then

            If IsEmpty(largest) Then Set largest = Item(i)
            isLarger = CompareReferenceTypes(Item(i), largest) > 0

            If isLarger Or IsEmpty(Max) Then
                Set largest = Item(i)
                Set Max = largest
            End If

        Else

            If IsEmpty(largest) Then largest = Item(i)
            isLarger = CompareValueTypes(Item(i), largest) > 0

            If isLarger Or IsEmpty(Max) Then
                largest = Item(i)
                Max = largest
            End If

        End If


    Next

End Function

'@Description "Returns the minimum value in the List."
Public Function Min() As Variant
Attribute Min.VB_Description = "Returns the minimum value in the List."

    Dim isRef As Boolean
    isRef = IsReferenceType

    Dim smallest As Variant
    Dim isSmaller As Boolean

    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = 1 To Count

        If isRef Then

            If IsEmpty(smallest) Then Set smallest = Item(i)
            isSmaller = CompareReferenceTypes(Item(i), smallest) < 0

            If isSmaller Or IsEmpty(Min) Then
                Set smallest = Item(i)
                Set Min = smallest
            End If

        Else

            If IsEmpty(smallest) Then smallest = Item(i)
            isSmaller = CompareValueTypes(Item(i), smallest) < 0

            If isSmaller Or IsEmpty(Min) Then
                smallest = Item(i)
                Min = smallest
            End If

        End If


    Next

End Function

'@Description "Reverses the order of the elements in the entire List."
Public Sub Reverse()
Attribute Reverse.VB_Description = "Reverses the order of the elements in the entire List."

    Dim temp As New list
    Do Until Count = 0

        temp.add Item(Count)
        RemoveAt Count

    Loop

    AddRange temp

End Sub

'@Description "Removes the first occurrence of specified object(s) from the List."
Public Sub Remove(ParamArray Values())
Attribute Remove.VB_Description = "Removes the first occurrence of specified object(s) from the List."

    '@Ignore UseMeaningfulName
    Dim i As Long
    Dim Index As Long

    For i = LBound(Values) To UBound(Values)

        Index = IndexOf(Values(i))
        If Index <> -1 Then RemoveAt Index

    Next

End Sub

'@Description "Removes the element at the specified index of the List."
Public Sub RemoveAt(ByVal Index As Long)
Attribute RemoveAt.VB_Description = "Removes the element at the specified index of the List."
'Removes the element at the specified index of the List.

    this.Encapsulated.Remove Index

End Sub

'@Description "Removes a range of elements from the List."
Public Sub RemoveRange(ByVal Index As Long, ByVal valuesCount As Long)
Attribute RemoveRange.VB_Description = "Removes a range of elements from the List."

    '@Ignore VariableNotUsed, UseMeaningfulName
    Dim i As Long
    For i = Index To Index + valuesCount - 1

        RemoveAt Index

    Next

End Sub

'@Description "Sorts the elements in the entire List."
Public Sub Sort()
Attribute Sort.VB_Description = "Sorts the elements in the entire List."

    Dim temp As list
    Dim minValue As Variant

    If Not IsSortable Then RaiseErrorMustImplementIComparable "Sort()"

    Dim isRef As Boolean
    isRef = IsReferenceType

    Set temp = New list
    Do Until Count = 0

        If isRef Then

            Set minValue = Min

        Else

            minValue = Min

        End If

        temp.add minValue
        Remove minValue

    Loop

    AddRange temp

End Sub

'@Description "Sorts the elements in the entire List, in descending order."
Public Sub SortDescending()
Attribute SortDescending.VB_Description = "Sorts the elements in the entire List, in descending order."
    
    Dim temp As list
    Dim maxValue As Variant

    If Not IsSortable Then RaiseErrorMustImplementIComparable "SortDescending()"

    Dim isRef As Boolean
    isRef = IsReferenceType

    Set temp = New list
    Do Until Count = 0

        If isRef Then
            Set maxValue = Max
        Else
            maxValue = Max
        End If

        temp.add maxValue
        Remove maxValue

    Loop

    AddRange temp

End Sub

'@Description "Copies the elements of the List to a new array."
Public Function ToArray() As Variant()
Attribute ToArray.VB_Description = "Copies the elements of the List to a new array."

    Dim result() As Variant
    ReDim result(1 To Count)

    '@Ignore UseMeaningfulName
    Dim i As Long
    If Count = 0 Then Exit Function

    If IsReferenceType Then
        For i = 1 To Count
            Set result(i) = Item(i)
        Next
    Else
        For i = 1 To Count
            result(i) = Item(i)
        Next
    End If

    ToArray = result

End Function

'@Description "Returns a string that represents the current List object."
Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a string that represents the current List object."

    ToString = StringType.Format("{0}<{1}>", TypeName(Me), StringType.Coalesce(this.ItemTypeName, "Variant"))

End Function
