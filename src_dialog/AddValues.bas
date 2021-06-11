Attribute VB_Name = "AddValues"
'@Folder("ValidateUserInput.Runner")
Option Explicit
Option Private Module

Public Sub Add()

    Dim workSheetName As String
    workSheetName = "Table1Values"

    Dim View As IView
    Set View = Table1View.Create()

    View.MinimumHeight 346
    View.MinimumWidth 318

    If View.ShowDialog(workSheetName) Then
        Debug.Print "Add Values Loaded."
    Else
        Debug.Print "Add Values cancelled."
    End If

End Sub
