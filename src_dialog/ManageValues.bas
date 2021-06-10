Attribute VB_Name = "ManageValues"
'@Folder("ValidateUserInput.Runner")
Option Explicit

'@ModuleDescription "Main entry in Application"
'Option Private Module

Public Sub Main()

    Dim workSheetName As String
    workSheetName = "TablesValues"

    Dim View As IView
    Set View = TablesValuesView.Create()

    View.MinimumHeight 330
    View.MinimumWidth 540

    If View.ShowDialog(workSheetName) Then
        Debug.Print "Manage Accounts Loaded."
    Else
        Debug.Print "Manage Accounts cancelled."
    End If

End Sub
