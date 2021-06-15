Attribute VB_Name = "Start"
Attribute VB_Description = "Main entry in Application"
'@Folder("ValidateUserInput.Main")
Option Explicit

'@ModuleDescription "Main entry in Application"
Option Private Module

Public Sub AppStart()

    Dim ViewModel As ValuesViewModel: Set ViewModel = ValuesViewModel.Create()
    Dim Source As String: Source = "TablesValues"
    
    ViewModel.DataSourceTable = Source
    ViewModel.Filter = "notSet"
    ViewModel.ModelHeight = 330
    ViewModel.ModelWidth = 540
    ViewModel.Id = -1

    Dim View As IView
    Set View = TablesView.Create(ViewModel)

    If View.ShowDialog() Then
        Debug.Print "Manage Values Loaded."
    Else
        Debug.Print "Manage Values cancelled."
    End If

End Sub
