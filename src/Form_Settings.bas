Attribute VB_Name = "Form_Settings"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: Form_Settings ****************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Const passWordLock As String = "fnextxx" 'Excel Application Password
Public DebtorFile As Workbook, WorkBookForm As String, autoClose As Boolean 'WorkBook Object variables
Public Const xRunWhat = "CloseMeX"  ' the name of the procedure to run
Private xTimerDown As Boolean
Private RunWhenX As Date
Dim UForm As Object
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'WorksheetsSetup ////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub WorksheetsSetup(UserAcces As Boolean)

    Application.ScreenUpdating = False
    WorkBookForm = ThisWorkbook.Name
    Set DebtorFile = Workbooks(WorkBookForm)
    DebtorFile.Unprotect passWordLock
        Select Case UserAcces
            Case Is = True
                With DebtorFile.Worksheets 'Visible
                    .Item(9).Visible = -1
                    '.Item(10).Visible = -1
                End With
                On Error Resume Next
                aa_valData.Activate
                Application.GoTo reference:=aa_valData.Range("N1").End(xlDown).End(xlToLeft).End(xlToLeft).Offset(-5), Scroll:=True
            Case Else
                With DebtorFile.Worksheets 'VeryHidden
                    .Item(1).Visible = 0
                    .Item(2).Visible = 0
                    .Item(3).Visible = 0
                    .Item(4).Visible = 0
                    .Item(5).Visible = 0
                    .Item(6).Visible = 0
                    .Item(9).Visible = 0
                    .Item(10).Visible = 0
                End With
                On Error Resume Next
                ZZ_INFO.Activate
                Application.GoTo reference:=ZZ_INFO.Range("A2"), Scroll:=True
        End Select
        Call SyncSheetsToA1
    DebtorFile.Protect passWordLock, True, True
    On Error GoTo 0
    Select Case Application.International(xlCountryCode)
        'Opening message to warn users what will happen if the workbook remain inactive
        Case Is = 49 'Messages displayed in German
            MsgBox "Bitte beachten Sie, dass diese Arbeitsmappe nach 30 Minuten Inaktivität geschlossen wird! " & vbCr & _
                    "Alle Änderungen, die Sie geschpeichert haben werden geschpeichert.", vbInformation, "Warning"
        Case Else   'Messages displayed in English
            MsgBox "Please note that this workbook will close after 30 minutes of inactivity." & vbCr & _
                    "Saved customer data will be saved.", vbInformation, "Warning"
    End Select
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'SetTimer ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub SetTimer()

    RunWhenX = Now + TimeValue("00:32:00")
    Application.OnTime RunWhenX, xRunWhat
    xTimerDown = True
    Debug.Print RunWhenX & " " & xTimerDown & " SetTimer"
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'StopTimer //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub StopTimer()

    On Error Resume Next
    Application.OnTime RunWhenX, xRunWhat, False, False
    xTimerDown = False
    Debug.Print RunWhenX & " " & xTimerDown & " StopTimer"
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'SyncSheetsToA1 /////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub SyncSheetsToA1()

    Dim UserSheet As Worksheet, sht As Worksheet, TopRow As Long, LeftCol As Integer, UserSel As String
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Application.ScreenUpdating = False
    Set UserSheet = ActiveSheet 'Remember the current sheet
    TopRow = 1
    LeftCol = 1
    UserSel = Range("A1").Address
    Application.EnableEvents = False
        For Each sht In ActiveWorkbook.Worksheets 'Loop through the worksheets
            'ActiveWindow.Zoom = 80
            If sht.Visible And Not sht.Name Like "Form" Then 'skip hidden sheets and Form Sheet
                sht.Activate
                Range(UserSel).Select
                ActiveWindow.ScrollRow = TopRow
                ActiveWindow.ScrollColumn = LeftCol
            End If
        Next sht
    Application.EnableEvents = True
    UserSheet.Activate 'Restore the original position
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CloseMeX ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub CloseMeX()

    Dim DebtorFile As Workbook, WorkBookForm As String
    WorkBookForm = ThisWorkbook.Name
    Set DebtorFile = Workbooks(WorkBookForm)
    Application.DisplayAlerts = False 'turn off any warning messages
    If IsLoaded("frm_Debitor") Then Unload UForm 'Closes the UserForm
    DebtorFile.Save 'Save
    DebtorFile.Close
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'FUNCTION - VBATrusted
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function VBATrusted() As Boolean

    On Error Resume Next
    VBATrusted = (Application.VBE.VBProjects.Count) > 0
End Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'FUNCTION - IsLoaded
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function IsLoaded(formName As String) As Boolean

    For Each UForm In VBA.UserForms
        If UForm.Name = formName Then IsLoaded = True: Exit Function
    Next UForm
    IsLoaded = False
End Function
