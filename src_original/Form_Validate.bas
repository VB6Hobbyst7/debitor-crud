Attribute VB_Name = "Form_Validate"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: Form_Validate ****************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'FUNCTION - ValidationFailed ////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function ValidationFailed(PPage As Object, xPage As Integer, Optional xLang As Long, Optional pName As String) As Boolean

    Dim ctrl As MSForms.Control
'    Application.Run "Form_Settings.StopTimer" ' Stop Timer
'    Application.Run "Form_Settings.SetTimer" ' Start Timer
    For Each ctrl In PPage(xPage).Controls 'Loop through Parent Object
        'MusFill Fields are not empty, check
        If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Or TypeOf ctrl Is MSForms.CheckBox Then
            If ctrl.Object.BackColor = RGB(255, 255, 204) And ctrl.Visible = True Then 'BackColor is yellow?
                Application.Run "Form_Messages.ShowMessage", xLang, "MustFill", ctrl, PPage, xPage, pName 'Message for User
                ctrl.SetFocus
                ValidationFailed = True
                Exit Function 'Terminate Check
            End If
        End If
        If TypeOf ctrl Is MSForms.TextBox Then 'Entered Value type must be Valid(Number, Telefon, Fax or Email), check
            If ctrl.Name Like "*SAPNr" And ctrl.value <> vbNullString Then 'Entered value must be an SAP Number, check
                If Not IsNumeric(ctrl.value) Or Len(ctrl.Text) < 6 Then 'Entered value must be a Number and min length of 6
                    With ctrl
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ForeColor = RGB(255, 0, 0) 'Red
                    End With
                    Application.Run "Form_Messages.ShowMessage", xLang, "OnlyNumbers", ctrl, PPage, xPage, pName 'Message for User
                    ctrl.SetFocus
                    ValidationFailed = True
                    Exit Function 'Terminate Check
                End If
            End If
            If ctrl.Name Like "*Telefon*" And ctrl.value <> vbNullString Then 'Entered value must be a Telefon Number, check
                If Not IsNumeric(Replace(ctrl.Text, " ", "")) Then
                    With ctrl
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ForeColor = RGB(255, 0, 0) 'Red
                    End With
                    Application.Run "Form_Messages.ShowMessage", xLang, "TelFax", ctrl, PPage, xPage, pName 'Message for User
                    ctrl.SetFocus
                    ValidationFailed = True
                    Exit Function 'Terminate Check
                End If
            End If
            If ctrl.Name Like "*Fax*" And ctrl.value <> vbNullString Then 'Entered value must be a Fax Number, check
                If Not IsNumeric(Replace(ctrl.Text, " ", "")) Then
                    With ctrl
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ForeColor = RGB(255, 0, 0) 'Red
                    End With
                    Application.Run "Form_Messages.ShowMessage", xLang, "TelFax", ctrl, PPage, xPage, pName 'Message for User
                    ctrl.SetFocus
                    ValidationFailed = True
                    Exit Function 'Terminate Check
                End If
            End If
            If ctrl.Name Like "*Email*" And ctrl.value <> vbNullString Then 'Entered value must be a valid E-Mail, check
                If Not IsValidEmail(ctrl.Text) Then
                    With ctrl
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ForeColor = RGB(255, 0, 0) 'Red
                    End With
                    Application.Run "Form_Messages.ShowMessage", xLang, "Email", ctrl, PPage, xPage, pName 'Message for User
                    ctrl.SetFocus
                    ValidationFailed = True
                    Exit Function 'Terminate Check
                End If
            End If
        End If
    Next
End Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'FUNCTION - IsValidEmail ////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function IsValidEmail(value As String) As Boolean

    Dim RE As Object
    Set RE = CreateObject("vbscript.RegExp") 'Validate Email Regular Expressions entry
    RE.Pattern = "^(([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5}){1,10})+([;.](([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5}){1,10})+)*$"
    IsValidEmail = RE.test(value)
    Set RE = Nothing
End Function
