Attribute VB_Name = "Form_Messages"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: Form_Messages ****************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'ShowMessage ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub ShowMessage(xLang As Long, msgOperator As String, Control As Object, _
                        Optional CtrlParent As Object, Optional ParentIdx As Integer, Optional pName As String)

    If pName = "mpg_AllgemeineDaten" Then CtrlParent.Parent.mpg_Vertriebsbereichsdaten.Visible = False
    If pName = "mpg_Vertriebsbereichsdaten" Then CtrlParent.Parent.mpg_AllgemeineDaten.Visible = False
    If Not CtrlParent.Visible Then CtrlParent.Visible = True
    CtrlParent.value = ParentIdx
    Select Case xLang
        Case Is = 49 'Messages displayed in German
        Select Case msgOperator
            Case "MustFill"
                MsgBox "Die gelben Felder muss ausgefüllet werden !" & vbNewLine & _
                        vbNewLine & " ", vbExclamation, "Muss-Felder"
                Application.Run "Form_Rebuild.UpdateFormControls", CtrlParent.Parent, CtrlParent
                Control.SetFocus
            Case "OnlyNumbers"
                MsgBox "SAP Nr. sind nur Zahlen erlaubt, Mindest. Länge 'Siehe Tipps' Zeichen." & vbNewLine _
                        & vbNewLine & vbNewLine & "Darf nicht: " & Control.Text, _
                        vbExclamation, "Eingegebene Nummer überprüfen."
                Application.Run "Form_Rebuild.UpdateFormControls", CtrlParent.Parent, CtrlParent
                Control.SetFocus
            Case "TelFax"
                MsgBox "Es wird nur Zahlen erlaubt und darf keine Ländervorwahl oder sonderzeichen enthalten." _
                        & vbNewLine & vbNewLine & "Beispiel: " & "01 234 567 8910" _
                        & vbNewLine & vbNewLine & "Darf nicht: " & Control.Text, _
                        vbExclamation, "Telefonnummer darf keine Ländervorwahl enthalten. Bitte überprüfen."
                Application.Run "Form_Rebuild.UpdateFormControls", CtrlParent.Parent, CtrlParent
                Control.SetFocus
            Case "Email"
                MsgBox "Die email Adresse ist nicht korrekt geschrieben worden !" _
                        & vbNewLine & vbNewLine & "Bitte korrigieren: " & Control.Text, _
                        vbExclamation, "E-Mail ist nicht gültig. Bitte überprüfen."
                Application.Run "Form_Rebuild.UpdateFormControls", CtrlParent.Parent, CtrlParent
                Control.SetFocus
            Case Else
                'none
        End Select
    Case Else
        Select Case msgOperator 'Messages displayed in English
            Case "MustFill"
                MsgBox "Must enter values in yellow fields !" & vbNewLine & _
                        vbNewLine & " ", vbExclamation, "Must fields"
                Application.Run "Form_Rebuild.UpdateFormControls", CtrlParent.Parent, CtrlParent
                Control.SetFocus
            Case "OnlyNumbers" 'SAP numbers are only allowed numbers.
                MsgBox "SAP Nr. are only numbers are allowed, min. Length 'See Tipp' characters." & vbNewLine _
                        & vbNewLine & vbNewLine & "Must not: " & Control.Text, _
                        vbExclamation, "Check entered number."
                Application.Run "Form_Rebuild.UpdateFormControls", CtrlParent.Parent, CtrlParent
                Control.SetFocus
            Case "TelFax"
                MsgBox "Only numbers are allowed and without the country code or special characters." _
                        & vbNewLine & vbNewLine & "Example: " & "0123 456 789 10" _
                        & vbNewLine & vbNewLine & "Must not: " & Control.Text, _
                        vbExclamation, "Telephone number should not contain the country code. Please check."
                Application.Run "Form_Rebuild.UpdateFormControls", CtrlParent.Parent, CtrlParent
                Control.SetFocus
            Case "Email"
                MsgBox "The e-mail address is not an e-mail !" _
                        & vbNewLine & vbNewLine & "Please check: " & Control.Text, _
                        vbExclamation, "Not a valid e-mail. Verify."
                Application.Run "Form_Rebuild.UpdateFormControls", CtrlParent.Parent, CtrlParent
                Control.SetFocus
            Case Else
                'none
        End Select
    End Select
End Sub
