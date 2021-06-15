VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: clsCheckBox ******************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Top Navigation Buttons control
Dim UParent As Object
Const ButtonsHide As Integer = 70           'Hide Buttons to top
Private WithEvents lblInForm As MSForms.Label
Attribute lblInForm.VB_VarHelpID = -1
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Control ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Set Control(lblContainer As MSForms.Label)

    Set lblInForm = lblContainer
End Property
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'lblInForm_Click ////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub lblInForm_Click()

    Set UParent = lblInForm.Parent
    With lblInForm.Parent
        If lblInForm.Name Like "*go_Einstieg*" Then
            If Application.Run("Form_Validate.ValidationFailed", .mpg_Einstieg, .mpg_Einstieg.value, .tbx_UserLangVal) Then Exit Sub
            .top_lbl_go_Einstieg.Visible = False
            .top_lbl_go_AllgemeineDaten.Visible = False
            .top_lbl_B_AllgDaten_Next.Visible = False
            .top_lbl_B_Continue.Visible = True
            .mpg_Einstieg.Visible = True
            .mpg_AllgemeineDaten.Visible = False
            .mpg_Vertriebsbereichsdaten.Visible = False
            .top_lbl_H2_AllgemeineDaten.Visible = False
            .top_lbl_H2_Vertriebsbereichsdaten.Visible = False
            .top_lbl_H3_Adresse.Visible = False
            .top_lbl_H3_1_Steuerung_Zusatzdaten.Visible = False
            .top_lbl_H3_2_Ansprechpartner.Visible = False
            .top_lbl_H3_3_Warenempfaenger.Visible = False
            .top_lbl_H3_4_Rechnungsempfaenger.Visible = False
            Application.Run "Form_Rebuild.UpdateFormControls", UParent
        End If
        If lblInForm.Name Like "*AllgemeineDaten*" Then
            If Application.Run("Form_Validate.ValidationFailed", .mpg_Vertriebsbereichsdaten, _
                .mpg_Vertriebsbereichsdaten.value, .tbx_UserLangVal) Then Exit Sub
            If .mpg_AllgemeineDaten.value = .mpg_AllgemeineDaten.Pages.Count - 1 Then _
                .top_lbl_go_Vertriebsbereichsdaten.Visible = True
            .top_lbl_go_AllgemeineDaten.Visible = False
            .mpg_Vertriebsbereichsdaten.Visible = False
            .mpg_AllgemeineDaten.Visible = True
            .top_lbl_H2_AllgemeineDaten.Visible = False
            .top_lbl_H2_Vertriebsbereichsdaten.Visible = True
            .top_lbl_H3_Verkauf.Visible = False
            .top_lbl_H3_Adresse.Visible = True
            .top_lbl_H3_1_Versand.Visible = False
            .top_lbl_H3_1_Steuerung_Zusatzdaten.Visible = True
            .top_lbl_H3_2_Faktura.Visible = False
            .top_lbl_H3_2_Ansprechpartner.Visible = True
            .top_lbl_H3_3_Partnerrolle_Zusatzdaten.Visible = False
            .top_lbl_H3_3_Warenempfaenger.Visible = True
            .top_lbl_H3_4_Rechnungsempfaenger.Visible = True
            If .mpg_AllgemeineDaten.value = 3 And .lbl_Warenempfaenger.Visible = False Then _
                .top_lbl_B1_Warenempfaenger.Visible = True
            If .mpg_AllgemeineDaten.value = 4 And .lbl_Rechnungsempfaenger.Visible = False Then _
                .top_lbl_B1_Rechnungsempfaenger.Visible = True
            Application.Run "Form_Rebuild.UpdateFormControls", UParent
        End If
        If lblInForm.Name Like "*Vertriebsbereichsdaten*" Then
            If Application.Run("Form_Validate.ValidationFailed", .mpg_AllgemeineDaten, _
                .mpg_AllgemeineDaten.value, .tbx_UserLangVal) Then Exit Sub
            If .mpg_Vertriebsbereichsdaten.value = 0 Then .top_lbl_go_AllgemeineDaten.Visible = True
            .top_lbl_go_Vertriebsbereichsdaten.Visible = False
            .top_lbl_B1_Warenempfaenger.Visible = False
            .top_lbl_B1_Rechnungsempfaenger.Visible = False
            .mpg_AllgemeineDaten.Visible = False
            .mpg_Vertriebsbereichsdaten.Visible = True
            .top_lbl_H2_Vertriebsbereichsdaten.Visible = False
            .top_lbl_H2_AllgemeineDaten.Visible = True
            .top_lbl_H3_Adresse.Visible = False
            .top_lbl_H3_Verkauf.Visible = True
            .top_lbl_H3_1_Steuerung_Zusatzdaten.Visible = False
            .top_lbl_H3_1_Versand.Visible = True
            .top_lbl_H3_2_Ansprechpartner.Visible = False
            .top_lbl_H3_2_Faktura.Visible = True
            .top_lbl_H3_3_Warenempfaenger.Visible = False
            .top_lbl_H3_3_Partnerrolle_Zusatzdaten.Visible = True
            .top_lbl_H3_4_Rechnungsempfaenger.Visible = False
            Application.Run "Form_Rebuild.UpdateFormControls", UParent
        End If
    End With
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'lblInForm_MouseMove ////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub lblInForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If lblInForm.Name Like "*Abbruch*" Or lblInForm.Name Like "top_lbl_Approve" Then
        Set UParent = lblInForm.Parent.Parent.Parent
        Application.Run "Form_Rebuild.UpdateFormControls", UParent, lblInForm
    ElseIf lblInForm.Name Like "*top_lbl_H*" Or lblInForm.Name Like "*top_lbl_B*" Or _
            lblInForm.Name Like "*top_lbl_C*" Or lblInForm.Name Like "*top_lbl_D*" Then
        Set UParent = lblInForm.Parent
        Application.Run "Form_Rebuild.UpdateFormControls", UParent, lblInForm
    End If
End Sub
