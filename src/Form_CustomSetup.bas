Attribute VB_Name = "Form_CustomSetup"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: Form_CustomSetup *************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim ctrl As MSForms.Control
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CustomSettings /////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub CustomSettings(UForm As Object)

    If UForm.Tag Like "New" Then
        With UForm
            'Partner Funktion Fields for 2961 = Automation and Reservation -> Partner 1 and Partner 2
            If .cbx_Kontengruppe.value Like "*KUNA*" And .cbx_Verkaufsorganisation.value Like "*2961*" Then
                If .tbx_UserLangVal.value = 49 Then 'German
                    If .cbx_Partnerrolle1.value = "" Then .cbx_Partnerrolle1.value = "ZF - Fax-/Mailempfänger"
                    If .cbx_Partnerrolle2.value = "" Then
                        If .cbx_Verkeaufergruppe.value Like "B06*" Then .cbx_Partnerrolle2.value = "ZP - Provisionsvertreter"
                        If .cbx_Verkeaufergruppe.value Like "A11*" Then .cbx_Partnerrolle2.value = "AP - Ansprechpartner"
                    End If
                ElseIf .tbx_UserLangVal.value = 1 Then 'English
                    If .cbx_Partnerrolle1.value = "" Then .cbx_Partnerrolle1.value = "ZF - Fax-/Email recipient"
                    If .cbx_Partnerrolle2.value = "" Then
                        If .cbx_Verkeaufergruppe.value Like "B06*" Then .cbx_Partnerrolle2.value = "ZP - Commision Repres."
                        If .cbx_Verkeaufergruppe.value Like "A11*" Then .cbx_Partnerrolle2.value = "CP - Contact Person (AP)"
                    End If
                End If
                If .cbx_Vertriebsweg.value Like "*HD*" Then
                    If .cbx_Partner_Nr1.value = "" Then .cbx_Partner_Nr1.value = "589327 - TTM GmbH Internationale"
                ElseIf .cbx_Vertriebsweg.value Like "*GY*" Then
                    If .cbx_Land.value Like "*DE*" Then
                        .cbx_Partner_Nr1.value = "644681 - DE Gygli Sammeladresse" 'If Country is Germany
                    Else
                        .cbx_Partner_Nr1.value = "645961 - EN Gygli Sammeladresse Export" 'If Country is Export
                    End If
                Else
                    If .cbx_Partner_Nr1.value = "" Then .cbx_Partner_Nr1.value = "650276 - Fiege Logistik Stiftung & Co."
                End If
                If .cbx_Verkeaufergruppe.value Like "B06*" And .cbx_Partner_Nr2.value = "" Then _
                        .cbx_Partner_Nr2.value = "531060 - B06 - P.Grossegoedinghaus" 'Provizionvertreter -> GGH Petra
                If .cbx_Verkeaufergruppe.value Like "A11*" And .cbx_Partner_Nr2.value = "" Then _
                        .cbx_Partner_Nr2.value = "130757 - Reiner" 'Ansprechpartner Firma REINER
                'Costum default field settings (Autocomplete)
                .chb_Komplettlief_vorgeschrieben.value = True
            End If
            'Automations Default settings for Sales Organization 3661 Italy
            If .cbx_Kontengruppe.value Like "*KUNA*" And .cbx_Verkaufsorganisation.value Like "*3661*" Then
                If .tbx_UserLangVal.value = 49 Then 'German
                    .chb_AuftrZusammenfuerung.value = True
                    If .cbx_TeilieferungJe_Position.value = "" Then .cbx_TeilieferungJe_Position.value = "_ - Teillieferung erlaubt"
                    If .tbx_Teillieferung_Max.value = "" Then .tbx_Teillieferung_Max.value = 9
                    .chb_Bonus.value = True
                    If .cbx_Rechnungstermine.value = "" Then .cbx_Rechnungstermine.value = "IT - Fabrikkalender Italien Standard"
                Else 'English
                    .chb_AuftrZusammenfuerung.value = True
                    If .cbx_TeilieferungJe_Position.value = "" Then .cbx_TeilieferungJe_Position.value = "_ - Partial delivery allowed"
                    If .tbx_Teillieferung_Max.value = "" Then .tbx_Teillieferung_Max.value = 9
                    .chb_Bonus.value = True
                    If .cbx_Rechnungstermine.value = "" Then .cbx_Rechnungstermine.value = "IT - Factory calendar Italy standard"
                End If
            End If
        End With
    End If
End Sub
