Attribute VB_Name = "Form_Rebuild"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: Form_Delete ******************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Form Pages Size and placement control
Const PagesSizeWidth As Integer = 1400      'Pages Width Size
Const PagesSizeLeft As Integer = 200        'Pages Left placement
Dim ctrl As MSForms.Control
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TYPES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Type ControlAndAnchors 'Ancoring to Form Objects
    ctl As MSForms.Control
    AnchorTop As Boolean
    AnchorLeft As Boolean
    AnchorBottom As Boolean
    AnchorRight As Boolean
End Type
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Position ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Position(UForm As Object, FormTag As String, Optional Lang As Long)

    Dim w As Long, h As Long
    With UForm
        .StartUpPosition = 1 ' Center Userform in the Middle of the Default Screen
        w = GetSystemMetrics32(0) ' Screen Resolution width in points
        h = GetSystemMetrics32(1) ' Screen Resolution height in points
        On Error Resume Next
            .Width = w * PointsPerPixel * 0.4  'Userform width= Width in Resolution * DPI * 40%
            .Height = h * PointsPerPixel * 0.7  'Userform height= Height in Resolution * DPI * 65%
        On Error GoTo 0
        Select Case Lang
            Case 49 'German
                If .cbx_Kontengruppe.value = "" Or .cbx_Kontengruppe.value = "( Please select )" Then _
                    .cbx_Kontengruppe.value = "( bitte auswählen )"
            Case Else 'English
                If .cbx_Kontengruppe.value = "" Or .cbx_Kontengruppe.value = "( bitte auswählen )" Then _
                    .cbx_Kontengruppe.value = "( Please select )"
        End Select
    End With
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Translate //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Translate(UForm As Object, SwichLang As Long, SwichFeed As String, _
                        Optional UserLanguage As String, Optional UserContTipp As String, _
                        Optional UserRowSourceLang As String, Optional VerkOrgMustFill As String, _
                        Optional VerkOrgRowSource As String, Optional VertWegRowSource As String, _
                        Optional UserTabIndex As String, Optional UChoise As Boolean)
                        
    Dim i As Long, xObjects As Long, SettingColumns As Excel.ListColumns, ControlsAndAnchors() As ControlAndAnchors
    Dim rFound As Range, foundSource As Boolean
    
    Set SettingColumns = xx_frmConst.ListObjects(1).ListColumns
    xObjects = UForm.Controls.Count
    ReDim ControlsAndAnchors(1 To xObjects)
    For i = 1 To xx_frmConst.ListObjects(1).DataBodyRange.Rows.Count
        With ControlsAndAnchors(i)
            Set .ctl = UForm.Controls(SettingColumns("ctrl").DataBodyRange.Rows(i).value)
    'Translate LABEL and tipptext for CHECKBOX and TEXTBOX
            If TypeOf .ctl Is MSForms.Label Or TypeOf .ctl Is MSForms.CheckBox Or TypeOf .ctl Is MSForms.TextBox Then
            'Initialize-------LABEL and tipptext for CHECKBOX and TEXTBOX------------------------------------------------------
                If SwichLang > 0 And SwichFeed = "Initialize" And Not UChoise Then
                    Set rFound = xx_frmConst.ListObjects(1).ListColumns(2).DataBodyRange.Find(.ctl.Name, , xlFormulas, xlWhole)
                    foundSource = Not rFound Is Nothing
                    If foundSource Then
                        If TypeOf .ctl Is MSForms.Label Or TypeOf .ctl Is MSForms.CheckBox Then
                            'Set UserLanguage.ctl.Object.Caption LABELS
                            .ctl.Object.Caption = SettingColumns(UserLanguage).DataBodyRange.Rows(i).value
                            .ctl.TabIndex = SettingColumns(UserTabIndex).DataBodyRange.Rows(i).value 'Fix Values
                            'Set UserContTipp.ctl.ControlTipText CHECKBOX or IMAGES
                            .ctl.ControlTipText = SettingColumns(UserContTipp).DataBodyRange.Rows(i).value
                        End If
                        If TypeOf .ctl Is MSForms.TextBox Then
                            'Set UserContTipp.ctl.ControlTipText TEXTBOX
                            .ctl.ControlTipText = SettingColumns(UserContTipp).DataBodyRange.Rows(i).value
                        End If
                    End If
                End If
            '-------------------------------------------------------------------------------------------------------------------
            'Translate-------LABEL and tipptext for CHECKBOX and TEXTBOX-Setup Controll top_lbl_C_UK_Click Or top_lbl_C_DE_Click
                If SwichLang > 0 And SwichFeed = "Translate" And UChoise Then
                    Set rFound = xx_frmConst.ListObjects(1).ListColumns(2).DataBodyRange.Find(.ctl.Name, , xlFormulas, xlWhole)
                    foundSource = Not rFound Is Nothing
                    If foundSource Then
                        If TypeOf .ctl Is MSForms.Label Or TypeOf .ctl Is MSForms.CheckBox Then
                            'Set UserLanguage.ctl.Object.Caption LABELS
                            .ctl.Object.Caption = SettingColumns(UserLanguage).DataBodyRange.Rows(i).value
                            .ctl.TabIndex = SettingColumns(UserTabIndex).DataBodyRange.Rows(i).value 'Fix Values
                            'Set UserContTipp.ctl.ControlTipText CHECKBOX or IMAGES
                            .ctl.ControlTipText = SettingColumns(UserContTipp).DataBodyRange.Rows(i).value
                        End If
                        If TypeOf .ctl Is MSForms.TextBox Then
                            'Set UserContTipp.ctl.ControlTipText TEXTBOX
                            .ctl.ControlTipText = SettingColumns(UserContTipp).DataBodyRange.Rows(i).value
                        End If
                    End If
                End If
            End If
        '------------------------------------------------------------------------------------------------------------------------
    'Translate COMBOBOX
            If TypeOf .ctl Is MSForms.ComboBox Then
            'Initialize-------COMBOBOX-------------------------------------------------------------------------------------------
                If SwichLang > 0 And SwichFeed = "Initialize" And Not UChoise Then
                    Set rFound = xx_frmConst.ListObjects(1).ListColumns(2).DataBodyRange.Find(.ctl.Name, , xlFormulas, xlWhole)
                    foundSource = Not rFound Is Nothing
                    If foundSource Then 'Set UserRowSourceLang.ctl.RowSource
                        .ctl.RowSource = SettingColumns(UserRowSourceLang).DataBodyRange.Rows(i).value
                        .ctl.TabIndex = SettingColumns(UserTabIndex).DataBodyRange.Rows(i).value 'Fix Values
                    End If
                End If
            '--------------------------------------------------------------------------------------------------------------------
            'Translate-------COMBOBOX-------------------------------------Setup Controll top_lbl_C_UK_Click Or top_lbl_C_DE_Click
                If SwichLang > 0 And SwichFeed = "Translate" And UChoise Then
                    Set rFound = xx_frmConst.ListObjects(1).ListColumns(2).DataBodyRange.Find(.ctl.Name, , xlFormulas, xlWhole)
                    foundSource = Not rFound Is Nothing
                    If foundSource Then 'Set UserRowSourceLang.ctl.RowSource
                    If UChoise Then .ctl.value = "" 'User Clicked on change Language Form Tag = "New"
                        .ctl.RowSource = SettingColumns(UserRowSourceLang).DataBodyRange.Rows(i).value
                        .ctl.TabIndex = SettingColumns(UserTabIndex).DataBodyRange.Rows(i).value 'Fix Values
                    End If
                End If
            '--------------------------------------------------------------------------------------------------------------------
            'Kontengruppe-------COMBOBOX----------------------------------------------OR---Setup Controll cbx_Kontengruppe_Change
                If SwichLang > 0 And SwichFeed = "Kontengruppe" And UChoise Then
                    'Set Defalut MustFillFieldsTag for All
                    Set rFound = xx_frmConst.ListObjects(1).ListColumns(2).DataBodyRange.Find(.ctl.Name, , xlFormulas, xlWhole)
                    foundSource = Not rFound Is Nothing
                    If foundSource Then 'Set UserRowSourceLang.ctl.RowSource
                    .ctl.RowSource = SettingColumns(UserRowSourceLang).DataBodyRange.Rows(i).value
                    .ctl.TabIndex = SettingColumns(UserTabIndex).DataBodyRange.Rows(i).value 'Fix Values
                    End If
                End If
            '--------------------------------------------------------------------------------------------------------------------
            'VerkOrg-----------COMBOBOX---------------------------------------OR---Setup Controll cbx_Verkaufsorganisation_Change
                If SwichLang = 0 And SwichFeed = "VerkOrg" Then
                    'Set Defalut VerkOrgMustFill for VerkOrg
                    Set rFound = xx_frmConst.ListObjects(1).ListColumns(2).DataBodyRange.Find(.ctl.Name, , xlFormulas, xlWhole)
                    foundSource = Not rFound Is Nothing
                    If foundSource Then 'Set VerkOrgMustFill.ctl.RowSource
                        If .ctl.RowSource = "" Then .ctl.RowSource = SettingColumns(VerkOrgRowSource).DataBodyRange.Rows(i).value
                        .ctl.TabIndex = SettingColumns(UserTabIndex).DataBodyRange.Rows(i).value 'Fix Values
                    End If
                    'Set MustFillAll.ctl.Tag
                    If foundSource And rFound.value <> "" Then .ctl.Tag = SettingColumns(VerkOrgMustFill).DataBodyRange.Rows(i).value
                End If
            '--------------------------------------------------------------------------------------------------------------------
            'VertWeg-----------COMBOBOX-----------------------------------------------OR---Setup Controll cbx_Vertriebsweg_Change
                If SwichLang = 0 And SwichFeed = "VertWeg" Then 'SwichLang = 0 and SwichFeed = "VertWegRowSource"
                    'Set Defalut VertWegRowSource for VertWeg
                    Set rFound = xx_frmConst.ListObjects(1).ListColumns(2).DataBodyRange.Find(.ctl.Name, , xlFormulas, xlWhole)
                    foundSource = Not rFound Is Nothing
                    If foundSource Then 'Set VertWegRowSource.ctl.RowSource
                        If .ctl.RowSource = "" Then .ctl.RowSource = SettingColumns(VertWegRowSource).DataBodyRange.Rows(i).value
                        .ctl.TabIndex = SettingColumns(UserTabIndex).DataBodyRange.Rows(i).value 'Fix Values
                    End If
                End If
            End If
        '-------------------------------------------------------------------------------------------------------------------------
    'Setup Tags fo TEXTBOX and CHECKBOX
            If TypeOf .ctl Is MSForms.TextBox Or TypeOf .ctl Is MSForms.CheckBox Then
                'VerkOrg----------TEXTBOX and CHECKBOX-------------------------OR---Setup Controll cbx_Verkaufsorganisation_Change
                If SwichLang = 0 And SwichFeed = "VerkOrg" Then 'SwichLang = 0 and SwichFeed = "VerkOrgMustFill"
                    'Set Defalut MustFillFieldsTag and TabIndex for All
                    Set rFound = xx_frmConst.ListObjects(1).ListColumns(2).DataBodyRange.Find(.ctl.Name, , xlFormulas, xlWhole)
                    foundSource = Not rFound Is Nothing
                    'Set MustFillAll.ctl.Tag
                    If foundSource And rFound.value <> "" Then .ctl.Tag = SettingColumns(VerkOrgMustFill).DataBodyRange.Rows(i).value
                    .ctl.TabIndex = SettingColumns(UserTabIndex).DataBodyRange.Rows(i).value 'Fix Values
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------------------------------------
        End With
    Next i
    SwichFeed = "" 'Reset Swich to nothing
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'UpdateFormControls /////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub UpdateFormControls(UForm As Object, Optional activeObject As Object)
    
    Dim NotTestMode As Boolean, UserDesigner As Boolean
    With UForm
        If .chb_Testmode_Ein_Aus.value = False Then NotTestMode = True
        If .mpg_Einstieg.Visible = True And InStr(1, .cbx_Kontengruppe.value, "(") <> 1 Then 'Condition to enable Next Button
            .top_lbl_B_Continue.Visible = True 'Display Continue Button
            If .top_lbl_C_UK.Visible = False Then .lbl_FormMode.Top = 6
        Else
            .top_lbl_B_Continue.Visible = False 'Hide Continue Buttons
            If .top_lbl_C_UK.Visible = False Then .lbl_FormMode.Top = 6
        End If
    '------------------------------------------------------------------------------------------------------------------------
    'Control Tag, Position and BackgroundColor ------------------------------------------------------------------------------
        For Each ctrl In .Controls
            If TypeOf ctrl Is MSForms.Label Then ' Change Controlls to appear Active
                If ctrl.Name Like "top_lbl_B_*" Then ctrl.Object.ForeColor = RGB(166, 166, 166) 'Gray if not active
                If ctrl.Name Like "top_lbl_*" Then
                    If Not ctrl.Name Like "top_lbl_B_*" Then
                        With ctrl.Object 'Label Back-, Font-, Border-Color setup if inactive Object MouseMove
                            .BackStyle = fmBackStyleTransparent
                            .ForeColor = RGB(0, 112, 192) 'Blue
                            .BorderColor = RGB(0, 112, 192) 'Blue
                        End With
                    End If
                End If
                If ctrl.Name Like "top_lbl_Approve" Then ctrl.Object.ForeColor = RGB(0, 176, 80) 'Green
                If Not activeObject Is Nothing Then
                    If ctrl.Name = activeObject.Name Then
                        With ctrl.Object 'Label Back-, Font-, Border-Color setup if active Object MouseMove
                            .BackStyle = fmBackStyleOpaque
                            .ForeColor = RGB(255, 255, 255) 'White
                            .BorderColor = RGB(0, 112, 192) 'Blue
                        End With
                        If activeObject.Name Like "top_lbl_B_*" Then ctrl.Object.ForeColor = RGB(0, 112, 192) 'Blue if active
                    End If
                End If
            End If
            If TypeOf ctrl Is MSForms.ComboBox Or TypeOf ctrl Is MSForms.TextBox Then ' Setup Combobox and Textbox wich are mandatory
            '**********Exception ***************
                If ctrl.Name Like "*Debitor_SAPNr*" And UForm.chb_Reaktivieren.value = False Then ctrl.Tag = ""
            '**********Mustfill Setup **********
                If ctrl.Tag = "MustFill" And ctrl.value = "" And NotTestMode Then ctrl.Object.BackColor = RGB(255, 255, 204) 'Yellow
            '**********NotFill Setup ***********
                'If ctrl.Tag = "NotFill" Then ctrl.Object.BackColor = RGB(242, 242, 242) 'Light Gray  'White
                If ctrl.Tag = "NotFill" Then ctrl.Object.BackColor = RGB(255, 255, 255) 'White
            End If
            'Approve Edit or View Modus
            If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Or TypeOf ctrl Is MSForms.CheckBox Then
                If Not ctrl.Name Like "chb_Testmode*" Then
                    If UForm.Tag = "Approve" And ctrl.Parent.Parent.Name <> "mpgControlPanel" Then
                        ctrl.Locked = True 'Lock for Editing
                            If UForm.chb_Approved.value = False Then 'Enable/ Disable Delete and Save if data Is Approved or Not Approved
                                UForm.lbl_FormMode.Caption = "Approve request/ Anforderung Freigeben"
                                UForm.top_lbl_D_Save.Visible = True
                            Else
                                UForm.lbl_FormMode.Caption = "Approve request/ Anforderung Freigeben <<Approved/ Freigeben>>"
                                UForm.top_lbl_D_Delete.Visible = False
                                UForm.top_lbl_D_Save.Visible = False
                            End If
                        UForm.top_lbl_B1_Warenempfaenger.Enabled = False
                        UForm.top_lbl_B1_Rechnungsempfaenger.Enabled = False
                    End If
                    If UForm.Tag = "Edit" And UForm.chb_Approved.value = True Then 'Edit Lock if Approved
                        UForm.lbl_FormMode.Caption = "Edit request/ Anforderung Ändern <<Approved/ Freigeben>>"
                        ctrl.Locked = True
                        UForm.top_lbl_D_Delete.Visible = False
                        UForm.top_lbl_D_Save.Visible = False
                        UForm.top_lbl_B1_Warenempfaenger.Enabled = False
                        UForm.top_lbl_B1_Rechnungsempfaenger.Enabled = False
                    End If
                    If UForm.Tag = "View" Then 'View Lock
                        ctrl.Locked = True
                        UForm.top_lbl_D_Delete.Visible = False
                        UForm.top_lbl_D_Save.Visible = False
                        UForm.top_lbl_B1_Warenempfaenger.Enabled = False
                        UForm.top_lbl_B1_Rechnungsempfaenger.Enabled = False
                    End If
                End If
            End If
        Next ctrl
    '------------------------------------------------------------------------------------------------------------------------
    '/* Display AllgemeineDaten MultiPage Labels and Buttons ----------------------------------------------------------------
        If .mpg_AllgemeineDaten.Visible = True Then 'Multipage is visible
            .mpg_Vertriebsbereichsdaten.Visible = False
            .top_lbl_B_VertDaten_Back.Visible = False
            .top_lbl_B_VertDaten_Next.Visible = False
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
            If .mpg_AllgemeineDaten.value = 1 Then .top_lbl_go_Einstieg.Visible = False 'Hide mpg_Einstieg Control Button
            If .mpg_AllgemeineDaten.value = 3 And .lbl_Warenempfaenger.Visible = False Then _
                .top_lbl_B1_Warenempfaenger.Visible = True 'Show Warenempfänger Label
            If .mpg_AllgemeineDaten.value = 4 And .lbl_Rechnungsempfaenger.Visible = False Then _
                .top_lbl_B1_Rechnungsempfaenger.Visible = True 'Show Rechnungsempfänger Label
            .top_lbl_B_AllgDaten_Back.Visible = .mpg_AllgemeineDaten.value > 0 'Enable Back Button if not first Page
            'Next Button Visible, hide Back Button
            If .mpg_AllgemeineDaten.value = 0 Then .top_lbl_go_AllgemeineDaten.Visible = False: .top_lbl_go_Einstieg.Visible = True
            ' Enable Next Button if not last Page
            .top_lbl_B_AllgDaten_Next.Visible = .mpg_AllgemeineDaten.value < .mpg_AllgemeineDaten.Pages.Count - 1
            If .mpg_AllgemeineDaten.value = .mpg_AllgemeineDaten.Pages.Count - 1 Then 'Hide Next Button
                .top_lbl_B_AllgDaten_Next.Visible = False 'Hide Next button for AllgemeineDaten Page next
                .top_lbl_go_Vertriebsbereichsdaten.Visible = True 'Show next Page Button
            End If
            'Buttons setup if the request is just for Sold to party
            If .mpg_AllgemeineDaten.value = 3 And InStr(1, .cbx_Kontengruppe.value, "KUNW") <> 0 Then
                .top_lbl_B_AllgDaten_Back.Visible = False
                .top_lbl_B_AllgDaten_Next.Visible = False
                .top_lbl_H2_Vertriebsbereichsdaten.Visible = False
                .top_lbl_H3_Adresse.Visible = False
                .top_lbl_H3_1_Steuerung_Zusatzdaten.Visible = False
                .top_lbl_H3_2_Ansprechpartner.Visible = False
                .top_lbl_H3_4_Rechnungsempfaenger.Visible = False
            End If
            'Color Top Label Buttons
            If .mpg_AllgemeineDaten.value = 0 Then 'Page Adresse
                .top_lbl_H3_Adresse.BackStyle = fmBackStyleOpaque
                .top_lbl_H3_Adresse.ForeColor = RGB(255, 255, 255) 'White
            End If
            If .mpg_AllgemeineDaten.value = 1 Then 'Steuerung_Zusatzdaten
                .top_lbl_H3_1_Steuerung_Zusatzdaten.BackStyle = fmBackStyleOpaque
                .top_lbl_H3_1_Steuerung_Zusatzdaten.ForeColor = RGB(255, 255, 255) 'White
            End If
            If .mpg_AllgemeineDaten.value = 2 Then  'Ansprechpartner
                .top_lbl_H3_2_Ansprechpartner.BackStyle = fmBackStyleOpaque
                .top_lbl_H3_2_Ansprechpartner.ForeColor = RGB(255, 255, 255) 'White
            End If
            If .mpg_AllgemeineDaten.value = 3 Then 'Warenempfaenger
                .top_lbl_B1_Rechnungsempfaenger.Visible = False
                If .lbl_Warenempfaenger.Visible = False Then
                    With .top_lbl_B1_Warenempfaenger
                        .Visible = True
                        .Left = PagesSizeLeft
                        .Width = PagesSizeWidth - 15
                    End With
                End If
                .top_lbl_H3_3_Warenempfaenger.BackStyle = fmBackStyleOpaque
                .top_lbl_H3_3_Warenempfaenger.ForeColor = RGB(255, 255, 255) 'White
                .top_lbl_Warenempfaenger_Abbruch.Width = PagesSizeWidth - 25
            Else
                .top_lbl_B1_Warenempfaenger.Visible = False
            End If
            If .mpg_AllgemeineDaten.value = 4 Then  'Rechnungsempfaenger
                .top_lbl_B1_Warenempfaenger.Visible = False
                If .lbl_Rechnungsempfaenger.Visible = False Then
                    With .top_lbl_B1_Rechnungsempfaenger
                        .Visible = True
                        .Left = PagesSizeLeft
                        .Width = PagesSizeWidth - 15
                    End With
                End If
                .top_lbl_H3_4_Rechnungsempfaenger.BackStyle = fmBackStyleOpaque
                .top_lbl_H3_4_Rechnungsempfaenger.ForeColor = RGB(255, 255, 255) 'White
                .top_lbl_Rechnungsempfaenger_Abbruch.Width = PagesSizeWidth - 25
            Else
                .top_lbl_B1_Rechnungsempfaenger.Visible = False
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------------
    '/* Display Vertriebsbereichsdaten MultiPage Labels and Buttons ---------------------------------------------------------
        If .mpg_Vertriebsbereichsdaten.Visible = True Then 'Multipage is visible
            .mpg_AllgemeineDaten.Visible = False
            .top_lbl_go_Einstieg.Visible = False
            .top_lbl_B_AllgDaten_Back.Visible = False
            .top_lbl_B_AllgDaten_Next.Visible = False
            .top_lbl_B1_Warenempfaenger.Visible = False
            .top_lbl_B1_Rechnungsempfaenger.Visible = False
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
            .top_lbl_B_VertDaten_Back.Visible = .mpg_Vertriebsbereichsdaten.value > 0 ' Enable Back Button if not first Page
            'Next Button Visible, hide Back Button
            If .mpg_Vertriebsbereichsdaten.value = 0 Then .top_lbl_go_AllgemeineDaten.Visible = True
            'Enable Next Button if not last Page
            .top_lbl_B_VertDaten_Next.Visible = .mpg_Vertriebsbereichsdaten.value < .mpg_Vertriebsbereichsdaten.Pages.Count - 1
            'Hide next to Page Button
            If .mpg_Vertriebsbereichsdaten.value = .mpg_Vertriebsbereichsdaten.Pages.Count - 1 Then _
                .top_lbl_B_VertDaten_Next.Visible = False
            'Color Top Label Buttons
            If .mpg_Vertriebsbereichsdaten.value = 0 Then   'Verkauf
                .top_lbl_H3_Verkauf.BackStyle = fmBackStyleOpaque
                .top_lbl_H3_Verkauf.ForeColor = RGB(255, 255, 255) 'White
            End If
            If .mpg_Vertriebsbereichsdaten.value = 1 Then   'Versand
                .top_lbl_H3_1_Versand.BackStyle = fmBackStyleOpaque
                .top_lbl_H3_1_Versand.ForeColor = RGB(255, 255, 255) 'White
            End If
            If .mpg_Vertriebsbereichsdaten.value = 2 Then   'Faktura
                .top_lbl_H3_2_Faktura.BackStyle = fmBackStyleOpaque
                .top_lbl_H3_2_Faktura.ForeColor = RGB(255, 255, 255) 'White
            End If
            If .mpg_Vertriebsbereichsdaten.value = 3 Then   'Partnerrolle_Zusatzdaten
                .top_lbl_H3_3_Partnerrolle_Zusatzdaten.BackStyle = fmBackStyleOpaque
                .top_lbl_H3_3_Partnerrolle_Zusatzdaten.ForeColor = RGB(255, 255, 255) 'White
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------------
    End With
End Sub
