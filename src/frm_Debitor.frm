VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Debitor 
   ClientHeight    =   27420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   27765
   OleObjectBlob   =   "frm_Debitor.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frm_Debitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: frm_Debitor ******************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'GLOBAL VARIABLES ///////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Top Navigation Buttons control
Const ButtonsTop As Integer = 6             'Move Buttons to Top
Const ButtonsStartLeft As Integer = 6       'Move Buttons to Left
Const ButtonsNext As Integer = 500          'Move Buttons to Left "Next Page Button"
Const ButtonsHide As Integer = 70           'Hide Buttons to top
Const ButtonsBack As Integer = 480          'Move Buttons to Left "Back Page Button"
Const ButtonSaveLeft As Integer = 420       'Move Buttons to Left "Save"
Const ButtonDelLeft As Integer = 380        'Move Buttons to Left "Delete"

'Form Size control
Const FormSizeHeight As Integer = 590       'Form Height Size
Const FormSizeWidth As Integer = 530        'Form Width Size
Const FormSeparatorTop As Integer = 28      'Form Top Line
Const FormSeparatorLeft As Integer = 198.5  'Form Top Line

'Left Navigation Labels control
Const HeaderPlaceTop_H1 As Integer = 32     'H1 Header Top placement
Const HeaderPlaceLeft_H1 As Integer = 6     'H1 Header Left placement
Const HeaderPlaceWidth_H1 As Integer = 185  'H1 Width Size
Const HeaderPlaceTop_H2 As Integer = 32     'H2 Header Top placement
Const HeaderPlaceLeft_H2 As Integer = 18    'H2 Header Left placement
Const HeaderPlaceWidth_H2 As Integer = 173  'H2 Width Size
Const HeaderPlaceTop_H3 As Integer = 55     'H3 Header Top placement
Const HeaderPlaceLeft_H3 As Integer = 30    'H3 Header Left placement
Const HeaderPlaceWidth_H3 As Integer = 161  'H3 Width Size

'Form Pages Size and placement control
Const PagesSizeHeight As Integer = 1400     'Pages Height Size
Const PagesSizeWidth As Integer = 1400      'Pages Width Size
Const PagesSizeTop As Integer = 55          'Pages Top placement
Const PagesSizeLeft As Integer = 200        'Pages Left placement

'Lagels General Tags and adjustments
Const LabelPrefix As String = "lbl"
Const MoveDown As Integer = 2

'Excel Application Password
Const passWordLock As String = "fnextxx"

'ClassModule variables, collections and control initialization
Dim cFormResizing As clsFormResizing
Dim ControlsAndAnchors() As ControlAndAnchors
Dim tbxColl As Collection, tbxObject As clsTextBox
Dim cbxColl As Collection, cbxObject As clsComboBox
Dim chbColl As Collection, chbObject As clsCheckBox
Dim lblColl As Collection, lblObject As clsLabel
'Dim mpgColl As Collection, mpgObject As clsMultiPage
Dim ctrl As MSForms.Control, PObject As MSForms.MultiPage
Dim UserLang As Long, clickPosRow As Long, xObjects As Long
Dim clickPos As Range, xLastColumn As Range, xRange As Range, findValue As Range

'Boolean values for Runtime Settings
Dim firstInit As Boolean, UserAdmin As Boolean, UserDesigner As Boolean, UChoise As Boolean

'Translate Language Settings
Dim SwichFeed As String, UserLanguage As String, UserContTipp As String, UserTabIndex As String, UserRowSourceLang As String

'MustFillFields Settings
Dim MustFillAll As String, VerkOrgMustFill As String

'Conditional "VerkOrg und VertWeg" RowSource
Dim VerkOrgRowSource As String, VertWegRowSource As String

'WorkBook Object variables
Public DebtorFile As Workbook, WorkBookForm As String
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////// ******** PROGRAM START ********* /////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'INITIALIZE - UserForm
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub UserForm_Initialize()

'UserForm Initialize variable setup
Dim i As Long, SettingColumns As Excel.ListColumns, rFound As Range, foundSource As Boolean

    Application.ScreenUpdating = False
    With Me 'Buttons Layout and pozition setup
        .Tag = "Initialize"
        .top_lbl_C_UK.Top = ButtonsTop
        .top_lbl_C_UK.Left = ButtonsStartLeft
        .top_lbl_C_DE.Top = ButtonsTop
        .top_lbl_C_DE.Left = ButtonsStartLeft + 40
        .top_lbl_D_Delete.Top = ButtonsTop
        .top_lbl_D_Delete.Left = ButtonDelLeft
        .top_lbl_D_Save.Top = ButtonsTop
        .top_lbl_D_Save.Left = ButtonSaveLeft
        .top_lbl_go_Einstieg.Top = ButtonsTop - 3
        .top_lbl_go_Einstieg.Left = ButtonsBack
        .top_lbl_B_Continue.Top = ButtonsTop - 3
        .top_lbl_B_Continue.Left = ButtonsNext
        .top_lbl_go_AllgemeineDaten.Top = ButtonsTop - 3
        .top_lbl_go_AllgemeineDaten.Left = ButtonsBack
        .top_lbl_B_AllgDaten_Next.Top = ButtonsTop - 3
        .top_lbl_B_AllgDaten_Next.Left = ButtonsNext
        .top_lbl_B_AllgDaten_Back.Top = ButtonsTop - 3
        .top_lbl_B_AllgDaten_Back.Left = ButtonsBack
        .top_lbl_go_Vertriebsbereichsdaten.Top = ButtonsTop - 3
        .top_lbl_go_Vertriebsbereichsdaten.Left = ButtonsNext
        .top_lbl_B_VertDaten_Next.Top = ButtonsTop - 3
        .top_lbl_B_VertDaten_Next.Left = ButtonsNext
        .top_lbl_B_VertDaten_Back.Top = ButtonsTop - 3
        .top_lbl_B_VertDaten_Back.Left = ButtonsBack
        .Separator_frm_Debitor_1.Top = FormSeparatorTop
        .Separator_frm_Debitor_2.Top = FormSeparatorTop
        .Separator_frm_Debitor_2.Left = FormSeparatorLeft
    End With
    With Me.mpgControlPanel 'Designer Panel setup
        .Top = HeaderPlaceTop_H1 + 23 + 155
        .Width = HeaderPlaceWidth_H1 + 13
        .Height = PagesSizeHeight - 155
    End With
    Me.tbx_TextSAP.Width = HeaderPlaceWidth_H1 + 13
    With Me.lbl_Informationen 'Information Text field pozition setup
        .Left = HeaderPlaceLeft_H1
        .Top = HeaderPlaceTop_H1 + 23 + 155
        .Width = HeaderPlaceWidth_H1
        .Height = PagesSizeHeight - 155
    End With
    Set tbxColl = New Collection
    Set cbxColl = New Collection
    Set chbColl = New Collection
    Set lblColl = New Collection
    'Set mpgColl = New Collection
    Set SettingColumns = xx_frmConst.ListObjects(1).ListColumns
    xObjects = Me.Controls.Count
    ReDim ControlsAndAnchors(1 To xObjects)
    If ActiveCell.Value2 <> "" Then 'Detect Excel installed Language setup or read value from sheet
        Set clickPos = aa_valData.Cells(ActiveCell.Row, 1)
        Set xLastColumn = aa_valData.Cells(ActiveCell.Row, aa_valData.UsedRange.Columns.Count)
        Set xRange = aa_valData.Range(clickPos, xLastColumn)
        Set findValue = xRange.Find(What:="tbx_UserLangVal", LookIn:=xlFormulas)
        If Not findValue Is Nothing Then
            If findValue.Offset(0, 1).value <> "" Then Me.tbx_UserLangVal.value = findValue.Offset(0, 1).value
        End If
    Else
        If Application.International(xlCountryCode) <> 49 Then Me.tbx_UserLangVal.value = 1 Else Me.tbx_UserLangVal.value = 49
    End If
    SwichFeed = "Initialize"
    UChoise = False
    Me.tbx_Boolean.value = UChoise
    If Me.tbx_UserLangVal = 49 Then
        Me.cbx_Kontengruppe.RowSource = "Kontengruppe_DE"
        UserLanguage = "Caption_[DE]" ''Excel in German Language installed
        UserContTipp = "ControlTipText_[DE]" ''Excel in German Language installed
        UserRowSourceLang = "RowSourceAll_[DE]" ''Excel in German Language installed
        UserTabIndex = "TabIndex" ''Fix Values
    Else 'English
        Me.tbx_UserLangVal = 1
        Me.cbx_Kontengruppe.RowSource = "Account_Group_EN"
        UserLanguage = "Caption_[EN]" ''default
        UserContTipp = "ControlTipText_[EN]" ''default
        UserRowSourceLang = "RowSourceAll_[EN]" ''default
        UserTabIndex = "TabIndex" ''Fix Values
    End If
    Application.Run "Form_Rebuild.Translate", Me, Me.tbx_UserLangVal.value, SwichFeed, UserLanguage, _
                    UserContTipp, UserRowSourceLang, "empty", "empty", "empty", UserTabIndex, UChoise 'Default translation
    'Count Objects on Form must be ExactMach with Table "tblBuilder" ctrl count
    For i = 1 To xx_frmConst.ListObjects(1).DataBodyRange.Rows.Count
        With ControlsAndAnchors(i)
            Set .ctl = Me.Controls(SettingColumns("ctrl").DataBodyRange.Rows(i).value)
                .AnchorTop = SettingColumns("AnchorTop").DataBodyRange.Rows(i).value
                .AnchorLeft = SettingColumns("AnchorLeft").DataBodyRange.Rows(i).value
                .AnchorBottom = SettingColumns("AnchorBottom").DataBodyRange.Rows(i).value
                .AnchorRight = SettingColumns("AnchorRight").DataBodyRange.Rows(i).value
            'Object is Multipage
            If TypeOf .ctl Is MSForms.MultiPage Then
                .ctl.Object.Style = fmTabStyleNone
                'Add Multipage to Multipage CollectionClass
                'Set mpgObject = New clsMultiPage: Set mpgObject.Control = .ctl: mpgColl.Add mpgObject
                If InStr(1, .ctl.Name, "mpg_") = 1 Then
                    firstInit = True 'Skip resize Class
                    With .ctl 'Setup MultiPage pozition in User Form
                        .value = 0
                        .Height = PagesSizeHeight
                        .Left = PagesSizeLeft
                        .Top = PagesSizeTop
                        .Width = PagesSizeWidth
                    End With
                    firstInit = False 'Reset Skip resize Class
                End If
            End If
            'Object is Label or CheckBox
            If TypeOf .ctl Is MSForms.Label Or TypeOf .ctl Is MSForms.CheckBox Then
                'Add Label to Label CollectionClass
                If TypeOf .ctl Is MSForms.Label Then _
                                    Set lblObject = New clsLabel: Set lblObject.Control = .ctl: lblColl.Add lblObject
                If Left$(.ctl.Name, Len(LabelPrefix)) = LabelPrefix Then .ctl.Top = .ctl.Top + MoveDown 'Move down an inch Labels
                If .ctl.Name Like "lbl_FormMode" Then 'Setup Header1 pozition in User Form
                    'H1_-------------------------------
                    With .ctl
                        .Left = HeaderPlaceLeft_H1
                        .Top = HeaderPlaceTop_H1
                        .Width = HeaderPlaceWidth_H1 + 165
                        .ZOrder (0)
                    End With
                End If
                If InStr(1, .ctl.Name, "top_lbl_H2_") = 1 Then 'Setup Header2 pozition in User Form
                    'H2--------------------------------
                    With .ctl
                        .Left = HeaderPlaceLeft_H2
                        .Top = HeaderPlaceTop_H2
                        .Width = HeaderPlaceWidth_H2
                    End With
                End If
                If InStr(1, .ctl.Name, "top_lbl_H3_") = 1 Then 'Setup Header3 pozition in User Form
                    'H3---------------------------------
                    With .ctl
                        .Left = HeaderPlaceLeft_H3
                        .Top = HeaderPlaceTop_H3
                        .Width = HeaderPlaceWidth_H3
                    End With
                End If
                If InStr(1, .ctl.Name, "top_lbl_H3_1") = 1 Then 'Setup Header3.1 pozition in User Form
                    'H3_1-------------------------------
                    With .ctl
                        .Left = HeaderPlaceLeft_H3
                        .Top = HeaderPlaceTop_H3 + 23
                        .Width = HeaderPlaceWidth_H3
                    End With
                End If
                If InStr(1, .ctl.Name, "top_lbl_H3_2") = 1 Then 'Setup Header3.2 pozition in User Form
                    'H3_2-------------------------------
                    With .ctl
                        .Left = HeaderPlaceLeft_H3
                        .Top = HeaderPlaceTop_H3 + 46 '(2x23)
                        .Width = HeaderPlaceWidth_H3
                    End With
                End If
                If InStr(1, .ctl.Name, "top_lbl_H3_3") = 1 Then 'Setup Header3.3 pozition in User Form
                    'H3_3-------------------------------
                    With .ctl
                        .Left = HeaderPlaceLeft_H3
                        .Top = HeaderPlaceTop_H3 + 69 '(3x23)
                        .Width = HeaderPlaceWidth_H3
                    End With
                End If
                If InStr(1, .ctl.Name, "top_lbl_H3_4") = 1 Then 'Setup Header3.4 pozition in User Form
                    'H3_4-------------------------------
                    With .ctl
                        .Left = HeaderPlaceLeft_H3
                        .Top = HeaderPlaceTop_H3 + 92 '(4x23)
                        .Width = HeaderPlaceWidth_H3
                    End With
                End If
            End If
            'Object is TextBox, ComboBox or CheckBox
            If TypeOf .ctl Is MSForms.TextBox Or TypeOf .ctl Is MSForms.ComboBox Or TypeOf .ctl Is MSForms.CheckBox Then
                'Add TextBox to TextBox CollectionClass
                If TypeOf .ctl Is MSForms.TextBox Then _
                                            Set tbxObject = New clsTextBox: Set tbxObject.Control = .ctl: tbxColl.Add tbxObject
                'Add ComboBox to ComboBox CollectionClass
                If TypeOf .ctl Is MSForms.ComboBox Then _
                                            Set cbxObject = New clsComboBox: Set cbxObject.Control = .ctl: cbxColl.Add cbxObject
                'Add CheckBox to CheckBox CollectionClass
                If TypeOf .ctl Is MSForms.CheckBox Then _
                                            Set chbObject = New clsCheckBox: Set chbObject.Control = .ctl: chbColl.Add chbObject
                Me.tbx_SelectedRowVal = 0 'Initialize tbx_SelectedRowVal
                If ActiveCell.Value2 <> "" Then 'Read Values from Sheet if User clicked edit button
                    UChoise = False 'UserChoiuse setup to Initialise
                    Me.tbx_Boolean.value = UChoise 'Initialize tbx_Boolean
                    clickPosRow = ActiveCell.Row 'ActiveCell Position
                    Me.tbx_SelectedRowVal = clickPosRow 'reset tbx_SelectedRowVal to selected cellRow
                    Set clickPos = aa_valData.Cells(clickPosRow, 1)
                    Set xLastColumn = aa_valData.Cells(clickPosRow, aa_valData.UsedRange.Columns.Count)
                    Set xRange = aa_valData.Range(clickPos, xLastColumn)
                    Set findValue = xRange.Find(What:=.ctl.Name, LookIn:=xlFormulas)
                    If Not findValue Is Nothing Then 'Value found pair with Object
                        If findValue.Offset(0, 1).value <> "" Then
                            If TypeOf .ctl Is MSForms.ComboBox Then .ctl.Object.Style = fmStyleDropDownCombo 'Set Combobox Style
                            .ctl.Object.value = findValue.Offset(0, 1).value 'Set Object value
                            If TypeOf .ctl Is MSForms.ComboBox And Not .ctl.Name Like "*_Partner_Nr*" Then _
                                .ctl.Object.Style = fmStyleDropDownList 'Reset Combobox Style
                        End If
                        If findValue.value Like "tbx_WE_Name1" And findValue.Offset(0, 1).value <> "" Then _
                                                Call top_lbl_B1_Warenempfaenger_click 'Display WE Page
                        If findValue.value Like "tbx_RE_Name1" And findValue.Offset(0, 1).value <> "" Then _
                                                Call top_lbl_B1_Rechnungsempfaenger_click 'Display RE Page
                    End If
                End If
            End If
        End With
    Next i
    'FreeUp Memory
    Set tbxObject = Nothing 'Distroy TextBox colection ObjectClass
    Set cbxObject = Nothing 'Distroy ComboBox colection ObjectClass
    Set chbObject = Nothing 'Distroy CheckBox colection ObjectClass
    Set lblObject = Nothing 'Distroy Label colection ObjectClass
    'Set mpgObject = Nothing 'Distroy MultiPage colection ObjectClass
    Me.tbx_Boolean.value = "True" 'UserChoiuse setup to reset
    Me.tbx_TextSAP.value = "" 'Clear Autofill
    Set cFormResizing = New clsFormResizing
    cFormResizing.Initialize Me, ControlsAndAnchors 'Resize controls
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'ACTIVATE - UserForm
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub UserForm_Activate()

    With Me
        .ScrollBars = fmScrollBarsBoth 'Display scrollbars
        If .InsideHeight < FormSizeHeight Then 'Setup scrollbar Height
            .ScrollHeight = FormSizeHeight * 1
        Else
            .ScrollHeight = .InsideHeight * 1
        End If
        If .InsideWidth < FormSizeWidth Then 'Setup scrollbar Width
            .ScrollWidth = FormSizeWidth
        Else
            .ScrollWidth = .InsideWidth
        End If
        'Setup Form if User is Admin
        If Application.WorksheetFunction.CountIf(xx_frmConst.Range("Admin_WorkBook"), _
                                                Application.UserName) = 1 Then UserAdmin = True
        'Setup Form if User is Designer
        If Application.WorksheetFunction.CountIf(xx_frmConst.Range("Designer_WorkSheet"), _
                                                Application.UserName) = 1 Then UserDesigner = True
    '------------------------------------------------------------------------------------------------------------------------
            If .Tag = "New" Then 'Form Mode = New Add new Data
                .lbl_FormMode.Caption = "New request/ Neue Anforderung"
                .lbl_FormMode.Top = HeaderPlaceTop_H1
                'Delete & Save Bubttons Controlled in Continue Button Event
                .top_lbl_C_UK.Visible = True
                .top_lbl_C_DE.Visible = True
                .tbx_Anforderer.value = Application.UserName
                .tbx_Anforderer_am.value = CDate(Now)
            End If
    '------------------------------------------------------------------------------------------------------------------------
            If .Tag = "Edit" Then 'Form Mode = Edit data as User
                .lbl_FormMode.Caption = "Edit request/ Anforderung Ändern"
                .lbl_FormMode.Top = ButtonsTop
                .top_lbl_C_UK.Visible = False
                .top_lbl_C_DE.Visible = False
                .top_lbl_D_Delete.Visible = True
                .top_lbl_D_Save.Visible = True
                .tbx_Geaendert.value = Application.UserName
                .tbx_Geaendert_am.value = CDate(Now)
                If .chb_Reaktivieren.value = True Then Call chb_Reaktivieren_Click
            End If
    '------------------------------------------------------------------------------------------------------------------------
            If .Tag = "Approve" Or .Tag = "View" Then 'Form Mode = Approve Data as Designer or Admin Or just View Data
                If .Tag = "Approve" Then 'Form Mode = Approve Data as Designer or Admin
                    .lbl_FormMode.Caption = "Approve request/ Anforderung Freigeben <<Approved/ Freigeben>>"
                    If .chb_Approved.value = False Then .lbl_FormMode.Caption = "Approve request/ Anforderung Freigeben"
                    .lbl_FormMode.Top = ButtonsTop
                    .top_lbl_C_UK.Visible = False
                    .top_lbl_C_DE.Visible = False
                    .lbl_Informationen.Visible = False
                    .mpgControlPanel.Visible = UserDesigner
                    .tbx_Freigegeben.value = Application.UserName
                    .tbx_Freigegeben_am.value = CDate(Now)
                    If InStr(1, .cbx_Kontengruppe.value, "KUNA") <> 0 Then
                    .tbx_NewKUNA_SAPNr.Tag = "MustFill"
                    .tbx_NewKUNW_SAPNr.Visible = .tbx_WE_Name1.Visible
                    .lbl_NewKUNW_SAPNr.Visible = .lbl_WE_Namen.Visible
                    If .tbx_WE_Name1.Visible Then .tbx_NewKUNW_SAPNr.Tag = "MustFill"
                    .tbx_NewKUNR_SAPNr.Visible = .tbx_RE_Name1.Visible
                    .lbl_NewKUNR_SAPNr.Visible = .lbl_RE_Namen.Visible
                    If .tbx_RE_Name1.Visible Then .tbx_NewKUNR_SAPNr.Tag = "MustFill"
                        .top_lbl_Warenempfaenger_Abbruch.Visible = False
                        .top_lbl_Rechnungsempfaenger_Abbruch.Visible = False
                    Else
                        'WE SAPNr mustfill if ther are data filled in
                        If InStr(1, .cbx_Kontengruppe.value, "KUNW") <> 0 Then
                        .tbx_NewKUNW_SAPNr.Tag = "MustFill"
                        .lbl_NewKUNA_SAPNr.Visible = False
                        .lbl_NewKUNW_SAPNr.Top = 18
                        .tbx_NewKUNA_SAPNr.Visible = False
                        .tbx_NewKUNW_SAPNr.Top = 18
                        .lbl_NewKUNR_SAPNr.Visible = False
                        .tbx_NewKUNR_SAPNr.Visible = False
                        End If
                    End If
                    If .chb_Reaktivieren.value = True Then Call chb_Reaktivieren_Click
                    If .chb_Approved.value = True Then .chb_Approved.Tag = "Counted"
                End If
                '-------------------------------------------------------------------------------------------------------------
                    If .Tag = "View" Then 'Form Mode = just View Data
                        .lbl_FormMode.Caption = "View request/ Anforderung Anzeigen <<Locked/ Gesperrt>>"
                        .lbl_FormMode.Top = ButtonsTop
                        .top_lbl_C_UK.Visible = False
                        .top_lbl_C_DE.Visible = False
                        .top_lbl_D_Delete.Visible = False
                        .top_lbl_D_Save.Visible = False
                        .lbl_Informationen.Visible = False
                        .top_lbl_Warenempfaenger_Abbruch.Visible = False
                        .top_lbl_Rechnungsempfaenger_Abbruch.Visible = False
                        If .chb_Reaktivieren.value = True Then Call chb_Reaktivieren_Click
                    End If
                '-------------------------------------------------------------------------------------------------------------
            End If
    '-------------------------------------------------------------------------------------------------------------------------
        If InStr(1, .cbx_Kontengruppe.value, "(") <> 0 Then
            For Each ctrl In .mpg_Einstieg(0).Controls
                ctrl.Visible = False
            Next ctrl
            .lbl_Einstieg.Visible = True
            .SeparatorEinstieg_P0_1.Visible = True
            .lbl_Kontengruppe.Visible = True
            .cbx_Kontengruppe.Visible = True
        Else
            For Each ctrl In .mpg_Einstieg(0).Controls
                ctrl.Visible = True
            Next ctrl
        End If
        .chb_Testmode_Ein_Aus.Visible = UserAdmin 'TestMode Control CheckBox
    End With
    If Not Me.Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me
    If Me.Tag = "Approve" Then Application.Run "Form_TextSAP.RemarksToDesigner", Me
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'RESIZE EVENTS - UserForm
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub UserForm_Resize()

    If firstInit = True Then Exit Sub
    cFormResizing.ResizeControls
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'LAYOUT EVENTS - mpgControlPanel
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mpgControlPanel_Layout(ByVal index As Long)

    If firstInit = True Then Exit Sub
    cFormResizing.ResizeControls
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'LAYOUT EVENTS - mpg_Einstieg ///////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mpg_Einstieg_Layout(ByVal index As Long)

    If firstInit = True Then Exit Sub
    cFormResizing.ResizeControls
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'LAYOUT EVENTS - mpg_AllgemeineDaten ////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mpg_AllgemeineDaten_Layout(ByVal index As Long)

    If firstInit = True Then Exit Sub
    cFormResizing.ResizeControls
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'LAYOUT EVENTS - mpg_Vertriebsbereichsdaten /////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub mpg_Vertriebsbereichsdaten_Layout(ByVal index As Long)

    If firstInit = True Then Exit Sub
    cFormResizing.ResizeControls
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_C_UK
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_C_UK_Click()

    SwichFeed = "Translate"
    UserLanguage = "Caption_[EN]" ''default
    UserContTipp = "ControlTipText_[EN]" ''default
    UserRowSourceLang = "RowSourceAll_[EN]" ''default
    Application.Run "Form_Rebuild.Translate", Me, 1, SwichFeed, UserLanguage, _
                    UserContTipp, UserRowSourceLang, "empty", "empty", "empty", UserTabIndex, True
    If Me.cbx_Kontengruppe.value = "" Then Me.cbx_Kontengruppe.value = "( Please select )"
    Me.cbx_Verkaufsorganisation.Enabled = True
    Me.cbx_Vertriebsweg.Enabled = True
    Me.tbx_UserLangVal.value = 1
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_C_DE ////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_C_DE_Click()

    SwichFeed = "Translate"
    UserLanguage = "Caption_[DE]" '
    UserContTipp = "ControlTipText_[DE]" '
    UserRowSourceLang = "RowSourceAll_[DE]" '
    Application.Run "Form_Rebuild.Translate", Me, 49, SwichFeed, UserLanguage, _
                    UserContTipp, UserRowSourceLang, "empty", "empty", "empty", UserTabIndex, True
    If Me.cbx_Kontengruppe.value = "" Then Me.cbx_Kontengruppe.value = "( bitte auswählen )"
    Me.cbx_Verkaufsorganisation.Enabled = True
    Me.cbx_Vertriebsweg.Enabled = True
    Me.tbx_UserLangVal.value = 49
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - chb_Testmode_Ein_Aus ////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub chb_Testmode_Ein_Aus_Click()

    Application.ScreenUpdating = False
    WorkBookForm = ThisWorkbook.Name
    Set DebtorFile = Workbooks(WorkBookForm)
    DebtorFile.Unprotect passWordLock 'Unprotect Workbook
    Select Case Me.chb_Testmode_Ein_Aus.value
        Case Is = True
            With DebtorFile.Worksheets 'Visible
                .Item(1).Visible = -1
                .Item(2).Visible = -1
                .Item(3).Visible = -1
                .Item(4).Visible = -1
                .Item(5).Visible = -1
                .Item(6).Visible = -1
                .Item(9).Visible = -1
                .Item(10).Visible = -1
            End With
            With Application
'                .ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",true)"
                .DisplayFormulaBar = Me.chb_Testmode_Ein_Aus.value
                .DisplayStatusBar = Me.chb_Testmode_Ein_Aus.value
            End With
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
            With Application
'                .ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",false)"
                .DisplayFormulaBar = Me.chb_Testmode_Ein_Aus.value
                .DisplayStatusBar = Me.chb_Testmode_Ein_Aus.value
            End With
    End Select
    If UserDesigner Then DebtorFile.Worksheets.Item(9).Visible = -1
    DebtorFile.Protect passWordLock, True, True 'Protect Workbook
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_H3_Adresse //////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_H3_Adresse_Click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_AllgemeineDaten, Me.mpg_AllgemeineDaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_AllgemeineDaten.value = 0
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_H3_1_Steuerung_Zusatzdaten //////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_H3_1_Steuerung_Zusatzdaten_Click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_AllgemeineDaten, Me.mpg_AllgemeineDaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_AllgemeineDaten.value = 1
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_H3_2_Ansprechpartner ////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_H3_2_Ansprechpartner_Click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_AllgemeineDaten, Me.mpg_AllgemeineDaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_AllgemeineDaten.value = 2
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_H3_3_Warenempfaenger ////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_H3_3_Warenempfaenger_Click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_AllgemeineDaten, Me.mpg_AllgemeineDaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_AllgemeineDaten.value = 3
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_H3_4_Rechnungsempfaenger ////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_H3_4_Rechnungsempfaenger_Click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_AllgemeineDaten, Me.mpg_AllgemeineDaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_AllgemeineDaten.value = 4
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_H3_Verkauf //////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_H3_Verkauf_click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_Vertriebsbereichsdaten, Me.mpg_Vertriebsbereichsdaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_Vertriebsbereichsdaten.value = 0
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_H3_1_Versand ////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_H3_1_Versand_click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_Vertriebsbereichsdaten, Me.mpg_Vertriebsbereichsdaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_Vertriebsbereichsdaten.value = 1
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_H3_2_Faktura ////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_H3_2_Faktura_click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_Vertriebsbereichsdaten, Me.mpg_Vertriebsbereichsdaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_Vertriebsbereichsdaten.value = 2
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_H3_3_Partnerrolle_Zusatzdaten ///////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_H3_3_Partnerrolle_Zusatzdaten_click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_Vertriebsbereichsdaten, Me.mpg_Vertriebsbereichsdaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_Vertriebsbereichsdaten.value = 3
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_B_Continue //////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_B_Continue_Click()

    With Me
        If Application.Run("Form_Validate.ValidationFailed", _
                            Me.mpg_Einstieg, Me.mpg_Einstieg.value, Me.tbx_UserLangVal) Then Exit Sub
        If Not .Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me, Me.mpg_AllgemeineDaten
        If InStr(1, .cbx_Kontengruppe.value, "KUNW") <> 0 Then 'Kontengruppe is KUNW - Warenempfänger/Rechnungsempfänger
            If .Tag = "New" Then .top_lbl_D_Delete.Visible = True
            If .Tag = "New" Then .top_lbl_D_Save.Visible = True
            .mpg_Einstieg.Visible = False
            .mpg_AllgemeineDaten.Visible = True
            .mpg_AllgemeineDaten.value = 3
             top_lbl_B1_Warenempfaenger_click 'Call click event
            .top_lbl_H3_3_Warenempfaenger.Top = HeaderPlaceTop_H3
            .top_lbl_Warenempfaenger_Abbruch.Visible = False
        Else  'Kontengruppe is KUNA - Auftraggeber/Regulierer
            If .Tag = "New" Then .top_lbl_D_Delete.Visible = True
            If .Tag = "New" Then .top_lbl_D_Save.Visible = True
            .mpg_Einstieg.Visible = False
            .mpg_AllgemeineDaten.Visible = True
            .top_lbl_H2_Vertriebsbereichsdaten.Visible = True
            .top_lbl_H3_Adresse.Visible = True
            .top_lbl_H3_1_Steuerung_Zusatzdaten.Visible = True
            .top_lbl_H3_2_Ansprechpartner.Visible = True
            .top_lbl_H3_3_Warenempfaenger.Visible = True
            .top_lbl_H3_4_Rechnungsempfaenger.Visible = True
        End If
        .top_lbl_go_Einstieg.Visible = True
        .top_lbl_C_UK.Visible = False
        .top_lbl_C_DE.Visible = False
    End With
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_B_AllgDaten_Back ////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_B_AllgDaten_Back_Click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_AllgemeineDaten, Me.mpg_AllgemeineDaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_AllgemeineDaten.value = Me.mpg_AllgemeineDaten.value - 1
    If Not Me.Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me, Me.mpg_AllgemeineDaten
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_B_AllgDaten_Next ////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_B_AllgDaten_Next_Click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_AllgemeineDaten, Me.mpg_AllgemeineDaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_AllgemeineDaten.value = Me.mpg_AllgemeineDaten.value + 1
    If Not Me.Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me, Me.mpg_AllgemeineDaten
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_B_VertDaten_Back ////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_B_VertDaten_Back_Click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_Vertriebsbereichsdaten, Me.mpg_Vertriebsbereichsdaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_Vertriebsbereichsdaten.value = Me.mpg_Vertriebsbereichsdaten.value - 1
    If Not Me.Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me, Me.mpg_Vertriebsbereichsdaten
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_B_VertDaten_Next ////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_B_VertDaten_Next_Click()

    If Application.Run("Form_Validate.ValidationFailed", _
                        Me.mpg_Vertriebsbereichsdaten, Me.mpg_Vertriebsbereichsdaten.value, Me.tbx_UserLangVal) Then Exit Sub
    Me.mpg_Vertriebsbereichsdaten.value = Me.mpg_Vertriebsbereichsdaten.value + 1
    If Not Me.Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me, Me.mpg_Vertriebsbereichsdaten
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_D_Save //////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_D_Save_Click()

    Dim MPages As MSForms.Control, xPage As MSForms.Page, X As Integer
    Application.ScreenUpdating = False
    If Not Me.Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me 'Rebuild User Form Layout
    If Me.mpg_Einstieg.Visible = True Then Call top_lbl_B_Continue_Click 'Skipp Einstieg Page
    If InStr(1, Me.cbx_Kontengruppe.value, "KUNW") <> 0 Then
        'Validate AllgemeineDaten
        If Application.Run("Form_Validate.ValidationFailed", _
                            Me.mpg_AllgemeineDaten, Me.mpg_AllgemeineDaten.value, Me.tbx_UserLangVal) Then Exit Sub
        'Validate ControlPanel
        If Application.Run("Form_Validate.ValidationFailed", _
                            Me.mpgControlPanel, Me.mpgControlPanel.value, Me.tbx_UserLangVal) Then Exit Sub
    Else
        For Each MPages In Me.Controls 'Validate all MultiPages
            If TypeOf MPages Is MSForms.MultiPage Then
                For X = 0 To MPages.Pages.Count - 1
                    ' Validate all Pages in MultiPage
                    If Application.Run("Form_Validate.ValidationFailed", _
                                        MPages, X, Me.tbx_UserLangVal, MPages.Name) Then Exit Sub
                Next X
            End If
        Next
    End If
    Application.Run "Form_Save.Save", Me, Me.tbx_SelectedRowVal.value, Me.chb_Approved.value 'Save Data
    Unload Me 'Destroy Form from Memory
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_Approve /////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_Approve_Click()
    
    If Me.chb_Approved.value = True Then
        If Me.chb_E_Rechnung.value Or Me.chb_WE_E_Lieferschein.value Or Me.chb_RE_E_Rechnung.value Then
            Application.Run "Form_Email.EmialToEDocumentsSetter", Me 'Create Email for E-Doc setter
        End If
    End If
    Call top_lbl_D_Save_Click 'Clone Save Button with approved role
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_D_Delete ////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_D_Delete_Click()

    Dim UserAnswer As Long, MsgText As String, MsgTitel As String
    Select Case Me.tbx_UserLangVal.value
        Case Is = 49 'German
            MsgText = "Wollen Sie wirklich Löschen?"
            MsgTitel = "Vom gewählten Debitor werden alle Daten gelöscht!"
        Case Else 'English
            MsgText = "Do you really want to delete?"
            MsgTitel = "Entered values from the chosen customer will be deleted!"
    End Select
    UserAnswer = MsgBox(MsgText, vbQuestion + vbYesNo + vbDefaultButton2, MsgTitel) 'Ask user if he wants to delete data
    If UserAnswer = vbNo Then Exit Sub 'User clicked cancel button
    If UserAnswer = vbYes Then
        Me.tbx_Geaendert.value = Application.UserName
        Me.tbx_Geaendert_am.value = CDate(Now)
        'If Data exists then Move existing data to Deleted Sheet
        Application.Run "Form_Delete.Delete", Me, Me.tbx_SelectedRowVal.value
    End If
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_B1_Warenempfaenger //////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_B1_Warenempfaenger_click()

    Me.top_lbl_B1_Warenempfaenger.Visible = False
    For Each ctrl In Me.mpg_AllgemeineDaten(3).Controls
        ctrl.Visible = True
    Next ctrl
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - chb_WE_E_Lieferschein  //////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub chb_WE_E_Lieferschein_Click()

    If Me.chb_WE_E_Lieferschein.value = False Then _
            Me.tbx_WE_Email.Tag = "": Me.tbx_WE_Email.BackColor = RGB(255, 255, 255) 'White
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_Warenempfaenger_Abbruch /////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_Warenempfaenger_Abbruch_click()

    With Me
        For Each ctrl In .mpg_AllgemeineDaten(3).Controls
            If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Then ctrl.Object.value = ""
            If TypeOf ctrl Is MSForms.CheckBox Then ctrl.Object.value = False
            ctrl.Visible = False
            ctrl.Object.BackColor = RGB(255, 255, 255) 'White
        Next ctrl
            .top_lbl_B1_Warenempfaenger.Visible = True
            .top_lbl_Warenempfaenger_Abbruch.BackColor = RGB(0, 112, 192) 'Blue
    End With
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_B1_Rechnungsempfaenger //////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_B1_Rechnungsempfaenger_click()

    Me.top_lbl_B1_Rechnungsempfaenger.Visible = False
    For Each ctrl In Me.mpg_AllgemeineDaten(4).Controls
        ctrl.Visible = True
    Next ctrl
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - chb_RE_E_Rechnung ///////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub chb_RE_E_Rechnung_Click()

    If Me.chb_RE_E_Rechnung.value = False Then _
            Me.tbx_RE_Email.Tag = "": Me.tbx_RE_Email.BackColor = RGB(255, 255, 255) 'White
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - top_lbl_Rechnungsempfaenger_Abbruch /////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub top_lbl_Rechnungsempfaenger_Abbruch_click()

    With Me
        For Each ctrl In .mpg_AllgemeineDaten(4).Controls
            If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Then ctrl.Object.value = ""
            If TypeOf ctrl Is MSForms.CheckBox Then ctrl.Object.value = False
            ctrl.Visible = False
            ctrl.Object.BackColor = RGB(255, 255, 255) 'White
        Next ctrl
            .top_lbl_B1_Rechnungsempfaenger.Visible = True
            .top_lbl_Rechnungsempfaenger_Abbruch.BackColor = RGB(0, 112, 192) 'Blue
    End With
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLICK EVENTS - chb_Approved ////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub chb_Approved_Click()

    If Not Me.Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me 'Validation Setup
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CHANGE EVENTS - cbx_Verkaufsorganisation
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cbx_Kontengruppe_Change()

    With Me
        If InStr(1, .cbx_Kontengruppe.value, "(") <> 1 Then
            .top_lbl_B_Continue.Visible = True
            For Each ctrl In .mpg_Einstieg(0).Controls
                ctrl.Visible = True
            Next ctrl
            .chb_Testmode_Ein_Aus.Visible = UserAdmin
            UserLang = Me.tbx_UserLangVal.value 'Change Feed Data
            SwichFeed = "Kontengruppe"
            UChoise = Me.tbx_Boolean.value
            If UserLang = 49 Then UserRowSourceLang = "RowSourceAll_[DE]" ''German
            If UserLang = 1 Then UserRowSourceLang = "RowSourceAll_[EN]"   ''default
            Application.Run "Form_Rebuild.Translate", Me, 1, SwichFeed, "empty", "empty", UserRowSourceLang, _
                            "empty", "empty", "empty", "TabIndex", UChoise
            .cbx_Verkaufsorganisation.Tag = "MustFill"
            .cbx_Vertriebsweg.Tag = "MustFill"
            .cbx_Kontengruppe.Enabled = False
            .cbx_Sparte = "10 - Interlining"
            .cbx_Sparte.Enabled = False
        Else
            For Each ctrl In .mpg_Einstieg(0).Controls
                ctrl.Visible = False
            Next ctrl
            .lbl_Einstieg.Visible = True
            .SeparatorEinstieg_P0_1.Visible = True
            .lbl_Kontengruppe.Visible = True
            .cbx_Kontengruppe.Visible = True
            .cbx_Kontengruppe.Enabled = True
            .top_lbl_B_Continue.Visible = False
        End If
        If Not .Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me
    End With
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CHANGE EVENTS - cbx_Verkaufsorganisation ///////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cbx_Verkaufsorganisation_Change()

    Dim indexValue As String
    UserTabIndex = "TabIndex" ''Fix Values
    indexValue = Me.cbx_Verkaufsorganisation.value
    Select Case indexValue
        Case Is = "0361 - Interlining ES"
            UserLang = 0 'Change Feed Data
            SwichFeed = "VerkOrg"
            VerkOrgMustFill = "MustFieldTag_[0361]"
            VerkOrgRowSource = "RowSource_[0361]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", "empty", _
                            VerkOrgMustFill, VerkOrgRowSource, "empty", UserTabIndex, False
            Me.cbx_Vertriebsweg.value = ""
            Me.cbx_Verkaufsorganisation.Enabled = False
        Case Is = "2561 - Interlining TR"
            UserLang = 0 'Change Feed Data
            SwichFeed = "VerkOrg"
            VerkOrgMustFill = "MustFieldTag_[2561]"
            VerkOrgRowSource = "RowSource_[2561]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", "empty", _
            VerkOrgMustFill, VerkOrgRowSource, "empty", UserTabIndex, False
            Me.cbx_Vertriebsweg.value = ""
            Me.cbx_Verkaufsorganisation.Enabled = False
        Case Is = "2961 - Interlining DE"
            UserLang = 0 'Change Feed Data
            SwichFeed = "VerkOrg"
            VerkOrgMustFill = "MustFieldTag_[2961]"
            VerkOrgRowSource = "RowSource_[2961]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", "empty", _
            VerkOrgMustFill, VerkOrgRowSource, "empty", UserTabIndex, False
            Me.cbx_Vertriebsweg.value = ""
            Me.cbx_Verkaufsorganisation.Enabled = False
            Me.lbl_Rechtsform.Visible = True
            Me.tbx_Rechtsform.Visible = True
            Me.chb_Rechtsform.Visible = True
            Me.lbl_WebShop.Visible = True
            Me.cbx_WebShop.Visible = True
        Case Is = "3661 - FPM Apparel IT"
            UserLang = 0 'Change Feed Data
            SwichFeed = "VerkOrg"
            VerkOrgMustFill = "MustFieldTag_[3661]"
            VerkOrgRowSource = "RowSource_[3661]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", "empty", _
            VerkOrgMustFill, VerkOrgRowSource, "empty", UserTabIndex, False
            Me.cbx_Vertriebsweg.value = ""
            Me.cbx_Verkaufsorganisation.Enabled = False
        Case Else
            Me.cbx_Verkaufsorganisation.Enabled = False
    End Select
    If Me.cbx_Verkaufsorganisation <> "" Then Me.cbx_Verkaufsorganisation.BackColor = RGB(255, 255, 255) 'White
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CHANGE EVENTS - cbx_Vertriebsweg ///////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cbx_Vertriebsweg_Change()

    Dim indexValue As String
    indexValue = Me.cbx_Vertriebsweg.value
    Select Case indexValue
        Case Is = "01 - Industrie" 'for Verk.Org 2961 - Interlining DE
            UserLang = 0 'Change Feed Data
            SwichFeed = "VertWeg"
            VertWegRowSource = "RowSource_[2961_01]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", _
            "empty", "empty", "empty", VertWegRowSource, "TabIndex", False
            Me.cbx_Vertriebsweg.Enabled = False
        Case Is = "01 - Industry" 'for Verk.Org 3661 - FPM Apparel IT
            Me.cbx_Preisliste.Tag = "NotFill"
            Me.cbx_Vertriebsweg.Enabled = False
        Case Is = "BO - Bochum" 'for Verk.Org 2961 - Interlining DE
            UserLang = 0 'Change Feed Data
            SwichFeed = "VertWeg"
            VertWegRowSource = "RowSource_[2961_BO]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", _
            "empty", "empty", "empty", VertWegRowSource, "TabIndex", False
            Me.cbx_Vertriebsweg.Enabled = False
        Case Is = "GY - Gygli" 'for Verk.Org 2961 - Interlining DE
            UserLang = 0 'Change Feed Data
            SwichFeed = "VertWeg"
            VertWegRowSource = "RowSource_[2961_GY]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", _
            "empty", "empty", "empty", VertWegRowSource, "TabIndex", False
            Me.cbx_Vertriebsweg.Enabled = False
        Case Is = "HD - Heidelberg" 'for Verk.Org 2961 - Interlining DE
            UserLang = 0 'Change Feed Data
            SwichFeed = "VertWeg"
            VertWegRowSource = "RowSource_[2961_HD]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", _
            "empty", "empty", "empty", VertWegRowSource, "TabIndex", False
            Me.cbx_Vertriebsweg.Enabled = False
        Case Is = "IF - IL France" 'for Verk.Org 2961 - Interlining DE
            UserLang = 0 'Change Feed Data
            SwichFeed = "VertWeg"
            VertWegRowSource = "RowSource_[2961_IF]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", _
            "empty", "empty", "empty", VertWegRowSource, "TabIndex", False
            Me.cbx_Vertriebsweg.Enabled = False
        Case Is = "IU - IL UK" 'for Verk.Org 2961 - Interlining DE
            UserLang = 0 'Change Feed Data
            SwichFeed = "VertWeg"
            VertWegRowSource = "RowSource_[2961_IU]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", _
            "empty", "empty", "empty", VertWegRowSource, "TabIndex", False
            Me.cbx_Vertriebsweg.Enabled = False
        Case Is = "NA - IL N-Afrika" 'for Verk.Org 2961 - Interlining DE
            UserLang = 0 'Change Feed Data
            SwichFeed = "VertWeg"
            VertWegRowSource = "RowSource_[2961_NA]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", _
            "empty", "empty", "empty", VertWegRowSource, "TabIndex", False
            Me.cbx_Vertriebsweg.Enabled = False
        Case Is = "PL - IL Polen" 'for Verk.Org 2961 - Interlining DE
            UserLang = 0 'Change Feed Data
            SwichFeed = "VertWeg"
            VertWegRowSource = "RowSource_[2961_PL]"
            Application.Run "Form_Rebuild.Translate", Me, UserLang, SwichFeed, "empty", "empty", _
            "empty", "empty", "empty", VertWegRowSource, "TabIndex", False
            Me.cbx_Vertriebsweg.Enabled = False
        Case Else 'for Other
            Me.cbx_Vertriebsweg.Enabled = False
    End Select
    If Me.Tag Like "New" And Me.cbx_Kontengruppe.value Like "*KUNA*" Then Application.Run "Form_CustomSetup.CustomSettings", Me
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CHANGE EVENTS - chb_Reaktivieren ///////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub chb_Reaktivieren_Click()

    Dim X As Integer
    If chb_Reaktivieren.value = True Then
        For X = 0 To 1
            For Each ctrl In Me.mpg_AllgemeineDaten(X).Controls
                If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Or TypeOf ctrl Is MSForms.CheckBox Then
                    If ctrl.Tag = "MustFill" Then ctrl.Tag = "NotFill": ctrl.Object.BackColor = RGB(255, 255, 255) 'White
                    If Me.cbx_Kontengruppe.value Like "*KUNA*" Then Me.tbx_Name1.Tag = "MustFill"
                End If
            Next
        Next X
    Else
        For X = 0 To 1
            For Each ctrl In Me.mpg_AllgemeineDaten(X).Controls
                If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Or TypeOf ctrl Is MSForms.CheckBox Then
                    If ctrl.Tag = "NotFill" Then ctrl.Tag = "MustFill"
                End If
            Next
        Next X
    End If
    If Not Me.Tag = "Initialize" Then Application.Run "Form_Rebuild.UpdateFormControls", Me
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CHANGE EVENTS - cbx_Land_Change ////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub chb_Rechtsform_Click()

    If Me.chb_Rechtsform.value = True Then
        Me.tbx_Rechtsform.Tag = "NotFill": Me.tbx_Rechtsform.BackColor = RGB(255, 255, 255) 'White
    Else
        If Me.tbx_Rechtsform.value = "" Then Me.tbx_Rechtsform.Tag = "MustFill": Me.tbx_Rechtsform.BackColor = RGB(255, 255, 204) 'Yellow
    End If
End Sub
Private Sub cbx_Land_Change()

    Application.Run "Form_CustomSetup.CustomSettings", Me
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CHANGE EVENTS - cbx_Verkeaufergruppe_Change ////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cbx_Verkeaufergruppe_Change()

    Application.Run "Form_CustomSetup.CustomSettings", Me
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CHANGE EVENTS - cbx_Zahlungsbedingung //////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cbx_Zahlungsbedingung_Change()

    On Error Resume Next
    With Me ' Displays description if Combobox Zahlungskondition Ithem is selected
        .lbx_Zahlungsbedingung.Clear
        .lbx_Zahlungsbedingung.AddItem Me.cbx_Zahlungsbedingung.Column(1)
        .lbx_Zahlungsbedingung.AddItem Me.cbx_Zahlungsbedingung.Column(2)
        .lbx_Zahlungsbedingung.AddItem Me.cbx_Zahlungsbedingung.Column(3)
    End With
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'EXIT EVENTS - tbx_Ort_Exit
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''Private Sub tbx_Ort_Exit(ByVal Cancel As MSForms.ReturnBoolean)
''
''    If Me.Tag = "New" Then Me.tbx_Incoterms_Ort.value = Me.tbx_Ort.value
''End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CLOSE USERFORM EVENTS - UserForm_QueryClose
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Dim UserAnswer As Long, MsgText As String, MsgTitel As String
    If Me.Tag = "Edit" And Me.lbl_FormMode.Caption Like "*<<*" Then Cancel = False: Exit Sub 'Do nothing
    If Me.Tag = "View" Or Me.Tag = "Approve" Then Cancel = False: Exit Sub 'Close the userform
        If CloseMode = vbFormControlMenu Then 'Ask User if to Quit?
            Cancel = True
        Select Case Me.tbx_UserLangVal.value
            Case Is = 49 'German
                MsgText = "Wollen Sie wirklich abbrechen?"
                MsgTitel = "Die Daten worden noch nicht gespeichert!"
            Case Else 'English
                MsgText = "Do you really want to cancel?"
                MsgTitel = "Entered data was not saved!"
        End Select
        'Ask user if he wants to save before exiting
        UserAnswer = MsgBox(MsgText, vbQuestion + vbYesNo + vbDefaultButton2, MsgTitel)
        If UserAnswer = vbNo Then Exit Sub 'Do nothing'User clicked cancel button
        If UserAnswer = vbYes Then Cancel = False 'Close the userform
    End If
End Sub
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////// ******** PROGRAM FINISH ********* ////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'----------------------------------------------------------------------------------------------------------------------------
