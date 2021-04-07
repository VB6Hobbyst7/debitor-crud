VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormResizing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: clsFormResizing **************************************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Types and Events ///////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Type ControlAnchorsAndValues
    ctl As MSForms.Control
    AnchorTop As Boolean
    AnchorLeft As Boolean
    AnchorBottom As Boolean
    AnchorRight As Boolean
    StartingTop As Double
    StartingLeft As Double
    StartingHeight As Double
    StartingWidth As Double
    ParentStartingHeight As Double
    ParentStartingWidth As Double
End Type
Private WithEvents m_UserForm As MSForms.UserForm
Attribute m_UserForm.VB_VarHelpID = -1
Private m_ControlsAnchorsAndVals() As ControlAnchorsAndValues
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Initialize /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub Initialize(UF As MSForms.UserForm, ControlsAndAnchors() As ControlAndAnchors)

    Dim i As Long
    Set m_UserForm = UF
    ReDim m_ControlsAnchorsAndVals(LBound(ControlsAndAnchors) To UBound(ControlsAndAnchors))
    For i = LBound(ControlsAndAnchors) To UBound(ControlsAndAnchors)
        With m_ControlsAnchorsAndVals(i)
            Set .ctl = ControlsAndAnchors(i).ctl
            .AnchorTop = ControlsAndAnchors(i).AnchorTop
            .AnchorLeft = ControlsAndAnchors(i).AnchorLeft
            .AnchorBottom = ControlsAndAnchors(i).AnchorBottom
            .AnchorRight = ControlsAndAnchors(i).AnchorRight
            .StartingTop = .ctl.Top
            .StartingLeft = .ctl.Left
            .StartingHeight = .ctl.Height
            .StartingWidth = .ctl.Width
            .ParentStartingHeight = .ctl.Parent.InsideHeight
            .ParentStartingWidth = .ctl.Parent.InsideWidth
        End With
    Next i
    MakeFormResizable m_UserForm, True
    ShowMinimizeButton m_UserForm, False
    ShowMaximizeButton m_UserForm, False
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'ResizeControls /////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub ResizeControls()

    Dim i As Long
    For i = LBound(m_ControlsAnchorsAndVals) To UBound(m_ControlsAnchorsAndVals)
        With m_ControlsAnchorsAndVals(i)
            If .AnchorTop And .AnchorBottom Then
                .ctl.Top = .StartingTop
                .ctl.Height = Application.WorksheetFunction.Max(0, .StartingHeight + _
                (.ctl.Parent.InsideHeight - .ParentStartingHeight))
            ElseIf .AnchorTop And Not .AnchorBottom Then
                .ctl.Top = .StartingTop
            ElseIf Not .AnchorTop And .AnchorBottom Then
                .ctl.Top = .StartingTop + (.ctl.Parent.InsideHeight - .ParentStartingHeight)
            End If
            If .AnchorLeft And .AnchorRight Then
                .ctl.Left = .StartingLeft
                .ctl.Width = Application.WorksheetFunction.Max(0, .StartingWidth + _
                (.ctl.Parent.InsideWidth - .ParentStartingWidth))
            ElseIf .AnchorLeft And Not .AnchorRight Then
                .ctl.Left = .StartingLeft
            ElseIf Not .AnchorLeft And .AnchorRight Then
                .ctl.Left = .StartingLeft + (.ctl.Parent.InsideWidth - .ParentStartingWidth)
            End If
        End With
    Next i
End Sub
