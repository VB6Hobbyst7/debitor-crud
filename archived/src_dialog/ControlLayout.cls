VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates basic layout logic for dynamic MSForms controls."
'@Folder("ValidateUserInput.Model.ResizeView")
'@ModuleDescription "Encapsulates basic layout logic for dynamic MSForms controls."
'@Exposed
Option Explicit

Public Enum AnchorEdges
    LeftAnchor = 1
    TopAnchor = 2
    RightAnchor = 4
    BottomAnchor = 8
    AnchorAll = LeftAnchor + TopAnchor + RightAnchor + BottomAnchor
End Enum

Private Type TControlLayout
    Anchors As AnchorEdges
    Margins As SideMargins
    FormSize As ControlSize
    Size As ControlSize
    BoundControl As Control
End Type

Private this As TControlLayout

Public Sub Bind(ByVal Form As MsForms.UserForm, ByVal formControl As Control, ByVal Anchor As AnchorEdges)

    Set this.FormSize = New ControlSize
    Set this.Size = New ControlSize
    Set this.Margins = New SideMargins
    this.Anchors = Anchor

    this.FormSize.height = Form.InsideHeight
    this.FormSize.width = Form.InsideWidth

    this.Size.height = formControl.height
    this.Size.width = formControl.width

    If (Anchor And BottomAnchor) = BottomAnchor Then
        this.Margins.BottomMargin = this.FormSize.height - formControl.top - formControl.height
    End If

    If (Anchor And LeftAnchor) = LeftAnchor Then
        this.Margins.LeftMargin = formControl.left
    End If

    If (Anchor And RightAnchor) = RightAnchor Then
        this.Margins.RightMargin = this.FormSize.width - formControl.left - formControl.width
    End If

    If (Anchor And TopAnchor) = TopAnchor Then
        this.Margins.TopMargin = formControl.top
    End If

    Set this.BoundControl = formControl

End Sub

Public Sub Resize(ByVal Form As UserForm)

    this.FormSize.height = Form.InsideHeight
    this.FormSize.width = Form.InsideWidth

    Layout this.BoundControl

End Sub

Private Sub Layout(ByVal formControl As Control)

    If (Anchors And TopAnchor) = TopAnchor Then

    If (Anchors And BottomAnchor) = BottomAnchor Then
        formControl.height = this.FormSize.height - formControl.top - this.Margins.BottomMargin
    End If

    ElseIf (Anchors And BottomAnchor) = BottomAnchor Then
        formControl.top = this.FormSize.height - formControl.height - this.Margins.BottomMargin
    End If


    If (Anchors And LeftAnchor) = LeftAnchor Then

        If (Anchors And RightAnchor) = RightAnchor Then
            formControl.width = this.FormSize.width - formControl.left - this.Margins.RightMargin
        End If

    ElseIf (Anchors And RightAnchor) = RightAnchor Then
        formControl.left = this.FormSize.width - formControl.width - this.Margins.RightMargin
    End If

End Sub

Public Property Get Anchors() As AnchorEdges
    Anchors = this.Anchors
End Property

Public Property Let Anchors(ByVal ancor As AnchorEdges)
    this.Anchors = ancor
End Property

Public Property Get BoundControl() As Control
    Set BoundControl = this.BoundControl
End Property

Public Property Set BoundControl(ByVal bound As Control)
    Set this.BoundControl = bound
End Property
