VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContainerLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates basic layout logic for dynamic MSForms controls."
'@Folder AccountsManager.Infrastructure.View.Dynamic
'@ModuleDescription "Encapsulates basic layout logic for dynamic MSForms controls."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IContainerLayout

Private Type TState
    Parent As MSForms.Controls
    Direction As LayoutDirection
    NextControlTop As Long
    NextControlLeft As Long
    Margin As Long
End Type

Private Const DefaultMargin As Long = 3

Private this As TState

Public Function Create(ByVal Parent As MSForms.Controls, Optional ByVal Direction As LayoutDirection, Optional ByVal Top As Long = 0, Optional ByVal Left As Long = 0, Optional ByVal Margin As Long = DefaultMargin) As IContainerLayout
    Dim result As ContainerLayout
    Set result = New ContainerLayout
    Set result.Parent = Parent
    result.NextControlTop = Top + Margin
    result.NextControlLeft = Left + Margin
    result.Direction = Direction
    result.Margin = Margin
    Set Create = result
End Function

Public Property Get Parent() As MSForms.Controls
    Set Parent = this.Parent
End Property

Public Property Set Parent(ByVal RHS As MSForms.Controls)
    Set this.Parent = RHS
End Property

Public Property Get NextControlTop() As Long
    NextControlTop = this.NextControlTop
End Property

Public Property Let NextControlTop(ByVal RHS As Long)
    this.NextControlTop = RHS
End Property

Public Property Get NextControlLeft() As Long
    NextControlLeft = this.NextControlLeft
End Property

Public Property Let NextControlLeft(ByVal RHS As Long)
    this.NextControlLeft = RHS
End Property

Public Property Get Direction() As LayoutDirection
    Direction = this.Direction
End Property

Public Property Let Direction(ByVal RHS As LayoutDirection)
    this.Direction = RHS
End Property

Public Property Get Margin() As Long
    Margin = this.Margin
End Property

Public Property Let Margin(ByVal RHS As Long)
    this.Margin = RHS
End Property

Private Sub OffsetControlSize(ByVal Control As MSForms.Control, ByVal Margin As Long)
    Select Case this.Direction
    
        Case LayoutDirection.TopToBottom
            this.NextControlTop = Control.Top + Control.Height + Margin
        
        Case LayoutDirection.LeftToRight
            this.NextControlLeft = Control.Left + Control.Width + Margin
            
        Case Else
            Debug.Assert False ' not implemented...
            
    End Select
End Sub

Private Function IContainerLayout_Add(ByVal ProgID As String, Optional ByVal Height As Variant, Optional ByVal Width As Variant, Optional ByVal ApplyMargin As Boolean = True) As Object
    Dim result As Control
    Set result = this.Parent.add(ProgID)
    
    result.Top = this.NextControlTop
    result.Left = this.NextControlLeft
    
    If Not IsMissing(Height) Then
        result.Height = Height
    End If
    
    If Not IsMissing(Width) Then
        result.Width = Width
    End If
    
    OffsetControlSize result, IIf(ApplyMargin, this.Margin, 0)
    Set IContainerLayout_Add = result
    
End Function

Private Sub IContainerLayout_SkipPosition(Optional ByVal Margin As Long)
    Select Case this.Direction
        
        Case LayoutDirection.TopToBottom
            this.NextControlTop = this.NextControlTop + IIf(Margin > 0, Margin, this.Margin * 2)
            
        Case LayoutDirection.LeftToRight
            this.NextControlLeft = this.NextControlLeft + IIf(Margin > 0, Margin, this.Margin * 2)
            
    End Select
End Sub
