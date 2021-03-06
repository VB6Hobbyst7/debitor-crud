VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Builds dynamic MSForms UI components from a binding source."
'@Folder AccountsManager.Infrastructure.StringFormatting.View.Dynamic
'@ModuleDescription "Builds dynamic MSForms UI components from a binding source."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IDynamicControlBuilder

Private Type TState
    Context As IAppContext
    Container As IContainerLayout
End Type

Private this As TState

Public Function Create(ByVal Context As IAppContext, ByVal Container As IContainerLayout) As IDynamicControlBuilder
    Dim result As DynamicControls
    Set result = New DynamicControls
    Set result.Context = Context
    Set result.Container = Container
    Set Create = result
End Function

Public Property Get Container() As IContainerLayout
    Set Container = this.Container
End Property

Friend Property Set Container(ByVal RHS As IContainerLayout)
    Set this.Container = RHS
End Property

Public Property Get Context() As IAppContext
    Set Context = this.Context
End Property

Friend Property Set Context(ByVal RHS As IAppContext)
    Set this.Context = RHS
End Property

Private Sub BindCaptionSource(ByVal Source As Variant, ByVal Target As Object, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter)
    If VarType(Source) = vbString Then
        'late-bound
        Target.Caption = Source
        
    ElseIf TypeOf Source Is IBindingPath Then
        Dim SourcePath As IBindingPath
        Set SourcePath = Source
        this.Context.Bindings.BindPropertyPath SourcePath.Context, SourcePath.path, Target, _
            Converter:=Converter, _
            StringFormat:=FormatString
        
    Else
        GuardClauses.GuardExpression True, TypeName(Me), "Unsupported Source type."
    End If
End Sub

Private Function CreateLabel(ByVal SourceCaption As Variant, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter, Optional ByVal ApplyMargin As Boolean = True) As MSForms.Label
    Dim result As MSForms.Label
    Set result = this.Container.add(FormsProgID.LabelProgId, ApplyMargin:=ApplyMargin)
    result.AutoSize = True
    result.WordWrap = False
    BindCaptionSource SourceCaption, result, FormatString, Converter
    Set CreateLabel = result
End Function

Private Function ErrorAdornerOrDefault(ByVal Validator As IValueValidator, ByVal ErrorAdorner As IDynamicAdorner, ByVal Target As Object) As IDynamicAdorner
    Dim result As IDynamicAdorner
    If Not Validator Is Nothing Then
        If ErrorAdorner Is Nothing Then
            Set result = this.Context.Validation.AdornerFactory.Create(Target)
        Else
            Set result = ErrorAdorner
        End If
    End If
    Set ErrorAdornerOrDefault = result
End Function

Private Function IDynamicControlBuilder_CheckBoxFor(ByVal SourceValue As IBindingPath, ByVal SourceCaption As Variant) As MSForms.CheckBox
    
    Dim result As MSForms.CheckBox
    Set result = this.Container.add(FormsProgID.CheckBoxProgId)
    
    BindCaptionSource SourceCaption, result
    this.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.path, result
    
    Set IDynamicControlBuilder_CheckBoxFor = result
    
End Function

Private Function IDynamicControlBuilder_ComboBoxFor(ByVal SourceValue As IBindingPath, ByVal SourceItems As IBindingPath, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter, Optional ByVal Validator As IValueValidator, Optional ByVal ErrorAdorner As IDynamicAdorner, Optional ByVal TitleSource As Variant) As MSForms.ComboBox
    
    If Not IsEmpty(TitleSource) Then
        CreateLabel TitleSource, ApplyMargin:=False
    End If
    
    Dim result As MSForms.ComboBox
    Set result = this.Container.add(FormsProgID.ComboBoxProgId)
        
    this.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.path, result, _
        StringFormat:=FormatString, _
        Converter:=Converter, _
        Validator:=Validator, _
        ValidationAdorner:=ErrorAdornerOrDefault(Validator, ErrorAdorner, result)

    this.Context.Bindings.BindPropertyPath SourceItems.Context, SourceItems.path, result
    
    Set IDynamicControlBuilder_ComboBoxFor = result
    
End Function

Private Function IDynamicControlBuilder_CommandButtonFor(ByVal command As ICommand, ByVal BindingContext As Object, ByVal SourceCaption As Variant) As MSForms.CommandButton
    
    Dim result As MSForms.CommandButton
    Set result = this.Container.add(FormsProgID.CommandButtonProgId)
    
    Dim Button As MSForms.Control
    Set Button = result
    
    If TypeOf command Is CancelCommand Then
        Button.Cancel = True
    ElseIf TypeOf command Is AcceptCommand Then
        Button.Default = True
    End If
    
    BindCaptionSource SourceCaption, result
    this.Context.Commands.BindCommand BindingContext, result, command
    
    Set IDynamicControlBuilder_CommandButtonFor = result
    
End Function

Private Function IDynamicControlBuilder_LabelFor(ByVal SourceCaption As Variant, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter) As MSForms.Label
    Set IDynamicControlBuilder_LabelFor = CreateLabel(SourceCaption, FormatString, Converter)
End Function

Private Function IDynamicControlBuilder_ListBoxFor(ByVal SourceValue As IBindingPath, ByVal SourceItems As IBindingPath, Optional ByVal TitleSource As Variant) As MSForms.ListBox
    
    If Not IsEmpty(TitleSource) Then
        CreateLabel TitleSource, ApplyMargin:=False
    End If
    
    Dim result As MSForms.ListBox
    Set result = this.Container.add(FormsProgID.ListBoxProgId)
    
    this.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.path, result
    this.Context.Bindings.BindPropertyPath SourceItems.Context, SourceItems.path, result
    
    Set IDynamicControlBuilder_ListBoxFor = result
    
End Function

Private Function IDynamicControlBuilder_OptionButtonFor(ByVal SourceValue As IBindingPath, ByVal SourceCaption As Variant) As MSForms.OptionButton
    
    Dim result As MSForms.OptionButton
    Set result = this.Container.add(FormsProgID.OptionButtonProgId)
    
    BindCaptionSource SourceCaption, result
    this.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.path, result
    
    Set IDynamicControlBuilder_OptionButtonFor = result
    
End Function

Private Function IDynamicControlBuilder_TextAreaFor(ByVal SourceValue As IBindingPath, Optional ByVal Converter As IValueConverter, Optional ByVal Validator As IValueValidator, Optional ByVal ErrorAdorner As IDynamicAdorner, Optional ByVal TitleSource As Variant) As MSForms.TextBox
    
    If Not IsEmpty(TitleSource) Then
        CreateLabel TitleSource, ApplyMargin:=False
    End If
    
    Dim result As MSForms.TextBox
    Set result = IDynamicControlBuilder_TextBoxFor(SourceValue, Converter:=Converter, Validator:=Validator, ErrorAdorner:=ErrorAdorner)
    
    result.MultiLine = True
    result.WordWrap = True
    result.ScrollBars = fmScrollBarsVertical
    
    Set IDynamicControlBuilder_TextAreaFor = result
    
End Function

Private Function IDynamicControlBuilder_TextBoxFor(ByVal SourceValue As IBindingPath, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter, Optional ByVal Validator As IValueValidator, Optional ByVal ErrorAdorner As IDynamicAdorner, Optional ByVal TitleSource As Variant) As MSForms.TextBox
    
    If Not IsEmpty(TitleSource) Then
        CreateLabel TitleSource, ApplyMargin:=False
    End If
    
    Dim result As MSForms.TextBox
    Set result = this.Container.add(FormsProgID.TextBoxProgId)
    
    Dim Trigger As BindingUpdateSourceTrigger
'    'VF: Logic error?: would always be zero and that means never // missing: if not validator is nothing then ??? trigger not picked up from validator
'    'is updated later. Why not set here and skip update later?
'    If Validator Is Nothing Then
'        Trigger = Validator.Trigger '= 0 that is never
''    '    'VF:
''    '    Else ' set value of validator
''    '        Trigger = Validator.Trigger
''    '    End If
'    End If
    'VF: simplifies to unconditional
    'but causes DecimalValuator::IValueValidator_Trigger() to be called twice
    Trigger = Validator.Trigger
    
    
    this.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.path, result, _
        StringFormat:=FormatString, _
        Converter:=Converter, _
        Validator:=Validator, _
        ValidationAdorner:=ErrorAdornerOrDefault(Validator, ErrorAdorner, result), _
        UpdateTrigger:=Trigger
    
    Set IDynamicControlBuilder_TextBoxFor = result
    
End Function

