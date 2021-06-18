VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetCellBindingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder AccountsManager.Infrastructure.Bindings.PropertyBindings.Strategies
Option Explicit
Implements IBindingTargetStrategy

Private Function IBindingTargetStrategy_DefaultPropertyBindingFor(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As Object, _
ByVal Mode As BindingMode, _
ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
ByVal Validator As IValueValidator, _
ByVal Converter As IValueConverter, _
ByVal StringFormat As IStringFormatter, _
ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding

    Set IBindingTargetStrategy_DefaultPropertyBindingFor = WorksheetCellPropertyBinding _
        .Create(Context, Source, Target, _
            Mode:=Mode, _
            Validator:=Validator, _
            Converter:=Converter, _
            StringFormat:=StringFormat, _
            ValidationAdorner:=ValidationAdorner)

End Function

Private Function IBindingTargetStrategy_PropertyBindingFor(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As IBindingPath, _
ByVal Mode As BindingMode, _
ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
ByVal Validator As IValueValidator, _
ByVal Converter As IValueConverter, _
ByVal StringFormat As IStringFormatter, _
ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding

    Set IBindingTargetStrategy_PropertyBindingFor = WorksheetCellPropertyBinding _
        .Create(Context, Source, Target.Context, Target.path, _
            Mode:=Mode, _
            Validator:=Validator, _
            Converter:=Converter, _
            StringFormat:=StringFormat, _
            ValidationAdorner:=ValidationAdorner)

End Function