VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("AccountsManager.System.Excel")
Option Explicit

'These constants define the default restoration behaviors for the class
Private Const APP_CALCULATION_STATE = True
Private Const APP_DISPLAY_ALERTS_STATE = True
Private Const APP_ENABLE_EVENTS_STATE = True
Private Const APP_SCREEN_UPDATING_STATE = True
Private Const APP_VISIBIL = True
Private Const APP_DISPLAY_FORMULABAR_STATE = True
Private Const APP_ANIMATIONS_STATE = True

'Set this to true to ensure a persisted state is restored, even if the consumer forgets to restore
Private Const RESTORE_ON_TERMINATE = True

'Private members stored in a Type
Private this As TMembers

Private Type TMembers
  Calculation As XlCalculationState
  DisplayAlerts As Boolean
  EnableEvents As Boolean
  ScreenUpdating As Boolean
  Visible As Boolean
  DisplayFormulaBar As Boolean
  EnableAnimations As Boolean
  
  RestoreCalcualion As Boolean
  RestoreDisplayAlert As Boolean
  RestoreEventsEnable As Boolean
  RestoreScreenUpdate As Boolean
  RestoreAppVisible As Boolean
  RestoreDisplayFormulaBar As Boolean
  RestoreEnableAnimations As Boolean
  
  IsPersisted As Boolean
  IsRestored As Boolean
End Type

'Set the default restoration behaviours on intialize
Private Sub Class_Initialize()
  this.RestoreCalcualion = APP_CALCULATION_STATE
  this.RestoreDisplayAlert = APP_DISPLAY_ALERTS_STATE
  this.RestoreEventsEnable = APP_ENABLE_EVENTS_STATE
  this.RestoreScreenUpdate = APP_SCREEN_UPDATING_STATE
  this.RestoreAppVisible = APP_VISIBIL
  this.RestoreDisplayFormulaBar = APP_DISPLAY_FORMULABAR_STATE
  this.RestoreEnableAnimations = APP_ANIMATIONS_STATE
End Sub

'By default, restore the settings if we didn't do it explicitly
Private Sub Class_Terminate()
  If this.IsPersisted And Not this.IsRestored And RESTORE_ON_TERMINATE Then
    Me.RestoreApplicationProperties
  End If
End Sub

Public Property Get RestoreCalculation() As Boolean
  RestoreCalculation = this.RestoreCalcualion
End Property

Public Property Let RestoreCalculation(value As Boolean)
  this.RestoreCalcualion = value
End Property

Public Property Get RestoreDisplayAlerts() As Boolean
  RestoreDisplayAlerts = this.RestoreDisplayAlert
End Property

Public Property Let RestoreDisplayAlerts(value As Boolean)
  this.RestoreDisplayAlert = value
End Property

Public Property Get RestoreEnableEvents() As Boolean
  RestoreEnableEvents = this.RestoreEventsEnable
End Property

Public Property Let RestoreEnableEvents(value As Boolean)
  this.RestoreEventsEnable = value
End Property

Public Property Get RestoreScreenUpdating() As Boolean
  RestoreScreenUpdating = this.RestoreScreenUpdate
End Property

Public Property Let RestoreScreenUpdating(value As Boolean)
  this.RestoreScreenUpdate = value
End Property

Public Property Get RestoreAppVisible() As Boolean
  RestoreAppVisible = this.RestoreAppVisible
End Property

Public Property Let RestoreAppVisible(value As Boolean)
  this.RestoreAppVisible = value
End Property

Public Property Get RestoreDisplayFormulaBar() As Boolean
  RestoreDisplayFormulaBar = this.RestoreAppVisible
End Property

Public Property Let RestoreDisplayFormulaBar(value As Boolean)
  this.RestoreDisplayFormulaBar = value
End Property

Public Property Get RestoreEnableAnimations() As Boolean
  RestoreEnableAnimations = this.RestoreEnableAnimations
End Property

Public Property Let RestoreEnableAnimations(value As Boolean)
  this.RestoreEnableAnimations = value
End Property

Public Sub SaveApplicationProperties()
  If Not this.IsPersisted Then
    'Save all of the settings
    With Application
      this.Calculation = .Calculation
      this.DisplayAlerts = .DisplayAlerts
      this.EnableEvents = .EnableEvents
      this.ScreenUpdating = .ScreenUpdating
      this.Visible = .Visible
      this.DisplayFormulaBar = .DisplayFormulaBar
      this.EnableAnimations = .EnableAnimations
    End With
    this.IsPersisted = True
  Else
    Err.Raise CustomError, "ExcelProperties", "Properties have already been persisted."
  End If
End Sub

Public Sub RestoreApplicationProperties()
  'Only restore the settings that we want restored
  '(which by default is all of them)
  With Application
    If this.RestoreCalcualion Then
      .Calculation = this.Calculation
    End If
    If this.RestoreDisplayAlert Then
      .DisplayAlerts = this.DisplayAlerts
    End If
    If this.RestoreEventsEnable Then
      .EnableEvents = this.EnableEvents
    End If
    If this.RestoreScreenUpdate Then
      .ScreenUpdating = this.ScreenUpdating
    End If
    If this.RestoreAppVisible Then
        .Visible = this.Visible
    End If
    If this.RestoreDisplayFormulaBar Then
        .DisplayFormulaBar = this.DisplayFormulaBar
    End If
    If this.EnableAnimations Then
        .EnableAnimations = this.EnableAnimations
    End If
  End With
  this.IsRestored = True
End Sub

Public Sub SetThisApplicationProperties(Optional ByVal swich As Boolean = True)
    With Application
        .Calculation = IIf(swich, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not swich
        .ScreenUpdating = Not swich
        .Visible = Not swich
        .DisplayFormulaBar = Not swich
        .EnableAnimations = Not swich
    End With
 
End Sub
