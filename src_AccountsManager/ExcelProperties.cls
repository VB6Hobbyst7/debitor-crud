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
  
  RestoreCA As Boolean
  RestoreCalcualion As Boolean
  RestoreDA As Boolean
  RestoreDisplayAlert As Boolean
  RestoreEE As Boolean
  RestoreEventsEnable As Boolean
  RestoreSU As Boolean
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
    Me.Restore
  End If
End Sub

Public Property Get RestoreCalculation() As Boolean
  RestoreCalculation = this.RestoreCalcualion
End Property

Public Property Let RestoreCalculation(Value As Boolean)
  this.RestoreCalcualion = Value
End Property

Public Property Get RestoreDisplayAlerts() As Boolean
  RestoreDisplayAlerts = this.RestoreDisplayAlert
End Property

Public Property Let RestoreDisplayAlerts(Value As Boolean)
  this.RestoreDisplayAlert = Value
End Property

Public Property Get RestoreEnableEvents() As Boolean
  RestoreEnableEvents = this.RestoreEE
End Property

Public Property Let RestoreEnableEvents(Value As Boolean)
  this.RestoreEE = Value
End Property

Public Property Get RestoreScreenUpdating() As Boolean
  RestoreScreenUpdating = this.RestoreScreenUpdate
End Property

Public Property Let RestoreScreenUpdating(Value As Boolean)
  this.RestoreScreenUpdate = Value
End Property



Public Sub Save()
  If Not this.IsPersisted Then
    'Save all of the settings
    With Application
      this.Calculation = .Calculation
      this.DisplayAlerts = .DisplayAlerts
      this.EnableEvents = .EnableEvents
      this.ScreenUpdating = .ScreenUpdating
    End With
    this.IsPersisted = True
  Else
    Logging.logFATAL "ExcelProperties", "Properties have already been persisted."
    Err.Raise CustomError, "ExcelProperties", "Properties have already been persisted."
  End If
End Sub

Public Sub Restore()
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
  End With
  this.IsRestored = True
End Sub

Public Sub SetApplicationProperties()

    With Application
          .Calculation = xlCalculationManual
          .DisplayAlerts = False
          .EnableEvents = False
          .ScreenUpdating = False
          .Visible = False
          .DisplayFormulaBar = False
          .EnableAnimations = False
          
    End With

End Sub
