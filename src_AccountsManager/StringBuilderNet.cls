VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilderNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "StringBuilder implementation modified from https://codereview.stackexchange.com/q/196076"
'@Folder AccountsManager.Infrastructure.StringFormatting
'@ModuleDescription "StringBuilder implementation modified from https://codereview.stackexchange.com/q/196076"
'>> modified identifier names to better fit this project
'>> modified to prevent a stateful default instance
'@IgnoreModule ParameterCanBeByVal: string parameters are passed ByRef to avoid copying a string of unknown length, for performance reasons.
'@PredeclaredId
'@Exposed
Option Explicit
'Based off vba_stringbuilder:Class - http://blog.livedoor.jp/midorityo/archives/50744656.html
Private Type TState
    AppendsBufferSize As Long
    Builder As Object 'NOTE: 'VF: referring to .net System.Text.StringBuilder <- must be late bound(?) so cannot be more specific? for intellisense/validation at compile time .capacity .length .MaxCapacity ?set reference and as System.Text.StringBuilder? <-dependency! without telling error handler message like .net not installed
    Index As Long
End Type
Private AppendsBuffer() As String
Private this As TState

Private Sub Class_Initialize()
    'NOTE: 'VF: why not pass through guardclause?
    If Not IsDefaultInstance Then
        Const DEFAULT_APPENDSBUFFER_SIZE As Long = 10000
        Set this.Builder = CreateObject("System.Text.StringBuilder")
        SetAppendsBufferSize DEFAULT_APPENDSBUFFER_SIZE
        this.Index = -1
    End If
End Sub

Public Property Get Capacity() As Long
    GuardClauses.GuardDefaultInstance Me, StringBuilderNet
    Capacity = this.Builder.Capacity
End Property

Private Property Let Capacity(ByVal RHS As Long)
    GuardClauses.GuardDefaultInstance Me, StringBuilderNet
    this.Builder.Capacity = RHS
End Property

Private Property Get Length() As Long
    GuardClauses.GuardDefaultInstance Me, StringBuilderNet
    Length = this.Builder.Length
End Property

Public Property Let Length(ByVal RHS As Long)
    GuardClauses.GuardDefaultInstance Me, StringBuilderNet
    this.Builder.Length = RHS
End Property

Public Function MaxCapacity() As Long
    GuardClauses.GuardDefaultInstance Me, StringBuilderNet
    MaxCapacity = this.Builder.MaxCapacity
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is StringBuilderNet
End Property

Public Function Append(ByRef text As String, Optional RepeatCount As Long = 0) As StringBuilderNet
    If IsDefaultInstance Then
        Dim result As StringBuilderNet
        Set result = New StringBuilderNet
        Set Append = result.Append(text, RepeatCount)
        Exit Function
    End If
    
    this.Index = this.Index + 1
    AppendsBuffer(this.Index) = text
    If RepeatCount > 0 Then result.Append text, (RepeatCount - 1)
    If this.Index = this.AppendsBufferSize Then Compact
    Set Append = Me
End Function

Public Function AppendFormat(ByVal Format As Variant, ParamArray Args() As Variant) As StringBuilderNet
    Dim LocalArgs() As Variant
    If IsArray(Args(LBound(Args))) Then
        LocalArgs = Args(LBound(Args))
    Else
        ReDim LocalArgs(LBound(Args) To UBound(Args))
        Dim Index As Long
        For Index = LBound(Args) To UBound(Args)
            LocalArgs(Index) = Args(Index)
        Next
    End If
    
    If IsDefaultInstance Then
        Dim result As StringBuilderNet
        Set result = New StringBuilderNet
        Set AppendFormat = result.AppendFormat(Format, LocalArgs)
        Exit Function
    End If
    
    Compact

    If Format = vbNullString Then
        Set AppendFormat = Me.Append(VBA.Strings.Join(LocalArgs))
        Exit Function
    End If

    this.Builder.AppendFormat_4 Format, LocalArgs
    Set AppendFormat = Me
End Function

Public Function InsertFormat(ByVal Format As Variant, StartIndex As Long, RepeatCount As Long, ParamArray Args() As Variant) As StringBuilderNet
    Dim LocalArgs() As Variant
    If IsArray(Args(LBound(Args))) Then
        LocalArgs = Args(LBound(Args))
    Else
        ReDim LocalArgs(LBound(Args) To UBound(Args))
        Dim Index As Long
        For Index = LBound(Args) To UBound(Args)
            LocalArgs(Index) = Args(Index)
        Next
    End If
    
    If IsDefaultInstance Then
        Dim result As StringBuilderNet
        Set result = New StringBuilderNet
        Set InsertFormat = result.InsertFormat(Format, StartIndex, RepeatCount, LocalArgs)
        Exit Function
    End If
    
    If Format = vbNullString Then
        Set InsertFormat = Me.Insert(StartIndex, VBA.Strings.Join(LocalArgs))
        Exit Function
    End If
    
    With New StringBuilderNet
        .AppendFormat Format, LocalArgs
        Me.Insert StartIndex, .ToString
    End With
    Set InsertFormat = Me
End Function

Public Function AppendLine(Optional ByRef text As String, Optional ByVal RepeatCount As Long = 0) As StringBuilderNet
    If IsDefaultInstance Then
        Dim result As StringBuilderNet
        Set result = New StringBuilderNet
        Set AppendLine = result.AppendLine(text, RepeatCount)
        Exit Function
    End If
    
    this.Index = this.Index + 1
    AppendsBuffer(this.Index) = text & vbNewLine
    If RepeatCount > 0 Then Me.AppendLine text, (RepeatCount - 1)
    If this.Index = this.AppendsBufferSize Then Compact
    Set AppendLine = Me
End Function

Public Sub Clear()
    GuardClauses.GuardDefaultInstance Me, StringBuilderNet
    Me.Length = 0
End Sub

Private Sub Compact()
    GuardClauses.GuardDefaultInstance Me, StringBuilderNet
    If this.Index > -1 Then
        this.Builder.Append_3 Join(AppendsBuffer, vbNullString)
        this.Index = -1
    End If
    ReDim AppendsBuffer(this.AppendsBufferSize)
End Sub

Public Function EnsureCapacity(ByVal Length As Long) As StringBuilderNet
    If IsDefaultInstance Then
        Dim result As StringBuilderNet
        Set result = New StringBuilderNet
        Set EnsureCapacity = result.EnsureCapacity(Length)
        Exit Function
    End If
    this.Builder.EnsureCapacity Length
    Set EnsureCapacity = Me
End Function

'@Ignore ParameterCanBeByVal: Text is passed ByRef to avoid copying a string of unknown length, for performance reasons.
Public Function Insert(ByVal StartIndex As Long, ByRef text As String, Optional ByVal RepeatCount As Long = 0) As StringBuilderNet
    If IsDefaultInstance Then
        Dim result As StringBuilderNet
        Set result = New StringBuilderNet
        Set Insert = result.Insert(StartIndex, text, RepeatCount)
        Exit Function
    End If
    Compact
    this.Builder.Insert_2 StartIndex, text
    If RepeatCount > 0 Then Me.Insert StartIndex, text, (RepeatCount - 1)
    Set Insert = Me
End Function

Public Function Remove(ByVal StartIndex As Long, ByVal Length As Long) As StringBuilderNet
    If IsDefaultInstance Then
        Dim result As StringBuilderNet
        Set result = New StringBuilderNet
        Set Remove = result.Remove(StartIndex, Length)
        Exit Function
    End If
    Compact
    this.Builder.Remove StartIndex, Length
    Set Remove = Me
End Function

Public Function Replace(ByRef OldValue As String, ByRef NewValue As String, Optional ByVal StartIndex As Long = -1, Optional ByVal RepeatCount As Long = 0) As StringBuilderNet
    If IsDefaultInstance Then
        Dim result As StringBuilderNet
        Set result = New StringBuilderNet
        Set Replace = result.Replace(OldValue, NewValue, StartIndex, RepeatCount)
        Exit Function
    End If
    
    Compact

    If StartIndex > -1 And RepeatCount > 0 Then
        this.Builder.Replace_2 OldValue, NewValue, StartIndex, RepeatCount
    Else
        this.Builder.Replace OldValue, NewValue
    End If
    Set Replace = Me
End Function

Private Sub SetAppendsBufferSize(ByVal Length As Long)
    GuardClauses.GuardDefaultInstance Me, StringBuilderNet
    this.AppendsBufferSize = Length
    Compact
End Sub

Public Function ToString(Optional ByVal StartIndex As Long = -1, Optional ByVal Length As Long = 0) As String
    GuardClauses.GuardDefaultInstance Me, StringBuilderNet
    Compact

    If StartIndex > -1 And Length > 0 Then
        ToString = this.Builder.ToString_2(StartIndex, Length)
    Else
        ToString = this.Builder.ToString
    End If
End Function
