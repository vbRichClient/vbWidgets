VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLayoutVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const enPrefix As String = "<voice required=""Language=409"">", enSuffix As String = "</voice>"

Public SpeechSupport As Boolean, MarginLeft, MarginTop, MarginRight, MarginBottom

Private WithEvents mParent As cWidgetForm, WithEvents tmrSpeech As cTimer
Attribute mParent.VB_VarHelpID = -1
Attribute tmrSpeech.VB_VarHelpID = -1
Private mBands As New Collection, mAutoAdjustFormHeight As Boolean, mMinWidth As Long
 
Public Sub AttachTo(ByVal Parent As cWidgetForm, Optional ByVal AutoAdjustFormHeight As Boolean, Optional ByVal OtherVoice As Object)
  Set mParent = Parent
  If Not OtherVoice Is Nothing Then Set Voice = OtherVoice
  mAutoAdjustFormHeight = AutoAdjustFormHeight
  mMinWidth = mParent.Width
  MarginLeft = 5: MarginTop = 5: MarginRight = 5: MarginBottom = 5
  Set tmrSpeech = New_c.Timer(500)
End Sub

Public Function BandCount() As Long
  BandCount = mBands.Count
End Function

Public Function Band(ByVal IndexZeroBased As Long) As cLayoutBand
  Set Band = mBands(IndexZeroBased + 1)
End Function

Public Function AddBandWithCells(ParamArray CellWidths()) As cLayoutBand
  If mParent Is Nothing Then Exit Function
  Dim TmpCellWidths(): TmpCellWidths = CellWidths 'copy the Param-Array
  If UBound(TmpCellWidths) = -1 Then ReDim TmpCellWidths(0 To 0): TmpCellWidths(0) = 1
  Set AddBandWithCells = New cLayoutBand
      AddBandWithCells.Init mParent, TmpCellWidths, mBands.Count
  mBands.Add AddBandWithCells
End Function

Public Sub Message(Msg, Optional ByVal LangID As Long)
  If SpeechSupport Then Speak IIf(LangID, "<voice required=""Language=" & LangID & """>" & Msg & "</voice>", Msg) Else VBA.MsgBox Msg, vbInformation
End Sub

Public Sub Resize(Optional ByVal NewWidth As Long, Optional ByVal NewHeight As Long)
  If mParent Is Nothing Then Exit Sub
  Dim Band As cLayoutBand, yOffs As Single
  If NewWidth = 0 Then NewWidth = mParent.ScaleWidth
  If NewHeight = 0 Then NewHeight = mParent.ScaleHeight
  
  mParent.Locked = True
    yOffs = MarginTop
    For Each Band In mBands
      Band.ResizeWith MarginLeft, yOffs, NewWidth / mParent.WidgetRoot.Zoom - MarginLeft - MarginRight
    Next
    If mAutoAdjustFormHeight Then
      NewHeight = yOffs * mParent.WidgetRoot.Zoom + mParent.Height - mParent.ScaleHeight
      mParent.SetMinMaxDimensions mMinWidth, NewHeight
      If mParent.Height < NewHeight Then mParent.Move mParent.Left, mParent.Top, NewWidth, NewHeight
    End If
  mParent.Locked = False
End Sub
 
Private Sub mParent_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
  If SpeechSupport = False Then Exit Sub
  On Error Resume Next
  
  Select Case EventName
    Case "W_GotFocus"
      SpeakDelayed GetControlInfo(Sender)
    Case "Change"
      If TypeOf Sender Is cwLabeledTextBox Then SpeakDelayed enPrefix & ", change: " & enSuffix & Sender.Text
    Case "SelChanged"
      If TypeOf Sender Is cwLabeledTextBox Then SpeakDelayed enPrefix & ", position-change: " & enSuffix & Mid$(Sender.Text, Sender.SelStart + 1, IIf(Sender.SelLength, Sender.SelLength, 100))
    Case "Click"
      If TypeOf Sender Is cwButton Then If Sender.ButtonStyle > 1 Then SpeakDelayed enPrefix & ", change: " & IIf(Sender.ButtonStyle < 4, Sender.Value, IIf(Sender.Value, "true", "false")) & enSuffix
    Case "DataSourceAddNew"
      If TypeOf Sender Is cwGrid Then Sender.Widget.ToolTip = "- after add-new "
    Case "DataSourceDelete"
      If TypeOf Sender Is cwGrid Then Sender.Widget.ToolTip = "- after delete "
    Case "DataSourcePositionChanged"
      If TypeOf Sender Is cwGrid Then SpeakDelayed enPrefix & ", grid-position change, " & Sender.Widget.ToolTip & ": " & GetRowInfo(Sender.DataSource): Sender.Widget.ToolTip = ""
  End Select
End Sub

Private Function GetControlInfo(Sender As Object) As String
Dim Name As String
  On Error Resume Next
  Name = Sender.Caption
  If Len(Name) = 0 Then Name = Sender.Widget.Tag
  If Len(Name) = 0 Then Name = Sender.Widget.Key
  Name = Replace(Name, "&", "")
  If TypeOf Sender Is cwButton Then
    Select Case Sender.ButtonStyle
      Case 0:    GetControlInfo = enPrefix & "Buttton" & enSuffix & Name 'Button deliberately written this way (correctly written it sounded terrible on my machine)
      Case 1:    GetControlInfo = enPrefix & "ToolButtton" & enSuffix & Name 'ToolButtton deliberately written this way (correctly written it sounded not nice on my machine)
      Case 2, 3: GetControlInfo = enPrefix & "CheckBox" & enSuffix & Name & enPrefix & "value: " & Sender.Value & enSuffix
      Case 4, 5: GetControlInfo = enPrefix & "OptionBox" & enSuffix & Name & enPrefix & "value: " & IIf(Sender.Value, "true", "false") & enSuffix
    End Select
  ElseIf TypeOf Sender Is cwGrid Then
    GetControlInfo = enPrefix & "Grid" & enSuffix & Name & enPrefix & ", " & GetRowInfo(Sender.DataSource)
  ElseIf TypeOf Sender Is cwTextBox Then
    GetControlInfo = enPrefix & "TextBox" & enSuffix & Name & enPrefix & IIf(Len(Sender.Text), "contains: " & enSuffix & Sender.Text, enSuffix)
  Else
    GetControlInfo = enPrefix & Mid$(TypeName(Sender), 3) & enSuffix & Name
  End If
End Function

Private Function GetRowInfo(DS As cDataSource, Optional ByVal i As Long) As String
  If DS Is Nothing Then Exit Function
  If DS.AbsolutePosition < 1 And DS.RecordCount > 0 Then Exit Function
  If DS.RecordCount = 0 Then
    GetRowInfo = "no records in grid" & enSuffix
  Else
    With New_c.StringBuilder
      .Append enSuffix
      For i = 0 To DS.FieldCount - 1
        .Append DS.FieldName(i) & ":"
        Select Case VarType(DS.FieldValue(i))
          Case vbEmpty: .Append "Null:"
          Case vbByte Or vbArray: .Append "Byte-Array:"
          Case Else: .Append DS.FieldValue(i) & ":"
        End Select
      Next
      GetRowInfo = .ToString
    End With
  End If
End Function
 
Private Sub SpeakDelayed(ByVal Text As String)
  Speak ""
  tmrSpeech.Enabled = False
  tmrSpeech.Enabled = True
  tmrSpeech.Tag = Text
End Sub

Private Sub mParent_ResizeWithDimensions(ByVal NewWidth As Long, ByVal NewHeight As Long)
  Resize NewWidth, NewHeight
End Sub

Private Sub tmrSpeech_Timer()
  tmrSpeech.Enabled = False
  Speak tmrSpeech.Tag
End Sub
