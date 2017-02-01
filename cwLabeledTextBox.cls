VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwLabeledTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Change()
Event Validate(Cancel As Boolean)
Event SelChanged(ByVal OldSelStart As Long, ByVal OldSelLength As Long, ByVal OldCaretPosition As Long)

Private WithEvents W As cWidgetBase, WithEvents TB As cwTextBox
Attribute W.VB_VarHelpID = -1
Attribute TB.VB_VarHelpID = -1
Private mCaption As String, mAlignment As AlignmentConstants, mLabelWidthPercent As Single

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
      W.ForwardFocus = True
  mAlignment = vbRightJustify
  mLabelWidthPercent = 0.4
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Public Property Get CaptionWidth() As Single
  Dim CC As cCairoContext
  Set CC = W.MeasureContext
  W.SelectFontSettingsInto CC
  CaptionWidth = CC.GetTextExtents(mCaption)
End Property

Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
Dim AccKey$, Pos&
  If mCaption = NewValue Then Exit Property Else mCaption = NewValue
  TB.Widget.Tag = NewValue
  W.Refresh
  Pos = InStr(Replace(mCaption, "&&", "--"), "&")
  If Pos Then AccKey = Mid$(Replace(mCaption, "&&", "--"), Pos + 1, 1)

  If Len(AccKey) Then W.AccessKeys = AccKey
End Property

Public Property Get Alignment() As AlignmentConstants
  Alignment = mAlignment
End Property
Public Property Let Alignment(ByVal NewValue As AlignmentConstants)
  If mAlignment = NewValue Then Exit Property Else mAlignment = NewValue
  W.Refresh
End Property

Public Property Get Text() As String
  If Not TB Is Nothing Then Text = TB.Text
End Property
Public Property Let Text(ByVal NewValue As String)
  If Not TB Is Nothing Then TB.Text = NewValue
End Property

Public Property Get SelStart() As Long
  SelStart = TB.SelStart
End Property
Public Property Let SelStart(ByVal NewValue As Long)
  TB.SelStart = NewValue
End Property

Public Property Get SelLength() As Long
  SelLength = TB.SelLength
End Property
Public Property Let SelLength(ByVal NewValue As Long)
  TB.SelLength = NewValue
End Property

Public Property Get LabelWidthPercent() As Single
  LabelWidthPercent = mLabelWidthPercent
End Property
Public Property Let LabelWidthPercent(ByVal NewValue As Single)
  If mLabelWidthPercent = NewValue Then Exit Property Else mLabelWidthPercent = NewValue
  If mLabelWidthPercent < 0.02 Then mLabelWidthPercent = 0.02
  If mLabelWidthPercent > 0.98 Then mLabelWidthPercent = 0.98
  W.Refresh
End Property

Public Sub Move(x, y, dx, dy, Optional ByVal LabelWidthPercent As Single)
  If LabelWidthPercent > 0 And LabelWidthPercent < 1 Then mLabelWidthPercent = LabelWidthPercent
  W.Move x, y, dx, dy
End Sub

Private Sub TB_Change()
  RaiseEvent Change
  W.RaiseBubblingEvent Me, "Change"
End Sub
Private Sub TB_Validate(Cancel As Boolean)
  RaiseEvent Validate(Cancel)
  W.RaiseBubblingEvent Me, "Validate", Cancel
End Sub
Private Sub TB_SelChanged(ByVal OldSelStart As Long, ByVal OldSelLength As Long, ByVal OldCaretPosition As Long)
  RaiseEvent SelChanged(OldSelStart, OldSelLength, OldCaretPosition)
  W.RaiseBubblingEvent Me, "SelChanged", OldSelStart, OldSelLength, OldCaretPosition
End Sub

Private Sub W_AccessKeyPress(KeyAscii As Integer)
  If TB Is Nothing Then Set TB = Widgets.Add(New cwTextBox, "_" & W.Key)
  TB.Widget.SetFocus
  TB.SelectAll
End Sub
 
Private Sub W_EnterFocus()
  If TB Is Nothing Then Set TB = Widgets.Add(New cwTextBox, "_" & W.Key)
  TB.SelectAll
End Sub

Private Sub W_Resize()
  If TB Is Nothing Then Set TB = Widgets.Add(New cwTextBox, "_" & W.Key)
  TB.Widget.Move W.Width * mLabelWidthPercent, 0, W.Width * (1 - mLabelWidthPercent), W.Height
End Sub
 
Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  W.SelectFontSettingsInto CC
  CC.DrawText 0, 4, dx_Aligned * mLabelWidthPercent - 3, dy_Aligned - 4, mCaption, True, mAlignment, 0, 0, dtHasAccelerators
End Sub
