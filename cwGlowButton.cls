VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwGlowButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public BlendInSteps As Long, BlendOutSteps As Long

'--------- default Widget-conventions for a "cwClass", as usual --------
Private Alpha As Single, dx As Single, dy As Single
Private ResHeightPerEntry As Long

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrBlend As cTimer, BlAlpha As Double, Pressed As Boolean
Attribute tmrBlend.VB_VarHelpID = -1
Private mCaption As String, mChecked As Boolean

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase

  BlendInSteps = 15
  BlendOutSteps = 30
End Sub
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property
'--------- end of default Widget-conventions -----------

Public Property Get Checked() As Boolean
  Checked = mChecked
End Property
Public Property Let Checked(ByVal NewValue As Boolean)
  If mChecked = NewValue Then Exit Property
  mChecked = NewValue
  W.Refresh
End Property

Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
  If mCaption = NewValue Then Exit Property
  mCaption = NewValue
  W.Refresh
End Property

Private Sub tmrBlend_Timer()
  Select Case tmrBlend.Tag
    Case "MouseEnter"
      If BlendInSteps Then BlAlpha = BlAlpha + 1 / BlendInSteps Else BlAlpha = 1.1
    Case "MouseLeave"
      If BlendOutSteps Then BlAlpha = BlAlpha - 1 / BlendOutSteps Else BlAlpha = -0.1
  End Select
  If BlAlpha < 0 Or BlAlpha > 1 Then Set tmrBlend = Nothing: Exit Sub
  W.Refresh
End Sub

Private Sub W_AddedToHierarchy()
  With Cairo.ImageList(W.ImageKey)
    ResHeightPerEntry = .Height \ 6
    W.Move W.Left, W.Top, .Width, ResHeightPerEntry
  End With
End Sub
 
Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  Pressed = True
  W.Refresh
End Sub
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  W.MouseIconImageKey = IIf(W.Enabled, "cur_hand,17,9", "cur_unavailable,8,8")
End Sub
Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  Pressed = False
  W.Refresh
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As cWidgetBase)
  If BlAlpha < 0 Then BlAlpha = 0
  If W.Enabled Then Set tmrBlend = New_c.Timer(15, True, "MouseEnter")
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As cWidgetBase)
  If BlAlpha > 1 Then BlAlpha = 1
  If W.Enabled Then Set tmrBlend = New_c.Timer(15, True, "MouseLeave")
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim Pat As cCairoPattern, M As cCairoMatrix, TOffs As Long

  Set Pat = Cairo.CreateSurfacePattern(Cairo.ImageList(W.ImageKey))
  If ResHeightPerEntry = 0 Then ResHeightPerEntry = Pat.Surface.Height \ 6
  Set M = Cairo.CreateIdentityMatrix

  If W.Enabled Then
    Set Pat.Matrix = M.TranslateCoords(0, ResHeightPerEntry * IIf((Pressed And W.MouseOver) Or mChecked, 2, 0))
    CC.Paint Alpha, Pat

    Set Pat.Matrix = M.TranslateCoords(0, ResHeightPerEntry) 'one additional shift when hovered (this translate adds to the one above)
    If BlAlpha > 0 Then CC.Paint Alpha * BlAlpha, Pat

    If W.Focused Then
      Set Pat.Matrix = M.TranslateCoords(0, ResHeightPerEntry * IIf((Pressed And W.MouseOver) Or mChecked, 2, 4))
      CC.Paint 1, Pat
    End If
  Else 'and the disabled state sits at last position 5 (at offset Pat.Surface.Height * 0.8)

    Set Pat.Matrix = M.TranslateCoords(0, ResHeightPerEntry * 4)
    CC.Paint Alpha, Pat
  End If
  
  If Len(mCaption) Then
    W.SelectFontSettingsInto CC
    TOffs = IIf(Pressed Or mChecked, 1, 0)
    CC.DrawText -1 + TOffs, 0 + TOffs, dx, dy, mCaption, True, vbCenter, 0, True
  End If
End Sub

