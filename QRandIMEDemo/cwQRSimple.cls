VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwQRSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents W As cWidgetBase, TBW As cwTBoxWrap, mCaption As String
Attribute W.VB_VarHelpID = -1

Public QRSrf As cCairoSurface

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase 'instantiate the W-WidgetBase which later on connects us to our "hosting environment"
      W.FontName = "Tahoma"
      W.FontSize = 11
      W.ForeColor = vbWhite
      W.ForwardFocus = True
      W.Moveable = True
      W.ImageKey = "QRico2"
      W.Alpha = 0.7
      
  Set TBW = Widgets.Add(New cwTBoxWrap, "TBW", 0, 0, 1, 28)
      TBW.TBox.Widget.FontName = "Tahoma"
      TBW.TBox.Widget.FontSize = 11
      TBW.TBox.Widget.FontBold = False
      TBW.Widget.ImageKey = "QRico"
End Sub
 
Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets: Set Widgets = W.Widgets: End Property

Public Property Get Text() As String
  Text = TBW.Text
End Property
Public Property Let Text(NewValue As String)
  TBW.Text = NewValue
  W.Refresh
End Property

Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(NewValue As String)
  mCaption = NewValue
  W.Refresh
End Property

Public Property Get TBox() As cwTextBox
  Set TBox = TBW.TBox
End Property
 
Private Sub W_EnterFocus()
  W.Refresh
End Sub
Private Sub W_ExitFocus()
  W.Refresh
End Sub

Private Sub W_Resize()
  TBW.Widget.Move 10, W.ScaleHeight - TBW.Widget.Height - 10, W.ScaleWidth - 20
End Sub

Private Sub W_KeyPress(KeyAscii As Integer)
  TBW.Widget.KeyPress KeyAscii 'just delegate it to the Text-Widget
End Sub
 
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  DrawShape CC, dx_Aligned, dy_Aligned
  DrawCaption CC, dx_Aligned, dy_Aligned
  DrawQRSurface CC, dx_Aligned, dy_Aligned
End Sub

Private Sub DrawShape(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single)
Const R As Double = 5, ShadowOffs& = 5
  
  'the main-shape
  CC.FillRule = CAIRO_FILL_RULE_EVEN_ODD
  CC.SetLineWidth 0.8
  CC.RoundedRect ShadowOffs, ShadowOffs, dx - 2 * ShadowOffs, dy - 2 * ShadowOffs, R, True
  CC.RoundedRect ShadowOffs + 8, ShadowOffs + 29, dx - (ShadowOffs + 21), dy - (ShadowOffs + 74), R, True
    CC.SetSourceColor &H88FFFF, 0.15
  CC.Fill True
    CC.SetSourceColor vbWhite, 0.15
  CC.Stroke
  
  'the Drop-Shadow
  Dim i As Long, j As Long, Pat As cCairoPattern
  Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
      Pat.AddColorStop 0, vbBlack, 0
      Pat.AddColorStop 0.05, IIf(W.Active, &H330033, &H111111), 0.5
      Pat.AddColorStop 1, IIf(W.Active, &H330033, &H111111), 0.8
  For i = 1 To ShadowOffs
    j = i - 1
    CC.RoundedRect ShadowOffs - j, ShadowOffs, dx - 2 * (ShadowOffs - j), dy - 2 * ShadowOffs + j, R + j * 1.2
    CC.RoundedRect ShadowOffs - i, ShadowOffs, dx - 2 * (ShadowOffs - i), dy - 2 * ShadowOffs + i, R + i * 1.2
    CC.Clip
      CC.Paint 1 / Exp(0.6 * i), Pat
    CC.ResetClip
  Next i
  
  'the Caption-Icon
  If Not W.Active Then CC.Operator = CAIRO_OPERATOR_HSL_LUMINOSITY
  CC.RenderSurfaceContent W.ImageKey, ShadowOffs + 6, ShadowOffs + 5, 21, 21
  CC.Operator = CAIRO_OPERATOR_OVER
End Sub
 
Private Sub DrawCaption(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single)
  W.SelectFontSettingsInto CC, &H111111
  CC.DrawText 34, 1, dx - 40, 42, mCaption, True, vbLeftJustify, 2, True
  W.SelectFontSettingsInto CC
  CC.DrawText 34, 0, dx - 40, 42, mCaption, True, vbLeftJustify, 2, True
End Sub

Private Sub DrawQRSurface(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single)
Dim x, y
  Cairo.CalcAspectFit 1, dx - 4, dy - 87, x, y, dx, dy
  If QRSrf Is Nothing Then Set QRSrf = Cairo.CreateSurface(1, 1) 'create an empty 1x1-Surface
  CC.RenderSurfaceContent QRSrf, x + 2, y + 40, dx, dy
End Sub

