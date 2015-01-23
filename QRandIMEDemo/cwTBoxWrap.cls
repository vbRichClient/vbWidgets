VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwTBoxWrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Public TBox As cwTextBox 'we make the wrapped cwTextBox also available on the outside
 
Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase  'instantiate the W-WidgetBase which later on connects us to our "hosting environment"
      W.ForwardFocus = True
      W.BackColor = -1
 
  Set TBox = Widgets.Add(New cwTextBox, "TBox", 0, 0, 1, 21)
      TBox.Widget.FontName = "Tahoma"
      TBox.Widget.FontSize = 12
      TBox.Widget.FontBold = True
      TBox.Widget.Alpha = 0.75
      TBox.Widget.BackColor = -1
      TBox.Widget.ForeColor = &HCCCCCC
      TBox.Widget.SelectionColor = vbYellow
      TBox.TextShadowColor = vbBlack
      TBox.TextShadowOffsetX = 1
      TBox.TextShadowOffsetY = 1.2
      TBox.Border = False
      TBox.VCenter = True
End Sub
 
Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets: Set Widgets = W.Widgets: End Property

Public Property Get Text() As String
  Text = TBox.Text
End Property
Public Property Let Text(NewValue As String)
  TBox.Text = NewValue
End Property

Private Sub W_Resize()
  TBox.Widget.Move W.ScaleHeight - 1, 0, W.ScaleWidth - W.ScaleHeight - 7, W.ScaleHeight - 3
End Sub

Private Sub W_EnterFocus()
  W.Refresh
End Sub
Private Sub W_ExitFocus()
  W.Refresh
End Sub

Private Sub W_KeyPress(KeyAscii As Integer)
  TBox.Widget.KeyPress KeyAscii 'just delegate it to the Text-Widget
End Sub
 
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  DrawShape CC, dx_Aligned, dy_Aligned
End Sub

Private Sub DrawShape(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single)
Dim i As Long, Pat As cCairoPattern

  Set Pat = Cairo.ImageList("bgTxtBlack").CreateSurfacePattern
      Pat.Extend = CAIRO_EXTEND_REPEAT
      CC.RoundedRect 1.25, 0.65, dx - 2.15, dy - 2.8, dy / 2
      CC.Fill , Pat

  Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
      Pat.AddGaussianStops_ThreeColors &H60606, &H111111, &HDDDDDD, 0.65, 0.25, 0.04, 0.4, gpLess
  CC.SetLineWidth 0.36
  For i = 0 To 4
    CC.SetSourcePattern Pat
    CC.RoundedRect 0.2, 0.35 + i * 0.35, dx, dy - 1.8, dy / 2, True
    CC.Stroke
  Next i
  
  CC.DrawLine dy / 2 - 1.2, 0, dx - dy / 2 + 1.2, 0, True, 1, vbBlack, 0.5
  CC.DrawLine dy / 2 - 1.2, dy - 1.2, dx - dy / 2 + 1.2, dy - 1.2, True, 1, &HDDDDDD, 0.06

  If Not W.Active Then CC.Operator = CAIRO_OPERATOR_HSL_LUMINOSITY
  If Len(W.ImageKey) Then CC.RenderSurfaceContent W.ImageKey, 1 + dy \ 7, 3, dy - 7, dy - 7
End Sub
  
