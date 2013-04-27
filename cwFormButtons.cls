VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwFormButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Alpha As Single, dx As Single, dy As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
 
Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
  W.CanGetFocus = False
  W.Move 0, 0, 70, 18
End Sub
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property
 
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim Pat As cCairoPattern
'  Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
'  Pat.AddColorStop 0, W.BackColor, Alpha, 1.1
'  Pat.AddColorStop 1, W.BackColor, Alpha, 0.9
'  CC.Rectangle 0, 0, dx, dy - 1, True
'  CC.Fill , Pat
'
'  CC.SetLineWidth 1, True
'  CC.SetLineCap CAIRO_LINE_CAP_SQUARE
'  CC.SetSourceColor W.BorderColor
'  CC.DrawLine 0, dy - 1, dx, dy - 1, True
'  CC.Stroke

End Sub
  
