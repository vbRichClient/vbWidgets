VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event OwnerDrawOnTop(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single)

Private Alpha As Single, dx As Single, dy As Single, mCaption As String

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
 
Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
      W.ForwardFocus = True
End Sub
Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets: Set Widgets = W.Widgets: End Property

Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
  If mCaption = NewValue Then Exit Property
  mCaption = NewValue
  W.Refresh
End Property

Private Sub W_AddedToHierarchy()
  AdjustAlignment
End Sub

Private Sub W_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
  If EventName = "W_ContainerResize" Then AdjustAlignment
End Sub

Private Sub AdjustAlignment()
  W.Move 0, W.Parent.ScaleHeight - W.Height, W.Parent.ScaleWidth, W.Height
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim Pat As cCairoPattern
  If dx = 0 Then dx = W.Width
  If dy = 0 Then dy = W.Height
  
  Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
      Pat.AddColorStop 0, W.BackColor, Alpha, 1.1
      Pat.AddColorStop 1, W.BackColor, Alpha, 0.9
  CC.Rectangle 0, 0, dx + 1, dy + 1
  CC.Fill , Pat
  
  CC.SetLineCap CAIRO_LINE_CAP_SQUARE
    CC.SetLineWidth 1, True
    CC.SetSourceColor W.BorderColor, 0.8
    CC.DrawLine 0, 0, dx, 0, True
  CC.Stroke
'  CC.SetSourceColor vbWhite, 0.8
'    CC.DrawLine 0, 1, dx, 1, True
'  CC.Stroke
'  CC.SetSourceColor vbWhite, 0.8, 0.85
'    CC.DrawLine 0, dy - 1, dx, dy - 1, True
'  CC.Stroke
  If Len(mCaption) Then CC.DrawText 7, 0, dx - 7, dy, mCaption, , vbLeftJustify, 2, True
  
  RaiseEvent OwnerDrawOnTop(CC, dx, dy)
  W.RaiseBubblingEvent Me, "OwnerDrawOnTop", CC, dx, dy
End Sub
 
