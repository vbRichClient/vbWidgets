VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwSeparatorLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Public LBL As cwLabel 'we make the wrapped cwLabel also available on the outside

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase 'instantiate the W-WidgetBase which later on connects us to our "hosting environment"
      W.CanGetFocus = False
      W.Alpha = 0.4
      
  Set LBL = Widgets.Add(New cwLabel, "LBL", 0, 0, 1, 27)
      LBL.BorderWidth = 0
      LBL.Widget.BackColor = -1
      LBL.Widget.FontSize = 14
      LBL.Widget.FontItalic = True
      LBL.TextShadowColor = vbBlack
      LBL.TextShadowOffsetX = 0.5
      LBL.TextShadowOffsetY = 1.8
End Sub
 
Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets: Set Widgets = W.Widgets: End Property
 
Public Sub SetCaptionAndImageKey(Caption As String, ImageKey As String, Optional TextColor As Long = &HDDDDDD)
  LBL.Widget.ForeColor = TextColor
  LBL.Caption = Caption
  W.ImageKey = ImageKey
  W.Refresh
End Sub
 
Private Sub W_ContainerResize()
  W.Move W.Left, W.Top, W.Parent.ScaleWidth - 2 * W.Left
End Sub

Private Sub W_Resize()
  LBL.Widget.Move 38, 0, W.ScaleWidth - 38, W.ScaleHeight
End Sub
 
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  CC.DrawLine 0, dy_Aligned - 2, dx_Aligned, dy_Aligned - 2, True, 1, vbBlack, W.Alpha
  CC.DrawLine 0, dy_Aligned - 1, dx_Aligned, dy_Aligned - 1, True, 1, &HDDDDDD, W.Alpha * 0.5
  CC.Operator = CAIRO_OPERATOR_HARD_LIGHT ' CAIRO_OPERATOR_DARKEN
  CC.RenderSurfaceContent W.ImageKey, 1, (dy_Aligned - 32) \ 2, 32, 32
End Sub
 
 
