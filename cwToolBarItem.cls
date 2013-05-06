VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwToolBarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Click()
Event ArrowClick()

Public Caption As String, IsCheckable As Boolean, Checked As Boolean
Public ArrowType As DropDownArrowType

Private BDown As Boolean, OutSide As Boolean, MOverArrow As Boolean

Private dx As Single, dy As Single, Alpha As Single
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
      W.CanGetFocus = False
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 1 Then BDown = True: W.Refresh
End Sub
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  
  OutSide = (x < 0 Or x > dx Or y < 0 Or y > dy)
  
  MOverArrow = IIf(ArrowType > 0 And Not OutSide And x > dx - 16, True, False)
  
  W.Refresh
End Sub
Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If BDown And Not OutSide Then CheckAndRaiseClick
  BDown = False: OutSide = False: W.Refresh
End Sub

Private Sub CheckAndRaiseClick()
  If MOverArrow Then
    RaiseEvent ArrowClick
    W.RaiseBubblingEvent Me, "ArrowClick"
  Else
    If IsCheckable Then Checked = Not Checked
    RaiseEvent Click
    W.RaiseBubblingEvent Me, "Click"
  End If
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As vbRichClient5.cWidgetBase)
  W.Refresh
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As vbRichClient5.cWidgetBase)
  MOverArrow = False
  W.Refresh
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned - 2
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim Srf As cCairoSurface, Pat As cCairoPattern, State As enmThemeDrawingState
Dim TxtWidth As Double, FontHeight As Double, xOffs As Single, IcoOffs As Long
Const ArrowSize& = 8
 
  If Caption = "-" Then
      W.Alpha = 0.85
        Cairo.Theme.DrawTo CC, W, thmTypeSeparatorLine, 0, 2, 1, 1, dy
      W.Alpha = Alpha
      Exit Sub
  End If
  
  IcoOffs = IIf(Len(Caption), 5, 3)
  
  W.SelectFontSettingsInto CC

  If W.Enabled Then
    W.SelectFontSettingsInto CC
    CC.SetLineWidth IIf(BDown Or Checked, 1.1, 1)
    
    If BDown Or Checked Then
      If MOverArrow Then CC.ClipExplicit dx - 16, 0, dx, dy
        CC.RoundedRect 0, 0, dx, dy, 2.2, True
        CC.SetSourceColor W.BackColor, Alpha, 0.9
        CC.Fill
      If MOverArrow Then CC.ResetClip
    End If
    
    If W.MouseOver Or Checked Then
      CC.RoundedRect 0, 0, dx, dy, 2.2, True
      Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
          Pat.AddColorStop 0, vbWhite, , IIf(BDown Or Checked, 0.48, 0.82)
          Pat.AddColorStop 1, vbWhite, , IIf(BDown Or Checked, 0.61, 0.48)
      CC.Stroke , Pat
      CC.DrawLine 2, dy, dx - 3, dy, True, 1, IIf(BDown Or Checked, vbWhite, vbBlack), 0.2

      If ArrowType <> ddNone And W.MouseOver Then
        W.Alpha = 0.85
          Cairo.Theme.DrawTo CC, W, thmTypeSeparatorLine, 0, dx - 16, 2, 1, dy - 3
        W.Alpha = Alpha
      End If
    End If
    
    If Len(W.ImageKey) Then
      CC.RenderSurfaceContent W.ImageKey, IcoOffs, 2.5, 19, 19, CAIRO_FILTER_BEST, W.Alpha, True
      xOffs = dy
    End If
    
    If Len(Caption) Then CC.TextOut 5 + xOffs, 4, Caption

  Else 'disabled state
    Set Srf = Cairo.ImageList(W.ImageKey).CreateSimilar(, , , True)
    Srf.AdjustColors -88, 100
    Srf.AdjustColors 100
    Srf.AdjustColors 100
    CC.RenderSurfaceContent Srf, IcoOffs, 2.5, 19, 19, CAIRO_FILTER_BEST, 0.3, True
    
    If Len(Caption) Then CC.TextOut 5 + xOffs, 4, Caption
  End If
  
  If ArrowType Then
    If MOverArrow Then State = thmStateHovered
    If ArrowType = ddCrumbBar Then
      Cairo.Theme.DrawTo CC, W, thmTypeArrow, State, dx - 11.5, (dy - ArrowSize) \ 2, ArrowSize, ArrowSize, 0, IIf(MOverArrow, thmDirectionDown, thmDirectionRight)
    Else
      Cairo.Theme.DrawTo CC, W, thmTypeArrow, State, dx - 11.5, (dy - ArrowSize) \ 2, ArrowSize, ArrowSize, 0, thmDirectionDown
    End If
  End If
End Sub



