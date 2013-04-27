VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private dx As Single, dy As Single, Alpha As Single
Private mCaption As String, mCaptionHeight As Single
Private mBorderRadius  As Single, mBorderWidth  As Single

Public UserdefinedHovering As Boolean, RoundedTopOnly As Boolean
Public OnActivateMoveToFront As Boolean

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Private Sub Class_Initialize()
  mBorderRadius = 4
  mBorderWidth = 1
  Set W = Cairo.WidgetBase
  W.SetClientAreaOffsets mBorderWidth, mBorderWidth, mBorderWidth, mBorderWidth
End Sub

Private Sub W_EnterFocus()
  If OnActivateMoveToFront Then
    W.MoveToFront
    W.Refresh
    'W.SetFocus
  End If
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  Alpha = W.AlphaInherited
  dx = dx_Aligned
  dy = dy_Aligned
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim Pat As cCairoPattern, BColor As Long, BWHalf As Single, LWAlign As Single, Zoom As Single
Dim TextWidth As Single, FontHeight As Double, yy As Single, dyy As Single
 
  BWHalf = mBorderWidth / 2
  Zoom = W.Zoom
  'first we determine the Caption-Extents (if there is one)
  If Len(mCaption) Then
    W.SelectFontSettingsInto CC
    TextWidth = CC.GetTextExtents(mCaption, FontHeight)
  End If
  
  BColor = W.BorderColor
  If W.Active And Me.OnActivateMoveToFront Then BColor = W.FocusColor
  If W.Focused Then BColor = W.FocusColor
  
  
  
  'now the Outline, the layout depending on, if we have a caption or not
  If Len(mCaption) = 0 Or mCaptionHeight <> 0 Then 'completely "closed" Outline
    CC.SetSourceColor BColor, Alpha
      CC.SetLineWidth mBorderWidth, True
      
      If RoundedTopOnly Then
        CC.RoundedRect 0, 0, dx, dy, mBorderRadius, True, cmTop
      Else
        CC.RoundedRect 0, 0, dx, dy, mBorderRadius, True, cmAll
      End If
    CC.Stroke
    
  ElseIf Len(mCaption) Then 'leave room for the caption
    CC.SetSourceColor BColor, Alpha
      
      CC.SetLineWidth mBorderWidth, True
      yy = CLng(FontHeight * 0.5 * W.Zoom) / W.Zoom
      dyy = CLng((dy - yy - 1) * W.Zoom) / W.Zoom
      LWAlign = CC.GetLineWidth
      
      'Now we draw only the Corners first (without the lines) and then the lines between Corners
      If RoundedTopOnly Then
        If mBorderRadius > 0 Then CC.RoundedRect 0, yy, dx, dyy, mBorderRadius, True, cmTop, True
        'now the lines between the Corners
        DrawAlignedLine CC, dx - LWAlign, yy + mBorderRadius, dx - LWAlign, yy + dyy - LWAlign, Zoom
        DrawAlignedLine CC, dx - LWAlign, yy + dyy - LWAlign, 0, yy + dyy - LWAlign, Zoom
        DrawAlignedLine CC, 0, yy + dyy - LWAlign, 0, yy + BorderRadius, Zoom
      Else
        If mBorderRadius > 0 Then CC.RoundedRect 0, yy, dx, dyy, mBorderRadius, True, cmAll, True
        'now the lines between the Corners
        DrawAlignedLine CC, dx - LWAlign, yy + mBorderRadius, dx - LWAlign, yy + dyy - mBorderRadius - LWAlign, Zoom
        DrawAlignedLine CC, dx - LWAlign - mBorderRadius, yy + dyy - LWAlign, mBorderRadius, yy + dyy - LWAlign, Zoom
        DrawAlignedLine CC, 0, yy + dyy - LWAlign - mBorderRadius, 0, yy + BorderRadius, Zoom
      End If
      DrawAlignedLine CC, mBorderRadius, yy, mBorderRadius + 2, yy, Zoom
      DrawAlignedLine CC, 5 + mBorderRadius + TextWidth + 2, yy, dx - mBorderRadius - LWAlign, yy, Zoom
      
    CC.Stroke
    
  End If

  If mCaptionHeight = 0 Then 'fill BackGround completely
    If W.BackColor <> -1 Then
      CC.SetSourceColor W.BackColor, Alpha
        CC.SetLineWidth mBorderWidth + mBorderWidth
        If Len(Caption) Then
          W.SetClientAreaOffsets mBorderWidth, mBorderWidth + yy, mBorderWidth, mBorderWidth
          CC.RoundedRect 0, yy, dx, dyy, mBorderRadius, True
        Else
          W.SetClientAreaOffsets mBorderWidth, mBorderWidth, mBorderWidth, mBorderWidth
          CC.RoundedRect 0, 0, dx, dy, mBorderRadius, True
        End If
      CC.Fill
    End If
    
    CC.TextOut 5 + mBorderRadius, 0, mCaption     '<-Cairo-text is always rendered at the BaseLine

  Else 'draw a caption first, followed by the normal BackGround-Fill of the rest
    'the caption-bar
    Set Pat = Cairo.CreateLinearPattern(0, 0, 0, mCaptionHeight + mBorderWidth)
      Pat.AddColorStop 0, W.ForeColor, Alpha, 1.25
      Pat.AddColorStop 0.2, W.ForeColor, Alpha, 0.9
      Pat.AddColorStop 0.88, W.ForeColor, Alpha, 1.05
      Pat.AddColorStop 1, W.ForeColor, Alpha, 0.7
      
      CC.RoundedRect BWHalf, BWHalf, dx - mBorderWidth, mCaptionHeight + mBorderWidth, mBorderRadius - 1, True, cmTop
    CC.Fill , Pat
    
    'now render the caption-string vertically centered
    CC.TextOut 6, mBorderWidth + (mCaptionHeight - FontHeight) * 0.5, mCaption       '<-Cairo-text is always rendered at the BaseLine
    
    If W.BackColor <> -1 Then 'and the rest of the filling in BackColor, following the caption
      CC.SetSourceColor W.BackColor, Alpha
        CC.SetLineWidth mBorderWidth + mBorderWidth
        If RoundedTopOnly Then
          CC.RoundedRect 0, mCaptionHeight - 0.1, dx, dy - mCaptionHeight + 0.2, mBorderRadius, True, cmNone
        Else
          CC.RoundedRect 0, mCaptionHeight - 0.1, dx, dy - mCaptionHeight + 0.2, mBorderRadius, True, cmBottom
        End If
      CC.Fill
    End If
  End If
End Sub

Private Sub DrawAlignedLine(CC As cCairoContext, ByVal x1!, ByVal y1!, ByVal x2!, ByVal y2!, ByVal ZoomFac As Single)
  x1 = CLng(x1 * ZoomFac) / ZoomFac
  y1 = CLng(y1 * ZoomFac) / ZoomFac
  x2 = CLng(x2 * ZoomFac) / ZoomFac
  y2 = CLng(y2 * ZoomFac) / ZoomFac

  CC.DrawLine x1, y1, x2, y2, True
End Sub

Public Property Get Caption$()
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewVal$)
  mCaption = NewVal
  W.Refresh
End Property

Public Property Get CaptionColor&()
  CaptionColor = W.ForeColor
End Property
Public Property Let CaptionColor(ByVal NewVal&)
  W.ForeColor = NewVal
End Property

Public Property Get BorderWidth() As Single
  BorderWidth = mBorderWidth
End Property
Public Property Let BorderWidth(ByVal NewVal As Single)
  If NewVal > 10 Then NewVal = 10
  If NewVal < 0.01 Then NewVal = 0.01
  If NewVal = mBorderWidth Then Exit Property
  mBorderWidth = NewVal
  W.SetClientAreaOffsets mBorderWidth, mBorderWidth + mCaptionHeight, mBorderWidth, mBorderWidth
End Property

Public Property Get CaptionHeight() As Single
  CaptionHeight = mCaptionHeight
End Property
Public Property Let CaptionHeight(ByVal NewVal As Single)
  If NewVal > 50 Then NewVal = 50
  If NewVal < 0 Then NewVal = 0
  If NewVal = mCaptionHeight Then Exit Property
  mCaptionHeight = NewVal
  W.SetClientAreaOffsets mBorderWidth, mBorderWidth + mCaptionHeight, mBorderWidth, mBorderWidth
End Property

Public Property Get BorderRadius() As Single
  BorderRadius = mBorderRadius
End Property
Public Property Let BorderRadius(ByVal NewVal As Single)
  If NewVal > 100 Then NewVal = 100
  If NewVal < 0 Then NewVal = 0
  If NewVal = mBorderRadius Then Exit Property
  mBorderRadius = NewVal
End Property

