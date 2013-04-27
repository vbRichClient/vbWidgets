VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwScoringLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Alpha As Single, dx As Single, dy As Single
Private mCaptionTop As String
Private mCaptionBottom As String
Private mFractionLineWidth As Long
Private mSingleLine As Boolean
Private mInnerSpace As Long
Private mBorderWidth As Long

'****---- Start of cwImplementation-Conventions ----****
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private Sub Class_Initialize()

    Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
    
    'some default-inits on our Widget-internal, local 'm' Variables
    mBorderWidth = 1
    mFractionLineWidth = 3
    mInnerSpace = 1 '1 Pixel Offset for the Text-Rendering - independent from (and additionally to) the BorderSize)
    mCaptionTop = "0"
    mCaptionBottom = "0"
    
    '...and the following "W-Defaults-adaptions" are done individually (as needed for the control in question)
    W.SetClientAreaOffsets mBorderWidth, mBorderWidth, mBorderWidth, mBorderWidth '<- this defines a widget-client-area
    W.CanGetFocus = False 'this way the Label will not be respected in the "Focus-Switch-Chain"
    
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property
'****---- End of cwImplementation-Conventions ----****


Public Property Get CaptionTop() As String
  CaptionTop = mCaptionTop
End Property
Public Property Let CaptionTop(ByVal NewValue As String)
  If mCaptionTop = NewValue Then Exit Property
  mCaptionTop = NewValue
  W.Refresh 'a change of the Caption will require a Redraw, so let's signal that over W
End Property

Public Property Get CaptionBottom() As String
  CaptionBottom = mCaptionBottom
End Property
Public Property Let CaptionBottom(ByVal NewValue As String)
  If mCaptionBottom = NewValue Then Exit Property
  mCaptionBottom = NewValue
  W.Refresh 'a change of the Caption will require a Redraw, so let's signal that over W
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = mBorderWidth
End Property
Public Property Let BorderWidth(ByVal NewValue As Long)
  If mBorderWidth = NewValue Then Exit Property
  mBorderWidth = NewValue
  W.SetClientAreaOffsets mBorderWidth, mBorderWidth, mBorderWidth, mBorderWidth
  W.Refresh 'a change of the BorderWidth will require a Redraw, so let's signal that over W
End Property

Public Property Get FractionLineWidth() As Long
  FractionLineWidth = mFractionLineWidth
End Property
Public Property Let FractionLineWidth(ByVal NewValue As Long)
  If mFractionLineWidth = NewValue Then Exit Property
  mFractionLineWidth = NewValue
  W.Refresh 'a change of the BorderWidth will require a Redraw, so let's signal that over W
End Property

Public Property Get InnerSpace() As Long
  InnerSpace = mInnerSpace
End Property
Public Property Let InnerSpace(ByVal NewValue As Long)
  If mInnerSpace = NewValue Then Exit Property
  mInnerSpace = NewValue
  W.Refresh 'a change of the InnerSpace will require a Redraw, so let's signal that over W
End Property

Public Property Get SingleLine() As Boolean
  SingleLine = mSingleLine
End Property
Public Property Let SingleLine(ByVal NewValue As Boolean)
  If mSingleLine = NewValue Then Exit Property
  mSingleLine = NewValue
  W.Refresh 'a change of the SingleLine-Mode will require a Redraw, so let's signal that over W
End Property
 
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  Alpha = W.AlphaInherited
  dx = dx_Aligned
  dy = dy_Aligned

  Draw CC, W.AbsLeft - xAbs, W.AbsTop - yAbs 'here we just delegate to our internal Drawing-Routine...
End Sub

Private Sub Draw(CC As cCairoContext, ByVal x As Single, ByVal y As Single)
    Dim captionTopWidth As Double
    Dim captionBottomWidth As Double
    Dim mFontHeight As Double
  
    CC.SetLineWidth mBorderWidth
    If W.BackColor <> -1 Then
        CC.Rectangle x, y, W.Width, W.Height, True '<-- note the last optional Param, which ensures a Pixel-aligned drawing inside the Widgets-Bounds
        CC.SetSourceColor W.BackColor, W.Alpha
        CC.Fill
    End If
    If W.BorderColor <> -1 Then
        CC.Rectangle x, y, W.Width, W.Height, True '<-- note the last optional Param, which ensures a Pixel-aligned drawing inside the Widgets-Bounds
        CC.SetSourceColor W.BorderColor, W.Alpha
        CC.Stroke
    End If
  
  CC.SelectFont W.FontName, W.FontSize, W.ForeColor, W.FontBold, W.FontItalic
  captionTopWidth = CC.GetTextExtents(mCaptionTop)
  captionBottomWidth = CC.GetTextExtents(mCaptionBottom)
  mFontHeight = CC.GetFontHeight
  
  CC.SetSourceColor W.ForeColor
  If captionTopWidth > captionBottomWidth Then
    CC.RoundedRect _
      (W.ScaleWidth - captionTopWidth) / 2, _
      (W.ScaleHeight - mFractionLineWidth) / 2, _
      captionTopWidth, _
      mFractionLineWidth, 5, True
  Else
    CC.RoundedRect _
      (W.ScaleWidth - captionBottomWidth) / 2, _
      (W.ScaleHeight - mFractionLineWidth) / 2, _
      captionBottomWidth, _
      mFractionLineWidth, 5, True
    
  End If
  CC.Fill
    
  CC.DrawText 0, (W.ScaleHeight - mFractionLineWidth) / 2 - mFontHeight, W.ScaleWidth, W.ScaleHeight / 2, mCaptionTop, mSingleLine, vbCenter, mInnerSpace, 0
  CC.DrawText 0, (W.ScaleHeight + mFractionLineWidth) / 2, W.ScaleWidth, W.ScaleHeight / 2, mCaptionBottom, mSingleLine, vbCenter, mInnerSpace, 0
End Sub

 
