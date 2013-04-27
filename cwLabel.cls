VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCaption As String
Private mWordWrap As Boolean
Private mAlignment As AlignmentConstants, mInnerSpace As Long, mVAlign As Long
Private mBorderWidth As Single, mBorderRadius As Single
Private mTextShadowOffsetX As Single, mTextShadowOffsetY As Single, mTextShadowColor As Long

Private dx As Single, dy As Single, Alpha As Single

'****---- Start of cwImplementation-Conventions ----****
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
      W.CanGetFocus = False
  mTextShadowColor = -1 '-1 Color-Values are (as always) treated as "no color"
  mAlignment = vbLeftJustify
  mVAlign = 1
  mInnerSpace = 1 '1 Pixel Offset for the Text-Rendering - independent from (and in addition to) the BorderSize)
  mBorderWidth = 1
End Sub
'****---- End of cwImplementation-Conventions ----****


'OK, let's implement our small Label-Widget, starting with its Caption-Property
Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
  If mCaption = NewValue Then Exit Property
  mCaption = NewValue
  W.Refresh
End Property

Public Property Get Alignment() As AlignmentConstants
  Alignment = mAlignment
End Property
Public Property Let Alignment(ByVal NewValue As AlignmentConstants)
  If mAlignment = NewValue Then Exit Property
  mAlignment = NewValue
  W.Refresh
End Property

Public Property Get VAlign() As Long
  VAlign = mVAlign
End Property
Public Property Let VAlign(ByVal NewValue As Long)
  If mVAlign = NewValue Then Exit Property
  mVAlign = NewValue
  W.Refresh
End Property

Public Property Get BorderWidth() As Single
  BorderWidth = mBorderWidth
End Property
Public Property Let BorderWidth(ByVal NewValue As Single)
  If NewValue = mBorderWidth Then Exit Property
  mBorderWidth = NewValue
  W.SetClientAreaOffsets mBorderWidth, mBorderWidth, mBorderWidth, mBorderWidth
  W.Refresh
End Property

Public Property Get BorderRadius() As Single
  BorderRadius = mBorderRadius
End Property
Public Property Let BorderRadius(ByVal NewValue As Single)
  If mBorderRadius = NewValue Then Exit Property
  mBorderRadius = NewValue
  W.Refresh
End Property

Public Property Get InnerSpace() As Long
  InnerSpace = mInnerSpace
End Property
Public Property Let InnerSpace(ByVal NewValue As Long)
  If mInnerSpace = NewValue Then Exit Property
  mInnerSpace = NewValue
  W.Refresh
End Property

Public Property Get WordWrap() As Boolean
  WordWrap = mWordWrap
End Property
Public Property Let WordWrap(ByVal NewValue As Boolean)
  If mWordWrap = NewValue Then Exit Property
  mWordWrap = NewValue
  W.Refresh
End Property

Public Property Get TextShadowOffsetX() As Single
  TextShadowOffsetX = mTextShadowOffsetX
End Property
Public Property Let TextShadowOffsetX(ByVal NewValue As Single)
  If mTextShadowOffsetX = NewValue Then Exit Property
  mTextShadowOffsetX = NewValue
  W.Refresh
End Property
Public Property Get TextShadowOffsetY() As Single
  TextShadowOffsetY = mTextShadowOffsetY
End Property
Public Property Let TextShadowOffsetY(ByVal NewValue As Single)
  If mTextShadowOffsetY = NewValue Then Exit Property
  mTextShadowOffsetY = NewValue
  W.Refresh
End Property
Public Property Get TextShadowColor() As Long
  TextShadowColor = mTextShadowColor
End Property
Public Property Let TextShadowColor(ByVal NewValue As Long)
  If mTextShadowColor = NewValue Then Exit Property
  mTextShadowColor = NewValue
  W.Refresh
End Property


Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  Alpha = W.AlphaInherited
  dx = dx_Aligned
  dy = dy_Aligned
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim i&
 
  If W.BackColor <> -1 Then 'draw a plain solid background
    CC.RoundedRect 0, 0, dx, dy, mBorderRadius
    CC.SetSourceColor W.BackColor, Alpha
    CC.Fill
  End If
  
  If W.BorderColor <> -1 And mBorderWidth > 0 Then
    CC.SetLineWidth mBorderWidth, True
    Cairo.Theme.DrawTo CC, W, thmTypeBorder, 0, 0, 0, dx, dy, mBorderRadius
  End If
  
  'and here the Text-Out
  W.SelectFontSettingsInto CC
  
  If mTextShadowColor <> -1 Then
    'note the last Param, which only ensures "a Path" (no Pixel-Rendering yet)
    CC.Save
      CC.TranslateDrawings mTextShadowOffsetX, mTextShadowOffsetY
      CC.TranslateDrawings 0.5, 0.5
      CC.DrawText mBorderWidth, mBorderWidth, dx, dy, mCaption, Not mWordWrap, mAlignment, mInnerSpace, mVAlign, , , True
      CC.SetLineJoin CAIRO_LINE_JOIN_ROUND
      CC.SetLineCap CAIRO_LINE_CAP_ROUND
      
      For i = 0 To 1 'this loop re-renders always on the same path (ensured by the True-Param in the .Stroke)
        CC.SetLineWidth 4 - 2 * i 'but with decreasing linewidth
        CC.SetSourceColor mTextShadowColor, 0.15 + i * 0.05  'and slightly increasing Opacity
        CC.Stroke True
      Next i
      
      CC.SetLineWidth 1
      CC.SetSourceColor mTextShadowColor, 0.2
      CC.Stroke
    CC.Restore
  End If

  CC.DrawText mBorderWidth, mBorderWidth, dx, dy, mCaption, Not mWordWrap, mAlignment, mInnerSpace, mVAlign
End Sub

