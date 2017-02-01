VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'a more efficiently implemented Scrollbar... Author Colin Edwards (December 2016)

Private Enum enumElement
   eNone
   eTopButton
   eBottomButton
   eScrollThumb
   eScrollTrackAbove
   eScrollTrackBelow
End Enum

Private WithEvents W As cWidgetBase, WithEvents mTimer As cTimer
Attribute W.VB_VarHelpID = -1
Attribute mTimer.VB_VarHelpID = -1
Private mMin As Double, mMax As Double, mSmallChange As Double, mLargeChange As Double, mValue As Double, mPreDragValue As Double
Private mScrollThumbY As Double, mScrollThumbH As Double, mButtonSize As Double, mBottomRightEdgeWidth As Long
Private mMouseX As Single, mMouseY As Single, mScrollTrackMouseDownY As Double, mMouseDownElement As enumElement, mMouseOverElement As enumElement
Private mVertical As Boolean, mWidth As Double, mHeight As Double, mYScale As Double, mAlpha As Single
 
Event Change(Value As Double) 'both events are only raised via user interaction...
Event Scroll(ExactValue As Double) '...setting the .Value property in code does not cause a change event! This is intentional.
Event MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)

Private Sub Class_Initialize()
  Set mTimer = New_c.Timer()
  Set W = Cairo.WidgetBase
      W.CanGetFocus = False
      W.ImplementsWheelMessages = True
  mMax = 9
  mLargeChange = 1
  mSmallChange = 1
  mVertical = True
End Sub

Private Sub Class_Terminate()
   Set mTimer = Nothing
End Sub

Public Property Get Widget() As cWidgetBase
   Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
   Set Widgets = W.Widgets
End Property
 
Public Sub SetParams(ByVal pValue As Double, ByVal pMin As Double, ByVal pMax As Double, ByVal pSmallChange As Double, ByVal pLargeChange As Double)
   mMin = pMin: mMax = IIf(pMax < pMin, pMin, pMax)
   mSmallChange = IIf(pSmallChange > mMax - mMin, 1, pSmallChange)
   mLargeChange = pLargeChange
   mValue = IIf(pValue < mMin Or pValue > mMax, mMin, pValue)
   W_Resize
End Sub

Public Property Let Vertical(ByVal RHS As Boolean)
  mVertical = RHS
End Property
Public Property Let Value(ByVal pValue As Double)
   If Not mMouseDownElement = eScrollThumb Then UpdateValue pValue, True  'don't allow the internal value to be changed from outside while thumb is being dragged
End Property
Public Property Let Min(ByVal pMin As Double)
   SetParams mValue, pMin, mMax, mSmallChange, mLargeChange
End Property
Public Property Let Max(ByVal pMax As Double)
   SetParams mValue, mMin, pMax, mSmallChange, mLargeChange
End Property
Public Property Let SmallChange(ByVal pSmallChange As Double)
   SetParams mValue, mMin, mMax, pSmallChange, mLargeChange
End Property
Public Property Let LargeChange(ByVal pLargeChange As Double)
   SetParams mValue, mMin, mMax, mSmallChange, pLargeChange
End Property
Public Property Let BottomRightEdgeWidth(ByVal NewValue As Long)
  If mBottomRightEdgeWidth = NewValue Then Exit Property Else mBottomRightEdgeWidth = NewValue
  W_Resize
End Property
Public Property Get BottomRightEdgeWidth() As Long
  BottomRightEdgeWidth = mBottomRightEdgeWidth
End Property

Public Property Get Vertical() As Boolean: Vertical = mVertical: End Property
Public Property Get Value() As Double: Value = mValue: End Property
Public Property Get Min() As Double: Min = mMin: End Property
Public Property Get Max() As Double: Max = mMax: End Property
Public Property Get SmallChange() As Double: SmallChange = mSmallChange: End Property
Public Property Get LargeChange() As Double: LargeChange = mLargeChange: End Property

Private Sub UpdateValue(ByVal pNewValue As Double, Optional ByVal NotifyChange As Boolean, Optional ByVal ForceNotify As Boolean)
   If pNewValue < mMin Then pNewValue = mMin
   If pNewValue > mMax Then pNewValue = mMax
   If mValue = pNewValue And Not ForceNotify Then Exit Sub
   mValue = pNewValue
   If W.LockRefresh Then Exit Sub
   W.Refresh
   If NotifyChange Then RaiseEvent Change(pNewValue)
End Sub
Private Sub Draw(CC As cCairoContext)
   mAlpha = W.AlphaInherited
   If Not mVertical Then
      CC.ScaleDrawings 1, -1
      CC.RotateDrawingsDeg -90
   End If
 
   mScrollThumbY = mWidth + (mValue - mMin) * mYScale
   DrawScrollTrack CC
   DrawScrollButton CC, eTopButton, 0
   DrawScrollButton CC, eBottomButton, mHeight - mButtonSize
   DrawScrollThumb CC
   If mBottomRightEdgeWidth Then
      Cairo.Theme.DrawTo CC, W, thmTypeBottomRightEdge, 0, 0, mHeight, mWidth, mBottomRightEdgeWidth
   End If
End Sub
Private Sub DrawScrollTrack(CC As cCairoContext)
Dim Alpha As Double
   CC.SetSourceColor W.BackColor, mAlpha * 0.6: CC.Paint
   If mMouseDownElement > 3 Then
      Alpha = W.Alpha: W.Alpha = 0.4
      Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, thmStatePressed Or thmStateHovered, 0, mScrollThumbY + mScrollThumbH / 2, mWidth, IIf(mMouseDownElement = eScrollTrackAbove, -mHeight, mHeight), 1.5, thmDirectionRight
      W.Alpha = Alpha
   End If
   CC.SetLineWidth 1
   CC.Rectangle 0, 0, mWidth, mHeight, True
   CC.SetSourceColor W.BorderColor, 0.15
   CC.Stroke
End Sub
Private Sub DrawScrollButton(CC As cCairoContext, Element As enumElement, y As Double)
Dim DrawState As enmThemeDrawingState, ArrowSize As Double
   ArrowSize = mButtonSize * 0.5
   DrawState = ElementDrawState(Element)
   If mMouseOverElement > eNone Or mMouseDownElement > eNone Then
      Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, DrawState, 0, y, mWidth - 0.7, mWidth, 1.5, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeShine, DrawState, 0, y + 1, mWidth - 1.5, mWidth - 2, 1.5, thmDirectionRight
      DrawOuterBorder CC, 0, y, mWidth, mButtonSize, W.BorderColor, mAlpha * IIf(DrawState, 0.7, 0.5)
      DrawOuterBorder CC, 0, y, mWidth, mButtonSize, IIf(DrawState, W.FocusColor, W.BorderColor), 0.5
   End If
   Cairo.Theme.DrawTo CC, W, thmTypeArrow, DrawState, (mWidth - ArrowSize) * 0.5, y + (mButtonSize - ArrowSize) * 0.5 + IIf(y, 0.75, -0.75), ArrowSize, ArrowSize, 0, IIf(Element = eTopButton, thmDirectionUp, thmDirectionDown)
End Sub
Private Sub DrawScrollThumb(CC As cCairoContext)
Dim DrawState As enmThemeDrawingState
   If mScrollThumbH = 0 Then Exit Sub
   DrawState = ElementDrawState(eScrollThumb)
   CC.SetLineWidth 1
   Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, DrawState, 0, mScrollThumbY, mWidth - 0.7, mScrollThumbH, 1.5, thmDirectionRight
   Cairo.Theme.DrawTo CC, W, thmTypeShine, DrawState, 0, mScrollThumbY + 1, mWidth - 1.5, mScrollThumbH - 2, 1.5, thmDirectionRight
   Cairo.Theme.DrawTo CC, W, thmTypeScrollBarGripperMarks, IIf(DrawState, thmStateHovered, 0), 0, mScrollThumbY, mWidth, Int(mScrollThumbH), 0, thmDirectionVert
   DrawOuterBorder CC, 0, mScrollThumbY, mWidth, mScrollThumbH, W.BorderColor, mAlpha * IIf(DrawState, 0.7, 0.5)
   DrawOuterBorder CC, 0, mScrollThumbY, mWidth, mScrollThumbH, IIf(DrawState, W.FocusColor, W.BorderColor), 0.3
End Sub
Private Sub DrawOuterBorder(CC As cCairoContext, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal BorderColor As Long, ByVal BorderAlpha As Double)
  CC.SetLineWidth 1
  CC.RoundedRect x, y, dx, dy, 1.5, True
  CC.SetSourceColor BorderColor, BorderAlpha * mAlpha
  CC.Stroke
End Sub
Private Function ElementDrawState(pElement As enumElement) As enmThemeDrawingState
   If mMouseDownElement <> pElement And mMouseDownElement <> eNone Then Exit Function
   If mMouseDownElement = pElement Then ElementDrawState = thmStateHovered Or thmStatePressed Else If mMouseOverElement = pElement Then ElementDrawState = thmStateHovered
End Function
Private Sub W_MouseLeave(ByVal MouseEnterWidget As vbRichClient5.cWidgetBase)
   mTimer.Enabled = Not mMouseDownElement = eNone
   If Not mTimer.Enabled Then mMouseOverElement = eNone: W.Refresh
End Sub
Private Sub W_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
   UpdateValue mValue - mSmallChange * Sgn(Rotation), True
   RaiseEvent MouseWheel(MouseKeys, Rotation, LineOffset, xAbs, yAbs)
End Sub

Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
   Draw CC
End Sub
Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
   If Not Button = vbLeftButton Then Exit Sub
   If mVertical Then mMouseX = x: mMouseY = y Else mMouseX = -y + mWidth: mMouseY = x
   mMouseDownElement = HitTest(mMouseX, mMouseY)
   W.Refresh
   If mMouseDownElement = eScrollThumb Then
      mScrollTrackMouseDownY = mMouseY - mScrollThumbY
      mPreDragValue = mValue
   End If
   mTimer_Timer
   mTimer.Enabled = True
End Sub
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim PrevMouseOverElement As enumElement
   If mVertical Then mMouseX = x: mMouseY = y Else mMouseX = -y + mWidth: mMouseY = x
   If mMouseDownElement = eNone Then mMouseOverElement = HitTest(mMouseX, mMouseY)
   If mMouseOverElement <> PrevMouseOverElement Then PrevMouseOverElement = mMouseOverElement: W.Refresh
   If Button = vbLeftButton And mMouseDownElement = eScrollThumb Then
      If mMouseX > -140 And mMouseX < mWidth + 140 Then
         UpdateValue mValue + ((mMouseY - mScrollTrackMouseDownY) / mYScale) - (mScrollThumbY / mYScale)
      Else
         UpdateValue mPreDragValue 'allows the thumb to 'snap back' to original value when mouse is moved too far away from the scrollbar
      End If
      RaiseEvent Scroll(mValue)
   Else
      If Not Button = vbLeftButton Then mMouseDownElement = eNone 'belt and braces!
   End If
End Sub
Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim ForceChangeEvent As Boolean
   mTimer.Enabled = False
   ForceChangeEvent = mMouseDownElement = eScrollThumb And mPreDragValue <> Round(mValue)
   mMouseDownElement = eNone
   W.Refresh
   UpdateValue Round(mValue), ForceChangeEvent, ForceChangeEvent
End Sub
Private Sub W_Resize()
   If mVertical Then
      mWidth = W.Width: mHeight = W.Height
   Else
      mWidth = W.Height: mHeight = W.Width - mBottomRightEdgeWidth
   End If
   mButtonSize = IIf(mWidth * 2 < mHeight, mWidth, mHeight \ 2 + 1)
   If mMin <> mMax Then
      If mLargeChange < 1 Then mLargeChange = 1
      mScrollThumbH = ((mHeight - mWidth * 2) / (mMax - mMin + mLargeChange)) * mLargeChange
      If mScrollThumbH < Int(mWidth * 0.7) Then mScrollThumbH = Int(mWidth * 0.7)
      If 3 * mWidth > mHeight Then mScrollThumbH = 0
      mYScale = (mHeight - mScrollThumbH - mWidth * 2) / (mMax - mMin)
   Else
      mScrollThumbH = (mHeight - mWidth * 2)
      If mScrollThumbH < 10 Then mScrollThumbH = 0
   End If
   W.Refresh
End Sub
Private Function HitTest(x As Single, y As Single) As enumElement
   Select Case True
      Case y <= mButtonSize And x > 0 And x < mWidth: HitTest = eTopButton
      Case y >= mHeight - mButtonSize And x > 0 And x < mWidth: HitTest = eBottomButton
      Case y >= mScrollThumbY And y <= mScrollThumbY + mScrollThumbH And mScrollThumbH > 0: HitTest = eScrollThumb
      Case y < mScrollThumbY And x > 0 And x < mWidth: HitTest = eScrollTrackAbove
      Case y > mScrollThumbY + mScrollThumbH And x > 0 And x < mWidth: HitTest = eScrollTrackBelow
      Case Else: HitTest = eNone
   End Select
End Function
Private Sub mTimer_Timer()
   mMouseOverElement = HitTest(mMouseX, mMouseY) 'need to check this as it can change even when the mouse is still (i.e. the thumb moving)
   If Not mMouseOverElement = mMouseDownElement Then
      W.Refresh
   Else
      Select Case mMouseDownElement
         Case eTopButton: UpdateValue mValue - mSmallChange, True
         Case eBottomButton: UpdateValue mValue + mSmallChange, True
         Case eScrollTrackAbove: UpdateValue mValue - mLargeChange, True
         Case eScrollTrackBelow: UpdateValue mValue + mLargeChange, True
      End Select
   End If
   mTimer.Interval = IIf(mTimer.Enabled, 50, 500)  'handle a repeat delay
End Sub
