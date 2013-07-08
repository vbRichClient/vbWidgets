VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwVScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Event Change()
Event Scroll()
Event MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
Event OwnerDrawScrollArea(ByVal CC As cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, ByVal TopAreaMouseDown As Boolean, ByVal BottomAreaMouseDown As Boolean, CancelBuiltIn As Boolean)
Event OwnerDrawButtons(ByVal CC As cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, ByVal ButtonSize As Single, ByVal TButtonDown As Boolean, ByVal BButtonDown As Boolean, CancelBuiltIn As Boolean)
Event OwnerDrawMButton(ByVal CC As cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, ByVal MButtonDown As Boolean, CancelBuiltIn As Boolean)

Private Const mButtonRadius As Double = 1.5

Private mVal&, mMin&, mMax&, mSmallChange&, mLargeChange&
Private ButtonSize&, MButtonSize&, FreePixels&, ScrollPixels&
Private dx As Single, dy As Single, Alpha As Single
Private TDown As Boolean, BDown As Boolean, mDown As Boolean
Private TLDown As Boolean, BLDown As Boolean
Private TOver As Boolean, BOver As Boolean, MOver As Boolean

Private DragPos As Single, MPos As Single, MOldButtonPos As Single
Private LastPos As Single, LastVal As Long, CurY As Single
Private WithEvents TScroll As cTimer, tScrollCounter As Long
Attribute TScroll.VB_VarHelpID = -1

Public UserdefinedHovering As Boolean

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
  W.CanGetFocus = False
  W.ImplementsWheelMessages = True

  mMax = 9
  mLargeChange = 1
  mSmallChange = 1
 
  W.Move 0, 0, Cairo.Theme.GetScrollerSize, 100
End Sub

Private Sub TScroll_Timer()
  If tScrollCounter < 6 Then
    tScrollCounter = tScrollCounter + 1
  Else
    LastVal = mVal
    If TDown Then If mVal > mMin Then mVal = mVal - mSmallChange: If mVal < mMin Then mVal = mMin
    If TLDown Then If mVal > mMin Then mVal = mVal - mLargeChange: If mVal < mMin Then mVal = mMin
    If BLDown Then If mVal < mMax Then mVal = mVal + mLargeChange: If mVal > mMax Then mVal = mMax
    If BDown Then If mVal < mMax Then mVal = mVal + mSmallChange: If mVal > mMax Then mVal = mMax
    If mVal = mMin Or mVal = mMax Then Set TScroll = Nothing
    W.Refresh
    If LastVal <> mVal Then RaiseEvent Change
  End If
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As cWidgetBase)
  If Not UserdefinedHovering Then W.Refresh
End Sub

Private Sub W_MouseLeave(ByVal MouseEnterWidget As cWidgetBase)
  TOver = False: BOver = False: MOver = False
  
  If Not UserdefinedHovering Then W.Refresh
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button <> 1 Then Exit Sub

  LastPos = y: LastVal = mVal
  Select Case y
    Case Is < ButtonSize: TDown = True 'Top-SmallChange
      If mVal > mMin Then mVal = mVal - mSmallChange: If mVal < mMin Then mVal = mMin
    Case Is < ButtonSize + MButtonPos: TLDown = True 'Top-LargeChange
      If mVal > mMin Then mVal = mVal - mLargeChange: If mVal < mMin Then mVal = mMin
    Case Is < ButtonSize + MButtonPos + MButtonSize 'ScrollButton-Down
      MPos = y: MOldButtonPos = MButtonPos: DragPos = MOldButtonPos: LastVal = mVal: mDown = True: GoTo ex
    Case Is < dy - ButtonSize: BLDown = True 'Bottom-LargeChange
      If mVal < mMax Then mVal = mVal + mLargeChange: If mVal > mMax Then mVal = mMax
    Case Else: BDown = True 'Bottom-SmallChange
      If mVal < mMax Then mVal = mVal + mSmallChange: If mVal > mMax Then mVal = mMax
  End Select
  tScrollCounter = 0
  If TScroll Is Nothing Then Set TScroll = W.Root.CreateTimer(30, True)
ex:
  W.Refresh
  If LastVal <> mVal Then RaiseEvent Change
End Sub

Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim DragY&

  CurY = y
  If mDown Then
    If ScrollPixels = 0 Then Exit Sub
    
    DragY = y - MPos
    If MOldButtonPos + DragY < 0 Then DragY = -MOldButtonPos
    If MOldButtonPos + MButtonSize + DragY > (dy - 2 * ButtonSize) Then DragY = (dy - 2 * ButtonSize) - (MOldButtonPos + MButtonSize)
    mVal = mMin + (mMax - mMin) * (MOldButtonPos + DragY) / ScrollPixels
    
    DragPos = MOldButtonPos + DragY
    W.Refresh

    If mVal <> LastVal Then LastVal = mVal: RaiseEvent Scroll
 
  ElseIf W.MouseOver And Not (TLDown Or BLDown Or TDown Or BDown Or mDown) Then
    TOver = False: BOver = False: MOver = False
    Select Case y
      Case Is < ButtonSize: TOver = True
      Case Is < ButtonSize + MButtonPos
      Case Is < ButtonSize + MButtonPos + MButtonSize: MOver = True
      Case Is < dy - ButtonSize
      Case Else: BOver = True
    End Select
    W.Refresh
  End If
End Sub

Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If mDown Then
    mDown = False
    DragPos = 0
    W.Refresh
    RaiseEvent Change
  Else
    TDown = False: BDown = False: TLDown = False: BLDown = False
    TOver = False: BOver = False: MOver = False
    If Not TScroll Is Nothing Then Set TScroll = Nothing 'W.Root.DestroyTimer(tScroll)
    W.Refresh
  End If
End Sub

Private Sub W_Resize()
  dx = W.Width
  dy = W.Height
  AdjustDimensions
End Sub

Private Sub AdjustDimensions()
  If dx > 20 Then ButtonSize = 20
  If dy < dx + dx Then ButtonSize = (dy \ 2) + 1 Else ButtonSize = dx
  If dy - 2 * ButtonSize < 0 Then FreePixels = 0 Else FreePixels = dy - 2 * ButtonSize
  If FreePixels < 12 Then
    MButtonSize = 0: ScrollPixels = 0
  Else
    MButtonSize = mLargeChange / (Abs(mMax - mMin) + mLargeChange) * FreePixels
    If MButtonSize > FreePixels Then MButtonSize = FreePixels
    If MButtonSize < 13 Then MButtonSize = 13
    ScrollPixels = FreePixels - MButtonSize
  End If
End Sub

Public Property Get Min() As Long
  Min = mMin
End Property
Public Property Let Min(ByVal NewValue As Long)
  If mMin = NewValue Then Exit Property
  mMin = NewValue
  ValueCheck
  If W.LockRefresh Then Exit Property
  AdjustDimensions
End Property

Public Property Get Max() As Long
  Max = mMax
End Property
Public Property Let Max(ByVal NewValue As Long)
  If mMax = NewValue Then Exit Property
  mMax = NewValue
  ValueCheck
  If W.LockRefresh Then Exit Property
  AdjustDimensions
End Property

Public Property Get LargeChange() As Long
  LargeChange = mLargeChange
End Property
Public Property Let LargeChange(ByVal NewValue As Long)
  If NewValue < 1 Then Exit Property
  If mLargeChange = NewValue Then Exit Property
  mLargeChange = NewValue
  AdjustDimensions
End Property

Public Property Get SmallChange() As Long
  SmallChange = mSmallChange
End Property
Public Property Let SmallChange(ByVal NewValue As Long)
  If NewValue < 1 Then Exit Property
  mSmallChange = NewValue
End Property

Public Property Get Value() As Long
  Value = mVal
End Property
Public Property Let Value(ByVal NewValue As Long)
Dim LastVal&
  If mVal = NewValue Then Exit Property
  LastVal = mVal
  mVal = NewValue
  ValueCheck
  
  If W.LockRefresh Then Exit Property

  W.Refresh
  If LastVal <> mVal Then RaiseEvent Change
End Property

Private Sub ValueCheck()
  If mMax < mMin Then
    If mVal < mMax Then mVal = mMax Else If mVal > mMin Then mVal = mMin
  Else
    If mVal < mMin Then mVal = mMin Else If mVal > mMax Then mVal = mMax
  End If
End Sub

Private Sub W_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
  RaiseEvent MouseWheel(MouseKeys, Rotation, LineOffset, xAbs, yAbs)
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  Alpha = W.AlphaInherited
  If dx <> dx_Aligned Or dy <> dy_Aligned Then
    dx = dx_Aligned
    dy = dy_Aligned
    AdjustDimensions
  End If
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim CancelBuiltIn As Boolean

    CancelBuiltIn = False: RaiseEvent OwnerDrawScrollArea(CC, 0, 0, dx, dy, TLDown, BLDown, CancelBuiltIn)
    If Not CancelBuiltIn Then DrawScrollArea CC
    
    CancelBuiltIn = False: RaiseEvent OwnerDrawButtons(CC, 0, 0, dx, dy, ButtonSize, TDown, BDown, CancelBuiltIn)
    If Not CancelBuiltIn Then DrawButtons CC
    
    If MButtonSize > 0 Then
      CancelBuiltIn = False: RaiseEvent OwnerDrawMButton(CC, 0, MButtonPos + ButtonSize, dx, MButtonSize, mDown, CancelBuiltIn)
      If Not CancelBuiltIn Then DrawMButton CC
    End If
End Sub

Private Sub DrawScrollArea(CC As cCairoContext)
Dim Pat As cCairoPattern, Color As Long
  
  If TLDown Or BLDown Then
    W.Alpha = 0.4
      Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, thmStatePressed Or thmStateHovered, _
                         0, ButtonSize, dx, dy - 2 * ButtonSize, mButtonRadius, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeShine, thmStatePressed Or thmStateHovered, _
                         1, ButtonSize + 1, dx - 1, dy - 2 * ButtonSize, mButtonRadius, thmDirectionRight
    W.Alpha = 1
    DrawOuterBorder CC, 0, ButtonSize, dx, dy - 2 * ButtonSize, W.BorderColor, Alpha * 0.6
    DrawOuterBorder CC, 0, ButtonSize, dx, dy - 2 * ButtonSize, W.FocusColor, Alpha * 0.3
  Else
    Color = W.BackColor
    CC.SetLineWidth 1
    Set Pat = Cairo.CreateLinearPattern(0, 0, dx, 0)
      Pat.AddColorStop 0, Color, Alpha * 0.6, 0.98
      Pat.AddColorStop 1, Color, Alpha * 0.6, 1.08
    CC.Rectangle 0, ButtonSize, dx, dy - 2 * ButtonSize, True
    CC.Fill , Pat
    
    
    CC.SetSourceColor Color, Alpha, 1
      CC.DrawLine 0, ButtonSize + 1, 0, dy - ButtonSize - 1
    CC.Stroke

    CC.SetSourceColor Color, Alpha, 1.09
      CC.DrawLine dx - 1, ButtonSize + 1, dx - 1, dy - ButtonSize - 1
    CC.Stroke
  End If
End Sub

Private Sub DrawButtons(CC As cCairoContext)
Dim y As Single, M As Single, ArrowSize As Single
Dim BorderAlpha As Double, State As enmThemeDrawingState
 
  ArrowSize = ButtonSize * 0.48
  M = dx * 0.5
  
  'top Button
  y = 0
  If TDown Then
    BorderAlpha = IIf(mMax <> mMin, 1, 0.1)
    State = thmStateHovered Or thmStatePressed
    If mMax <> mMin Then
      Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, 0, y, dx, ButtonSize, mButtonRadius, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 1, y + 1, dx - 1, ButtonSize - 2, mButtonRadius, thmDirectionRight
    End If
    Cairo.Theme.DrawTo CC, W, thmTypeArrow, State, M - ArrowSize / 2, y + (ButtonSize - ArrowSize) / 2 - 1, ArrowSize, ArrowSize, 0, thmDirectionUp
  Else
    BorderAlpha = IIf(W.MouseOver And mMax <> mMin, 0.7, 0.1)
    State = IIf(TOver, thmStateHovered, 0)
    If mMax <> mMin Then
      Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, 0, y, dx, ButtonSize, mButtonRadius, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 1, y + 1, Int(dx / 2 + 0.1) - 0.25, ButtonSize - 2, mButtonRadius, thmDirectionRight
    End If
    Cairo.Theme.DrawTo CC, W, thmTypeArrow, State, M - ArrowSize / 2, y + (ButtonSize - ArrowSize) / 2 - 1, ArrowSize, ArrowSize, 0, thmDirectionUp
  End If
  DrawOuterBorder CC, 0, y, dx, ButtonSize, W.BorderColor, BorderAlpha
  If TOver Or TDown Then DrawOuterBorder CC, 0, y, dx, ButtonSize, W.FocusColor, BorderAlpha * 0.4

  'bottom Button
  y = dy - ButtonSize
  If BDown Then
    BorderAlpha = IIf(mMax <> mMin, 1, 0.1)
    State = thmStateHovered Or thmStatePressed
    If mMax <> mMin Then
      Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, 0, y, dx, ButtonSize, mButtonRadius, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 1, y + 1, dx - 1, ButtonSize - 2, mButtonRadius, thmDirectionRight
    End If
    Cairo.Theme.DrawTo CC, W, thmTypeArrow, State, M - ArrowSize / 2, y + (ButtonSize - ArrowSize) / 2 + 1, ArrowSize, ArrowSize, 0, thmDirectionDown
  Else
    BorderAlpha = IIf(W.MouseOver And mMax <> mMin, 0.7, 0.1)
    State = IIf(BOver, thmStateHovered, 0)
    If mMax <> mMin Then
      Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, 0, y, dx, ButtonSize, mButtonRadius, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 1, y + 1, Int(dx / 2 + 0.1) - 0.25, ButtonSize - 2, mButtonRadius, thmDirectionRight
    End If
    Cairo.Theme.DrawTo CC, W, thmTypeArrow, State, M - ArrowSize / 2, y + (ButtonSize - ArrowSize) / 2 + 1, ArrowSize, ArrowSize, 0, thmDirectionDown
  End If
  DrawOuterBorder CC, 0, y, dx, ButtonSize, W.BorderColor, BorderAlpha
  If BOver Or BDown Then DrawOuterBorder CC, 0, y, dx, ButtonSize, W.FocusColor, BorderAlpha * 0.4
End Sub
 
Private Sub DrawMButton(CC As cCairoContext)
Dim y As Single, BorderAlpha As Double, State As enmThemeDrawingState

  y = ButtonSize + MButtonPos

  If mMax <> mMin Then
    If mDown Then
      BorderAlpha = Alpha * 1
      State = thmStateHovered Or thmStatePressed
      Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, 0, y, dx, MButtonSize, mButtonRadius, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 1, y + 1, dx - 1, MButtonSize - 2, mButtonRadius, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeScrollBarGripperMarks, State, 0, y, dx, MButtonSize, 0, thmDirectionVert
    Else
      BorderAlpha = Alpha * IIf(W.MouseOver, 0.7, 0.5)
      State = IIf(MOver, thmStateHovered, 0)
      Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, 0, y, dx, MButtonSize, mButtonRadius, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 1, y + 1, Int(dx / 2 + 0.1) - 0.25, MButtonSize - 2, mButtonRadius, thmDirectionRight
      Cairo.Theme.DrawTo CC, W, thmTypeScrollBarGripperMarks, State, 0, y, dx, MButtonSize, 0, thmDirectionVert
    End If
    DrawOuterBorder CC, 0, y, dx, MButtonSize, W.BorderColor, BorderAlpha
    If MOver Or mDown Then DrawOuterBorder CC, 0, y, dx, MButtonSize, W.FocusColor, BorderAlpha * 0.4
  End If
End Sub

Private Sub DrawOuterBorder(CC As cCairoContext, ByVal x As Double, ByVal y As Double, ByVal dx As Double, ByVal dy As Double, ByVal BorderColor As Long, ByVal BorderAlpha As Double)
  CC.SetLineWidth 1
  CC.RoundedRect x, y, dx, dy, mButtonRadius, True
  CC.SetSourceColor BorderColor, BorderAlpha * Alpha
  CC.Stroke
End Sub

Private Function MButtonPos() As Long
  If mMax - mMin = 0 Then Exit Function
  If DragPos Then
    MButtonPos = DragPos
  Else
    MButtonPos = ScrollPixels * (mVal - mMin) / (mMax - mMin)
  End If
End Function
