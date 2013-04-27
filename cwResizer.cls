VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ResizerStyle
  ResizerVerticalRightAligned
  ResizerVerticalLeftAligned
  ResizerHorizontalBottomAligned
  ResizerHorizontalTopAligned
End Enum

Event ResizerMoved(ByVal NewParentLeft As Single, ByVal NewParentTop As Single, ByVal NewParentWidth As Single, ByVal NewParentHeight As Single, ByVal Delta As Single)

Private Alpha As Single, dx As Single, dy As Single
Private mResizerStyle As ResizerStyle, mSizePixels As Single
Private mDownParentLeft As Single, mDownParentTop As Single
Private mDownParentWidth As Single, mDownParentHeight As Single
Private mDownX As Single, mDownY As Single, mDown As Boolean

Public MinWidth As Single, MaxWidth As Single, MinHeight As Single, MaxHeight As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
  
Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
      W.CanGetFocus = False
      
  mSizePixels = 6
  MinWidth = mSizePixels:  MaxWidth = 1000000
  MinHeight = mSizePixels: MaxHeight = 1000000
  
  On Error Resume Next
    With New_c.Crypt
      If Not Cairo.ImageList.Exists("VSplitCursor") Then Cairo.ImageList.AddImage "VSplitCursor", .Base64Dec(VSplitCursor_Png, True)
      If Not Cairo.ImageList.Exists("HSplitCursor") Then Cairo.ImageList.AddImage "HSplitCursor", .Base64Dec(HSplitCursor_Png, True)
    End With
  If Err Then Err.Clear
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Public Property Get SizePixels() As Single
  SizePixels = mSizePixels
End Property
Public Property Let SizePixels(ByVal NewValue As Single)
  If mSizePixels = NewValue Then Exit Property
  mSizePixels = NewValue
  W.Refresh
End Property

Public Sub Init(ByVal ResizerStyle As ResizerStyle, Optional ByVal MinWidth As Single = 6, Optional ByVal MinHeight As Single = 6, _
                                                    Optional ByVal MaxWidth As Single, Optional ByVal MaxHeight As Single)
  Me.ResizerStyle = ResizerStyle
  Me.MinWidth = MinWidth:   If MaxWidth > 0 Then Me.MaxWidth = MaxWidth
  Me.MinHeight = MinHeight: If MaxHeight > 0 Then Me.MaxHeight = MaxHeight
End Sub

Public Property Get ResizerStyle() As ResizerStyle
  ResizerStyle = mResizerStyle
End Property
Public Property Let ResizerStyle(ByVal NewValue As ResizerStyle)
  If mResizerStyle = NewValue Then Exit Property
  mResizerStyle = NewValue
End Property

Private Sub W_AddedToHierarchy()
  AlignMe
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As cWidgetBase)
  W.Refresh
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As cWidgetBase)
  W.Refresh
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  W.MoveToFront
  mDownParentLeft = W.Parent.Left: mDownParentTop = W.Parent.Top
  mDownParentWidth = W.Parent.Width: mDownParentHeight = W.Parent.Height
  mDownX = x: mDownY = y: mDown = True
End Sub
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim Delta As Single, PT As Single, PL As Single, PW As Single, PH As Single, MaxW As Single, MaxH As Single

  If Cairo.ImageList.Exists("VSplitCursor") Then
    W.MouseIconImageKey = IIf(mResizerStyle < 2, "VSplitCursor,16,16", "HSplitCursor,16,16")
  Else
    W.MousePointer = IIf(mResizerStyle < 2, IDC_SIZEWE, IDC_SIZENS)
  End If
  If Button = 0 Then Exit Sub
  
  PL = mDownParentLeft: PT = mDownParentTop
  PW = mDownParentWidth: PH = mDownParentHeight
  Select Case mResizerStyle
    Case ResizerVerticalRightAligned:    Delta = x - mDownX: PW = PW + Delta
    Case ResizerVerticalLeftAligned:     Delta = mDownX - x: PW = PW + Delta: PL = PL - Delta
    Case ResizerHorizontalBottomAligned: Delta = y - mDownY: PH = PH + Delta
    Case ResizerHorizontalTopAligned:    Delta = mDownY - y: PH = PH + Delta: PT = PT - Delta
  End Select
  
  MaxW = MaxWidth
  MaxH = MaxHeight
  If W.Parent Is W.Root Then
    Dim DlgFrm As cWidgetForm
    Set DlgFrm = Cairo.WidgetForms.ItemByHWnd(W.Root.DialogFormHwnd)
    If Not DlgFrm Is Nothing Then
      If DlgFrm.ScaleWidth < MaxW Then MaxW = DlgFrm.ScaleWidth
      If DlgFrm.ScaleHeight < MaxH Then MaxH = DlgFrm.ScaleHeight
    End If
  End If

  If PW < MinWidth Then PW = MinWidth:   PL = mDownParentLeft: Delta = 0
  If PW > MaxW - 1 Then PW = MaxW - 1:  PL = mDownParentLeft: Delta = 0
  If PH < MinHeight Then PH = MinHeight: PT = mDownParentTop:  Delta = 0
  If PH > MaxH - 1 Then PH = MaxH - 1: PT = mDownParentTop: Delta = 0
  
'  Debug.Print PT, PH, MaxH, MaxHeight
  RaiseEvent ResizerMoved(PL, PT, PW, PH, Delta)
  W.RaiseBubblingEvent Me, "ResizerMoved", PL, PT, PW, PH, Delta
  mDownParentLeft = W.Parent.Left: mDownParentTop = W.Parent.Top
  mDownParentWidth = W.Parent.Width: mDownParentHeight = W.Parent.Height
  AlignMe
End Sub

Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  AlignMe
  mDown = False
End Sub

Private Sub W_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
  If EventName = "W_ContainerResize" Then AlignMe
End Sub

Private Sub AlignMe()
Dim x As Single, y As Single, dx As Single, dy As Single
  dx = mSizePixels: dy = mSizePixels
  Select Case mResizerStyle
    Case ResizerVerticalRightAligned:    x = W.Parent.ScaleWidth - dx: dy = W.Parent.ScaleHeight
    Case ResizerVerticalLeftAligned:    dy = W.Parent.ScaleHeight
    Case ResizerHorizontalBottomAligned: y = W.Parent.ScaleHeight - dy: dx = W.Parent.ScaleWidth
    Case ResizerHorizontalTopAligned:   dx = W.Parent.ScaleWidth
  End Select
  W.Move x, y, dx, dy
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = IIf(W.MouseOver Or mDown, 0.6, 0.1)
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim Pat As cCairoPattern, PEndX As Long, PEndY As Long
  If dx = 0 Then dx = W.Width
  If dy = 0 Then dy = W.Height
  PEndX = dx
  PEndY = dy
  
  If mResizerStyle < 2 Then PEndY = 0 Else PEndX = 0
  Set Pat = Cairo.CreateLinearPattern(0, 0, PEndX, PEndY)
     Pat.AddColorStop 0, W.BackColor, Alpha, 1.4
     Pat.AddColorStop 1, W.BackColor, Alpha, 0.7
  CC.Rectangle 0, 0, dx, dy
  CC.Fill , Pat
End Sub
  
 



