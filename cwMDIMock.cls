VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwMDIMock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mBorderWidth As Long = 5
Private Const mCaptionHeight As Long = 22
Private mCaption As String

Private dx As Single, dy As Single, Alpha As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
  W.ForeColor = &H222222
  W.FontSize = 10
  W.SetClientAreaOffsets mBorderWidth, mCaptionHeight + mBorderWidth, mBorderWidth, mBorderWidth
End Sub

'the following two Public Props are a "required convention", needed in *each* cwWidget-class
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
  If mCaption = NewValue Then Exit Property
  mCaption = NewValue
  W.Refresh
End Property

Private Sub W_EnterFocus()
  W.Refresh
End Sub
Private Sub W_ExitFocus()
  W.Refresh
End Sub

Private Sub W_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
Dim PRecurse As cWidgetBase
  If EventName <> "W_MouseDown" Then Exit Sub
  W.MoveToFront: W.Refresh
  Set PRecurse = Sender.Widget
  Do Until PRecurse Is W
    If PRecurse.CanGetFocus Then Exit Do
    Set PRecurse = PRecurse.Parent
  Loop
  If PRecurse Is W Then W.SetFocus
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  W.MoveToFront: W.Refresh
  W.Moveable = y < mBorderWidth + mCaptionHeight
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited

  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim NCColor As Long, Pat As cCairoPattern
  
  'nonclient-area drawing (first we determine the NonClient-Color)
  NCColor = IIf(W.Active, W.SelectionColor, W.ShadeColor(W.BackColor, 0.95))

  'the outline, part 1 (a thin rectangle, to achieve TopLeft-lighting)
  CC.SetLineWidth 2
    CC.RoundedRect 1, 1, dx - 2, dy - 2, mBorderWidth, True, cmTop
    CC.SetSourceColor &HEEEEEE, Alpha * 0.95
  CC.Stroke

  'the outline, part 2 (a thicker one - this is mainly for the NC-colored Borders)
  CC.SetLineWidth 3
    CC.RoundedRect 2, 2, dx - 3.3, dy - 3.5, mBorderWidth - 1, True, cmTop
    CC.SetSourceColor NCColor, Alpha * 0.95, 0.95
  CC.Stroke
   
  'now the Caption-Gradient
  CC.SetLineWidth 0
  CC.RoundedRect mBorderWidth - 1, mBorderWidth - 2, dx - 2 * mBorderWidth + 2, mCaptionHeight + 1, mBorderWidth - 2, True, cmTop
  Set Pat = Cairo.CreateLinearPattern(0, mBorderWidth - 2, 0, mCaptionHeight + mBorderWidth)
    Pat.AddColorStop 0, NCColor, Alpha * 0.95, 0.88
    Pat.AddColorStop 1, NCColor, Alpha * 0.95, 1.03
  CC.Fill , Pat

 'the Min/Max/Close-ButtonGroup
  CC.SetLineWidth 1
  CC.RoundedRect dx - 76, 0, 70, Int(mCaptionHeight * 0.8), mBorderWidth - 0.4, True, cmBottom
  Set Pat = Cairo.CreateLinearPattern(0, 0, 0, Int(mCaptionHeight * 0.8) - 2)
    Pat.AddColorStop 0, W.BackColor, Alpha * 0.55, 1.2
    Pat.AddColorStop 0.6, W.BackColor, Alpha * 0.55, 1.05
    Pat.AddColorStop 0.61, W.BackColor, Alpha * 0.55, 0.85
    Pat.AddColorStop 1, W.BackColor, Alpha * 0.55, 0.8
    CC.Fill True, Pat
  CC.SetLineWidth 3
    CC.SetSourceColor vbWhite, Alpha * 0.3
  CC.Stroke True
  CC.SetLineWidth 1
    CC.SetSourceColor W.BorderColor, Alpha * 0.7
  CC.Stroke
  
  'outline, part 3 (the outermost line, a thin one)
  NCColor = IIf(W.Active, W.FocusColor, W.BorderColor)
  CC.SetLineWidth 1
    CC.RoundedRect 0, 0, dx, dy, mBorderWidth + 1, True, cmTop
    CC.SetSourceColor NCColor, Alpha * 0.9
  CC.Stroke
  
  'the "Form"-Icon...
  CC.RenderSurfaceContent "MDIIcon", 7, 4, 18, 18, , Alpha * 2
  
  '... and the "Caption"-Text
  If W.Active Then 'draw some light-grey Text-Imprint first
    W.SelectFontSettingsInto CC, &HE0E0D0
    CC.TextOut 27, 5, mCaption, , Alpha * 2
  End If
  W.SelectFontSettingsInto CC
  CC.TextOut 27, 6, mCaption, , Alpha * 2

 'Now the clientarea-coloring
  CC.SetLineWidth 1, True
    CC.Rectangle mBorderWidth - 1, mCaptionHeight + mBorderWidth - 1, dx - 2 * mBorderWidth + 2, dy - mCaptionHeight - 2 * mBorderWidth + 2, True
    CC.SetSourceColor W.BackColor, Alpha
  CC.Fill True
    CC.SetSourceColor W.BorderColor, Alpha * 0.8
  CC.Stroke
End Sub
