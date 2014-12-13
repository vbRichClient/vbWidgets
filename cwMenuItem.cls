VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Event ShowSubMenu(Sender As cwMenuItem)

Private dx As Single, dy As Single, Alpha As Single
Private mCaption As String, mShortcutHint As String, mSubMenuDS As cMenuItem
Private mDeactivatedViaMouseLeave As Boolean

Private mIsCheckable As Boolean, mChecked As Boolean, mIsOption As Boolean
Private WithEvents tmrSubMenuHover As cTimer
Attribute tmrSubMenuHover.VB_VarHelpID = -1

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
 
Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Public Property Get SubMenuDS() As cMenuItem
  Set SubMenuDS = mSubMenuDS
End Property
Friend Property Set SubMenuDS(NewValue As cMenuItem)
  Set mSubMenuDS = NewValue
End Property

Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
Dim AccKey$, Pos&
  If mCaption = NewValue Then Exit Property
  mCaption = NewValue
  W.Refresh
  Pos = InStr(Replace(mCaption, "&&", "--"), "&")
  If Pos Then AccKey = Mid$(Replace(mCaption, "&&", "--"), Pos + 1, 1)
  Pos = InStr(mCaption, "|")
  If Pos Then mShortcutHint = Mid$(mCaption, Pos + 1)

  If Len(AccKey) Then W.AccessKeys = AccKey
End Property

Public Property Get Checked() As Boolean
  Checked = mChecked
End Property
Friend Property Let Checked(ByVal NewValue As Boolean)
  mChecked = NewValue
End Property

Public Property Get IsCheckable() As Boolean
  IsCheckable = mIsCheckable
End Property
Friend Property Let IsCheckable(ByVal NewValue As Boolean)
  mIsCheckable = NewValue
  If Not NewValue Then mIsOption = False    'IsOption works only if IsCheckable
End Property
 
Public Property Get IsOption() As Boolean
  IsOption = mIsOption
End Property
Friend Property Let IsOption(ByVal NewValue As Boolean)
  mIsOption = NewValue
End Property
 
Public Property Get IsActive() As Boolean
   IsActive = W.Focused And (Not mSubMenuDS Is Nothing Or Not mDeactivatedViaMouseLeave)
End Property

Private Sub W_LostFocus()
  W.MouseLeave Nothing
  W.Refresh
End Sub

Private Sub W_GotFocus()
   mDeactivatedViaMouseLeave = False
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As cWidgetBase)
  If Not mSubMenuDS Is Nothing Then Set tmrSubMenuHover = New_c.Timer(350, True)
  mDeactivatedViaMouseLeave = False
  W.SetFocus
  W.Refresh
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As cWidgetBase)
  Set tmrSubMenuHover = Nothing
  mDeactivatedViaMouseLeave = True
  W.Refresh
End Sub
Private Sub tmrSubMenuHover_Timer()
  Set tmrSubMenuHover = Nothing
  If Not W.Enabled Then Exit Sub
  RaiseEvent ShowSubMenu(Me)
  W.RaiseBubblingEvent Me, "ShowSubMenu", W.object
End Sub

Private Sub W_AccessKeyPress(KeyAscii As Integer)
  If InStr(1, W.AccessKeys, Chr$(KeyAscii), vbTextCompare) Then
    If W.Enabled Then
      W.SetFocus
      W.RaiseBubblingEvent Me, "W_MouseUp", 1, 0, 1, 1
    End If
  End If
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Const IcoOffsX& = 30, IcoSize& = 16, ArrowSize& = 9
  
  CC.SetSourceColor W.BackColor
  CC.Paint
  
  CC.DrawLine IcoOffsX - 1, 0, IcoOffsX - 1, dy, True, 1, &HD0D0D0, Alpha
  CC.DrawLine IcoOffsX, 0, IcoOffsX, dy, True, 1, vbWhite, Alpha
  
  If mCaption = "-" Then
    CC.DrawLine IcoOffsX, 4, dx - 2, 4, True, 1, &HD0D0D0, Alpha
    CC.DrawLine IcoOffsX, 5, dx - 2, 5, True, 1, vbWhite, Alpha
    
  Else
    If IsActive Then
      W.Alpha = 0.6
      CC.SetLineWidth 1, True
      Cairo.Theme.DrawTo CC, W, thmTypeListSelection, 0, 2, 0, dx - 4, dy, 3
    End If

    If mIsCheckable Then
      W.Alpha = 0.3
      Cairo.Theme.DrawTo CC, W, thmTypeListSelection, 0, 2, 0, dy + 2, dy, 3
      If mChecked Then
        If mIsOption Then
          DrawOptionMark CC, IcoSize + 3
        Else
          DrawCheckMark CC, IcoSize + 3
        End If
      End If
    End If
    
    W.Alpha = Alpha
    If Not mChecked And Len(W.ImageKey) Then
      CC.RenderSurfaceContent W.ImageKey, 6.5, (dy - 16) \ 2 + 0.5, IcoSize, IcoSize, , Alpha
    End If
    If Not mSubMenuDS Is Nothing Then 'this is a ParentNode
      Cairo.Theme.DrawTo CC, W, thmTypeArrow, 0, dx - 18, (dy - ArrowSize) \ 2 + 1, ArrowSize, ArrowSize, 0, thmDirectionRight
    End If
    W.SelectFontSettingsInto CC
    CC.DrawText IcoOffsX, 0, dx - IcoOffsX, dy, Replace(mCaption, "|" & mShortcutHint, vbNullString), True, vbLeftJustify, 5, True, dtHasAccelerators, Alpha
    CC.DrawText IcoOffsX, 0, dx - IcoOffsX, dy, mShortcutHint, True, vbRightJustify, 5, True, dtHasAccelerators, Alpha
  End If
End Sub
  
Private Sub DrawCheckMark(CC As cCairoContext, ByVal CheckSize As Long)
Dim x As Double, y As Double
    y = (W.Height - CheckSize) / 2 + 2
    x = y + 3.3
    CC.SetLineCap CAIRO_LINE_CAP_ROUND
    CC.SetLineWidth 2
      CC.SetSourceColor W.FocusColor, 0.85, 0.27
      x = x + 2: CheckSize = CheckSize - 4
      y = y + 2: CheckSize = CheckSize - 4
      CC.MoveTo x + CheckSize - 1.75, y + 1
      CC.LineTo x + CheckSize * 0.395, y + CheckSize - 1.9
      CC.LineTo x + 1.75, y + CheckSize * 0.55
    CC.Stroke True
    CC.SetLineWidth 1
      CC.SetSourceColor W.FocusColor, 0.35
    CC.Stroke
End Sub

Private Sub DrawOptionMark(CC As cCairoContext, ByVal OptionSize As Long)
Dim x As Double, y As Double
    y = (W.Height - OptionSize) / 2 + 2
    x = y + 3.3
    CC.SetLineCap CAIRO_LINE_CAP_ROUND
 
      OptionSize = OptionSize - 8
      x = x + 2 + OptionSize / 2
      y = y + 2 + OptionSize / 2
     
    Dim Pat As cCairoPattern
    CC.ARC x, y, OptionSize / 2
  
    Set Pat = Cairo.CreateRadialPattern(x, y, OptionSize / 2, x + OptionSize / 6, y - OptionSize / 4, 0)
      Pat.AddColorStop 1, W.FocusColor, 0#
      Pat.AddColorStop 0, W.FocusColor, 0.85, 0.22
    CC.Fill True, Pat
      
    CC.SetLineWidth 1
      CC.SetSourceColor W.FocusColor, 0.35, 0.4
    CC.Stroke
End Sub

