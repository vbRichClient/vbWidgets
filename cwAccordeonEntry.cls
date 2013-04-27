VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwAccordeonEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Private Alpha As Single, dx As Single, dy As Single
Private mCaption As String

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

  If Len(AccKey) Then W.AccessKeys = AccKey
End Property

Public Property Get IsOpen() As Boolean
  IsOpen = W.Parent.Object.CurrentEntryIndex = CLng(Split(W.Key, "_")(0))
End Property

Public Property Get AssociatedWidget() As Object
  With W.Parent.Widgets 'the associated Widget is always one IndexPosition "below us"
    Set AssociatedWidget = .Item(.GetOneBasedChildIndexByKey(W.Key) + 1)
  End With
End Property

Private Sub W_GotFocus()
  If IsOpen Then AssociatedWidget.Widget.SetFocus
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As cWidgetBase)
  W.Refresh
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As cWidgetBase)
  W.Refresh
End Sub

Private Sub W_AccessKeyPress(KeyAscii As Integer)
  If InStr(1, W.AccessKeys, Chr$(KeyAscii), vbTextCompare) Then
    W.SetFocus
    W.RaiseBubblingEvent Me, "W_Click"
  End If
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim Pat As cCairoPattern, Color As Long, CaptExt As Double
Dim IconKeys() As String, IconKey As String
  
  Color = W.ShadeColor(W.BackColor, 0.98) 'default color is the greyish one
  If IsOpen Then Color = W.SelectionColor
  If W.MouseOver Then Color = W.HoverColor
  
  
  CC.SetLineWidth 1, True
  Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
    Pat.AddColorStop 0, &HF0F0F0, Alpha
    Pat.AddColorStop 0.03, &HF0F0F0, Alpha
    Pat.AddColorStop 0.05, Color, Alpha, 0.98
    Pat.AddColorStop 0.1, Color, Alpha, 0.96
    Pat.AddColorStop 0.88, Color, Alpha, 1.05
    Pat.AddColorStop 1, Color, Alpha, 1.2
    CC.Rectangle 0, 0, dx, dy, True
  CC.Fill , Pat
    CC.SetSourceColor W.BorderColor, Alpha
    CC.DrawLine 0, dy - 1, dx, dy - 1, True
  CC.Stroke
  
  
  'the focused area and a dotted rectangle
  If W.Focused Then
    CaptExt = CC.GetTextExtents(mCaption)  'measure the current Pixel-Len of the caption-text

    CC.RoundedRect dy - 3, 3, CaptExt + 4, dy - 7, 3, True
      CC.SetSourceColor W.FocusColor, Alpha * 0.3, 1.1
    CC.Fill
    Cairo.Theme.DrawTo CC, W, thmTypeDottedRectangle, 0, dy - 3, 3, CaptExt + 4, dy - 7, 2
  End If
  
  'the Icon (if there was a resource-info given in the W.ImageKey)
  If Len(W.ImageKey) Then
    IconKeys = Split(W.ImageKey, ",")
    IconKey = Trim$(IconKeys(0))
    If IsOpen Then 'we are the currently active (open and expanded) entry
      If UBound(IconKeys) > 0 Then IconKey = Trim$(IconKeys(1))
    End If
    CC.RenderSurfaceContent IconKey, 3, 3, dy - 7, dy - 7, , Alpha
  End If
  
  'the Caption-Text
  W.SelectFontSettingsInto CC
  CC.DrawText dy - 7, 1, dx, dy, mCaption, True, vbLeftJustify, 6, True, dtHasAccelerators, Alpha
End Sub
  

