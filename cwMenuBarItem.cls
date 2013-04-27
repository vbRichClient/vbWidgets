VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwMenuBarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Event Click(ByVal CurMenuItemPath As String)

Private mDataSource As cMenuItem
 
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

Public Property Get DataSource() As cMenuItem
  Set DataSource = mDataSource
End Property
Public Property Set DataSource(NewValue As cMenuItem)
  Set mDataSource = NewValue
End Property

Public Sub ShowPopUp()
Dim MenuBar As cwMenuBar
  Set MenuBar = W.Parent.object
  If Not MenuBar.CurPopPup Is Nothing Then
    If Not MenuBar.CurPopPup.DataSource Is mDataSource Then MenuBar.CurPopPup.DestroyPopup
  End If
  Set MenuBar.CurPopPup = New cwMenu
  MenuBar.CurPopPup.Widget.FontSize = W.FontSize
      MenuBar.CurPopPup.InitAndShow W, mDataSource, PopupBehaviourDropDownLeftAligned
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 1 Then ShowPopUp
End Sub
Private Sub W_MouseEnter(ByVal MouseLeaveWidget As vbRichClient5.cWidgetBase)
  If IsPopupVisible And Not IsOwnPopupVisible Then
    ShowPopUp
  End If
  W.Refresh
  W.Parent.Refresh
End Sub
Private Function IsOwnPopupVisible() As Boolean
Dim MenuBar As cwMenuBar
  Set MenuBar = W.Parent.object
  If IsPopupVisible Then
    If MenuBar.CurPopPup.DataSource Is Nothing Then Exit Function
    IsOwnPopupVisible = MenuBar.CurPopPup.DataSource Is mDataSource
  End If
End Function
Private Function IsPopupVisible() As Boolean
  IsPopupVisible = Not W.Parent.object.CurPopPup Is Nothing
End Function

Private Sub W_MouseLeave(ByVal MouseEnterWidget As vbRichClient5.cWidgetBase)
  W.Refresh
  W.Parent.Refresh
End Sub

Private Sub W_AccessKeyPress(KeyAscii As Integer)
  If W.Root.IsHidden Then Exit Sub
 
  If InStr(1, W.AccessKeys, Chr$(KeyAscii), vbTextCompare) Then
    W.SetFocus
    ShowPopUp
    W.Refresh
    W.Parent.Refresh
  End If
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Public Sub SetAccelerator(ByVal Caption As String)
Dim AccKey$, Pos&
 
  Pos = InStr(Replace(Caption, "&&", "--"), "&")
  If Pos Then AccKey = Mid$(Replace(Caption, "&&", "--"), Pos + 1, 1)

  If Len(AccKey) Then W.AccessKeys = AccKey
End Sub

Private Sub Draw(CC As cCairoContext)
Dim TxtWidth As Double, FontHeight As Double, Pat As cCairoPattern

  If mDataSource Is Nothing Then Exit Sub
  SetAccelerator mDataSource.Caption
  
  CC.SetLineWidth 1
  CC.SelectFont Cairo.Theme.GetIconFontName, W.FontSize, vbBlack
  
  TxtWidth = CC.GetTextExtents(mDataSource.Caption, FontHeight)

  CC.SelectFont Cairo.Theme.GetIconFontName, W.FontSize, vbWhite
  CC.DrawText 0, 1, dx, dy, mDataSource.Caption, True, vbCenter, 2, 1, dtHasAccelerators

  CC.SelectFont Cairo.Theme.GetIconFontName, W.FontSize, vbBlack
  CC.DrawText 0, 0, dx, dy, mDataSource.Caption, True, vbCenter, 2, 1, dtHasAccelerators

  If W.MouseOver Or IsOwnPopupVisible Then
    CC.RoundedRect 0, 0, dx, dy - 1, 2.2, True
    CC.SetSourceColor vbBlack
    Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
        Pat.AddColorStop 0, vbWhite, , 0.85
        Pat.AddColorStop 1, vbWhite, , 0.45
    CC.Stroke , Pat
    CC.DrawLine 2, dy, dx - 3, dy, False, 1, vbBlack, 0.2
  End If
End Sub

