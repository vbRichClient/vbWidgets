VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwMenuBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Click(ByVal CurMenuItemPath As String)

Public WithEvents CurPopPup As cwMenu
Attribute CurPopPup.VB_VarHelpID = -1
Private mDataSource As cMenuItem
Private WithEvents tmrCleanup As cTimer
Attribute tmrCleanup.VB_VarHelpID = -1
Private WithEvents KWatch As cKeyWatcher, ShiftOrCtrlKeyDown As Boolean, AltDown As Boolean
Attribute KWatch.VB_VarHelpID = -1
 
Private dx As Single, dy As Single, Alpha As Single, mBlockDestroy As Boolean
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
  Set KWatch = New_c.KeyWatcher("16,17,18") 'Shift,Ctrl,Alt - Keys
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
Dim i As Long, xOffs As Single, TxtWidth As Double, FontHeight As Double, CC As cCairoContext
Dim currentMenuBarItem As cwMenuBarItem

  Set mDataSource = NewValue
  Widgets.RemoveAll
 
  Set CC = W.MeasureContext
  CC.SelectFont Cairo.Theme.GetIconFontName, W.FontSize, vbBlack
  
  xOffs = 3
  For i = 0 To mDataSource.SubItemCount - 1
    TxtWidth = CC.GetTextExtents(mDataSource.SubItemByIndex(i).Caption, FontHeight)
    If InStr(mDataSource.SubItemByIndex(i).Caption, "&") Then
      TxtWidth = TxtWidth - CC.GetTextExtents("&")
    End If
    Set currentMenuBarItem = Widgets.Add(New cwMenuBarItem, "MenuBarItem" & i, xOffs, 2, TxtWidth + 15, FontHeight * 1.6666 - 3)
    Set currentMenuBarItem.DataSource = mDataSource.SubItemByIndex(i)
    currentMenuBarItem.Widget.FontSize = W.FontSize
    currentMenuBarItem.Widget.Enabled = mDataSource.SubItemByIndex(i).Enabled
    xOffs = xOffs + 15 + TxtWidth
  Next i
End Property
 
Private Sub CurPopPup_Click(ByVal CurMenuItemPath As String)
  RaiseEvent Click(CurMenuItemPath)
  W.RaiseBubblingEvent Me, "Click", CurMenuItemPath
End Sub

Private Sub CurPopPup_MenuBarEntryShift(ByVal ShiftLeft As Boolean)
Dim MenuBarItem As cwMenuBarItem, Index As Long, i As Long
  If Widgets.Count = 0 Then Exit Sub
  For Each MenuBarItem In Widgets
    If CurPopPup Is Nothing Then
      If MenuBarItem.Widget.MouseOver Then
        MenuBarItem.Widget.MouseLeave Nothing
        Exit For
      End If
    Else
      If MenuBarItem.DataSource Is CurPopPup.DataSource Then Exit For
    End If
    Index = Index + 1
  Next MenuBarItem
  
  If Index = Widgets.Count Then Exit Sub 'not found
  
  For i = 1 To Widgets.Count 'let's try to shift the entry (trying Widgets.Count times)
    Index = Index + IIf(ShiftLeft, -1, 1)
    If Index < 0 Then Index = Widgets.Count - 1
    If Index >= Widgets.Count Then Index = 0
    
    Set MenuBarItem = Widgets(Index + 1)
    If MenuBarItem.Widget.Enabled Then Exit For 'in case it is disabled, we shift once more
  Next i
  
  mBlockDestroy = True
  MenuBarItem.ShowPopUp
  MenuBarItem.Widget.Refresh
  W.Refresh
  mBlockDestroy = False
End Sub

Private Sub CurPopPup_MenuDestroyed(ByVal DestroyedByKeyPress As Boolean)
Dim MenuBarItem As cwMenuBarItem
  If mBlockDestroy Or W.Root Is Nothing Then Exit Sub
  If W.Root.GetWindowUnderCursor = W.Root.hWnd And Not DestroyedByKeyPress Then
    For Each MenuBarItem In Widgets
      If MenuBarItem.Widget.MouseOver Then Exit Sub
    Next
  End If
  W.Root.Widget.MouseUp 0, 0, -1, -1
  
  Set tmrCleanup = New_c.Timer(50, True)
End Sub
 
Private Sub KWatch_VKeyDown(ByVal vKey As Integer, ByVal MapIdx As Long)
  If vKey = 16 Or vKey = 17 Then ShiftOrCtrlKeyDown = True
  If vKey = 18 And Not ShiftOrCtrlKeyDown Then AltDown = True
End Sub

Private Sub KWatch_VKeyUp(ByVal vKey As Integer, ByVal MapIdx As Long)
  If AltDown And Not ShiftOrCtrlKeyDown And CurPopPup Is Nothing And Widgets.Count > 0 Then
    Dim MenuBarItem As cwMenuBarItem
    For Each MenuBarItem In Widgets
      If MenuBarItem.Widget.MouseOver Then
        MenuBarItem.Widget.MouseLeave Nothing
        W.Root.Widget.MouseUp 0, 0, -1, -1
        W.Root.Refresh
        Exit Sub
      End If
    Next
    For Each MenuBarItem In Widgets
      If MenuBarItem.Widget.Enabled Then
        W.Root.Widget.SetFocus
        MenuBarItem.Widget.MouseEnter Nothing
        Exit For
      End If
    Next
  End If
  If vKey = 16 Or vKey = 17 Then ShiftOrCtrlKeyDown = False
  If vKey = 18 Then AltDown = False
End Sub

Private Sub tmrCleanup_Timer()
  Set tmrCleanup = Nothing
  Set CurPopPup = Nothing
  W.Refresh
End Sub
 
Private Sub W_ContainerResize()
  W.Move 0, 0, W.Parent.Width, W.Height
  W.Refresh
End Sub
 
Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not CurPopPup Is Nothing Then Exit Sub
  Select Case KeyCode
    Case vbKeyLeft:  CurPopPup_MenuBarEntryShift True
    Case vbKeyRight: CurPopPup_MenuBarEntryShift False
  End Select
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim Pat As cCairoPattern
  Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
      Pat.AddColorStop 0, W.BackColor, , 1.11
      Pat.AddColorStop 1, W.BackColor, , 0.93
  CC.Paint , Pat
  CC.DrawLine 0, dy - 1, dx, dy - 1, True, 1, vbBlack, 0.3
End Sub
