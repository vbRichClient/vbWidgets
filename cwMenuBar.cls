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

'below API-call only used in case of a normal VB-HostForm (usage is made in the KWatch_VKeyDown-Event only)
Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long

Event Click(ByVal CurMenuItemPath As String)

Public WithEvents CurPopUp As cwMenu
Attribute CurPopUp.VB_VarHelpID = -1
Private mDataSource As cMenuItem
Private mKeyboardShortcuts As cCollection, mKeyWatcherCodes As String
Public LocalizedCtrlMarker As String, LocalizedShiftMarker As String

Private WithEvents tmrCleanup As cTimer, WithEvents tmrWatchActive As cTimer
Attribute tmrCleanup.VB_VarHelpID = -1
Attribute tmrWatchActive.VB_VarHelpID = -1
Private WithEvents KWatch As cKeyWatcher, TabOrShiftOrCtrlKeyDown As Boolean, MenuBarActive As Boolean
Attribute KWatch.VB_VarHelpID = -1

Private mCtrlDown As Long, mShiftDown As Long
Private dx As Single, dy As Single, Alpha As Single, mBlockDestroy As Boolean
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
 
  Set mKeyboardShortcuts = New_c.Collection
  'we initialize these Markers to the english locale by default
  LocalizedCtrlMarker = "Ctrl"
  LocalizedShiftMarker = "Shift"
  Set tmrWatchActive = New_c.Timer(60, True)
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
  Set KWatch = Nothing
  Set CC = W.MeasureContext
  CC.SelectFont Cairo.Theme.GetIconFontName, W.FontSize, vbBlack
  
  xOffs = 3
  mKeyWatcherCodes = "9,16,17,18" 'Tab,Shift,Ctrl,Alt - Keys as the default
  mKeyboardShortcuts.RemoveAll
  
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
    RegisterKeyboardShortcuts mDataSource.SubItemByIndex(i), mDataSource.SubItemByIndex(i).Key
  Next i

  W.Refresh
End Property
 
Friend Sub RegisterKeyboardShortcuts(pItem As cMenuItem, pPath As String)
Dim i As Long, s() As String, ModifierValue As Long, sKey As String, KeyCode As Long
'nothing bullet-proof or robust in here since it's undecided how to approach the storage of the shortcuts
'for now, just something that sorta works!
   If InStr(pItem.Caption, "|") Then
      s = Split(pItem.Caption, "|")
      ModifierValue = IIf(InStr(1, s(1), LocalizedCtrlMarker & "+", vbTextCompare), 1, 0) + _
                      IIf(InStr(1, s(1), LocalizedShiftMarker & "+", vbTextCompare), 2, 0)
      s = Split(s(1), "+")
      sKey = UCase$(s(UBound(s)))
      If Len(sKey) = 1 Then 'a single char (letter or number)
        KeyCode = Asc(sKey)
      ElseIf Left$(sKey, 1) = "F" Then 'F1-F1? keys
        KeyCode = Replace(sKey, "F", vbNullString) + 111
      Else
         'not valid/handled as yet
      End If
      If Not mKeyboardShortcuts.Exists(ModifierValue & "+" & KeyCode) Then
        mKeyboardShortcuts.Add pPath, ModifierValue & "+" & KeyCode
        mKeyWatcherCodes = mKeyWatcherCodes & "," & KeyCode
      End If
   End If
   For i = 0 To pItem.SubItemCount - 1
     RegisterKeyboardShortcuts pItem.SubItemByIndex(i), pPath & ">" & pItem.SubItemByIndex(i).Key
   Next i
End Sub

Public Sub CleanupActiveState(Optional ExcludeW As cWidgetBase)
Dim MenuBarItem As cwMenuBarItem
  For Each MenuBarItem In Widgets
    If Not ExcludeW Is MenuBarItem.Widget Then MenuBarItem.Widget.MouseLeave W
    MenuBarItem.Widget.Refresh
  Next
  W.Refresh
End Sub

Private Sub CurPopUp_Click(ByVal CurMenuItemPath As String)
  RaiseEvent Click(CurMenuItemPath)
  W.RaiseBubblingEvent Me, "Click", CurMenuItemPath
  W.Refresh
End Sub

Private Sub CurPopUp_MenuBarEntryShift(ByVal ShiftLeft As Boolean)
Dim MenuBarItem As cwMenuBarItem, Index As Long, i As Long
  If Widgets.Count = 0 Then Exit Sub
  For Each MenuBarItem In Widgets
    If CurPopUp Is Nothing Then
      If MenuBarItem.Widget.MouseOver Then
        MenuBarItem.Widget.MouseLeave Nothing
        Exit For
      End If
    Else
      If MenuBarItem.DataSource Is CurPopUp.DataSource Then Exit For
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
    If MenuBarActive Then
      MenuBarItem.Widget.MouseEnter Nothing
    Else
      MenuBarItem.ShowPopUp True
    End If
    MenuBarItem.Widget.Refresh
    W.Refresh
  mBlockDestroy = False
End Sub

Private Sub CurPopUp_MenuDestroyed(ByVal DestroyedByKeyPress As Boolean)
Dim MenuBarItem As cwMenuBarItem
  If mBlockDestroy Or W.Root Is Nothing Then Exit Sub
  If W.Root.GetWindowUnderCursor = W.Root.hWnd And Not DestroyedByKeyPress Then
    For Each MenuBarItem In Widgets
      If MenuBarItem.Widget.MouseOver Then Exit Sub
    Next
  End If
  W.Root.Widget.MouseUp 0, 0, -1, -1

  Set tmrCleanup = New_c.Timer(60, True)
End Sub
 
Private Sub KWatch_VKeyDown(ByVal vKey As Integer, ByVal MapIdx As Long)
  If vKey = 9 Or vKey = 16 Or vKey = 17 Then TabOrShiftOrCtrlKeyDown = True
  If vKey = 17 Then mCtrlDown = 1
  If vKey = 16 Then mShiftDown = 2
 
  CheckForKeyBoardShortcut vKey
  
  If vKey = 18 Then
    W.Root.WidgetForm.SetFocus
    W.SetFocus
    If Not CurPopUp Is Nothing Then CurPopUp.DestroyPopup: Set CurPopUp = Nothing: MenuBarActive = True
 
    If Not Screen.ActiveForm Is Nothing Then 'we send an Esc-KeyUp-Event in case of a VB-Form as TopLevel-Host
      Dim Evt(0 To 27) As Byte: Evt(0) = 1: Evt(4) = vbKeyEscape: Evt(8) = 2 '<- flagged as KeyUp
      SendInput 1, Evt(0), UBound(Evt) + 1
    End If
  End If
End Sub

Friend Sub CheckForKeyBoardShortcut(vKey As Integer)
Dim pMenuItemPath As String
   If Not CurPopUp Is Nothing Or mDataSource Is Nothing Then Exit Sub
   If mKeyboardShortcuts.Exists(mCtrlDown + mShiftDown & "+" & vKey) Then
      pMenuItemPath = mKeyboardShortcuts(mCtrlDown + mShiftDown & "+" & vKey)
      If mDataSource.SubItemByPathKey(mDataSource.Key & ">" & pMenuItemPath).Enabled Then
         CurPopUp_Click pMenuItemPath
      End If
   End If
End Sub

Private Sub KWatch_VKeyUp(ByVal vKey As Integer, ByVal MapIdx As Long)
  If vKey = 17 Then mCtrlDown = 0
  If vKey = 16 Then mShiftDown = 0
  If vKey = 18 Then
    CleanupActiveState
    If Not TabOrShiftOrCtrlKeyDown And Widgets.Count > 0 Then MenuBarActive = Not MenuBarActive
  End If
  
  If MenuBarActive And Not TabOrShiftOrCtrlKeyDown And CurPopUp Is Nothing And Widgets.Count > 0 Then
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
        W.Root.Widget.MouseUp 0, 0, -1, -1
        W.Root.Widget.SetFocus
        MenuBarItem.Widget.MouseEnter Nothing
        W.SetFocus
        Exit For
      End If
    Next
  End If
  If vKey = 9 Or vKey = 16 Or vKey = 17 Then TabOrShiftOrCtrlKeyDown = False
End Sub
 
Private Sub tmrCleanup_Timer()
  Set tmrCleanup = Nothing
  Set CurPopUp = Nothing
  W.Refresh
End Sub

Private Sub W_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
  If TypeOf Sender Is cwMenuBarItem And EventName = "AccessKeyPress" Then
    MenuBarActive = True
    W_KeyDown CInt(P1), 0
    Set tmrCleanup = Nothing
  End If
End Sub

Private Sub W_ContainerResize()
  W.Move 0, 0, W.Parent.Width, W.Height
  W.Refresh
End Sub

Private Sub W_ExitFocus()
  CleanupActiveState
End Sub

Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
Dim MenuBarItem As cwMenuBarItem
  If CurPopUp Is Nothing And MenuBarActive And Widgets.Count > 0 Then 'only go in here, when no Menu is currently shown
    For Each MenuBarItem In Widgets
      If MenuBarItem.Widget.MouseOver And KeyCode = vbKeyDown Then
        Exit For
      ElseIf InStr(1, MenuBarItem.Widget.AccessKeys, Chr$(KeyCode), vbTextCompare) Then
        CleanupActiveState
        KeyCode = vbKeyDown 'pretend the down arrow was pressed
        Exit For
      End If
    Next
    Select Case KeyCode
      Case vbKeyEscape
         CleanupActiveState
         MenuBarActive = False
'        If Not MenuBarItem Is Nothing Then MenuBarItem.Widget.MouseLeave Nothing
'        W.Root.Widget.MouseUp 0, 0, -1, -1
'        W.Root.Refresh
      Case vbKeyDown:
        If Not MenuBarItem Is Nothing Then
          MenuBarItem.Widget.MouseLeave Nothing
          MenuBarActive = False: MenuBarItem.ShowPopUp True
          MenuBarItem.Widget.Refresh
        End If
      Case vbKeyLeft:  CurPopUp_MenuBarEntryShift True
      Case vbKeyRight: CurPopUp_MenuBarEntryShift False
    End Select
  End If
End Sub
 
Private Sub tmrWatchActive_Timer()
  If Not CurPopUp Is Nothing Then MenuBarActive = False
  If W.Root Is Nothing Then Exit Sub
  ShortcutActivation W.Root.IsTopWindowActive
End Sub

Private Sub ShortcutActivation(ByVal IsTopWindowActive As Boolean)
  If IsTopWindowActive Then
    If KWatch Is Nothing Then Set KWatch = New_c.KeyWatcher(mKeyWatcherCodes)
  Else 'deactivate KeyWatching, in case the TopWindow is not the active one
    If Not KWatch Is Nothing Then Set KWatch = Nothing: CleanupActiveState: MenuBarActive = False
  End If
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

Private Sub Class_Terminate()
  Set tmrCleanup = Nothing
  Set tmrWatchActive = Nothing
  Set KWatch = Nothing
  Set mDataSource = Nothing
  If Not CurPopUp Is Nothing Then CurPopUp.DestroyPopup
  Set CurPopUp = Nothing
End Sub

