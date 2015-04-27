VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum enmMenuPopUpPosition
  PopupBehaviourMousePosExpandLeft = -2
  PopupBehaviourMousePosExpandRight = -1
  PopupBehaviourDropDownRightAligned = 0
  PopupBehaviourDropDownLeftAligned = 1
End Enum

Event ReplaceCaption(ByVal CurItem As cMenuItem, Caption As String)
Event Click(ByVal CurMenuItemPath As String)
Event MenuBarEntryShift(ByVal ShiftLeft As Boolean)
Event MenuDestroyed(ByVal DestroyedByKeyPress As Boolean)

'To have a option-like MenuItem, you must set the IsCheckable property to True and set the IconKey property
'to this value:
Private Const OptionMenuItemIconKey = "OptMenuItem"
'No real IconResource is needed at all in the global Cairo.ImageList - the IconImageKey is just used as
'the "magic Indicator" for the OptionBox-Rendering-Workaround.
'This is done because the cMenuItem class have not a IsOption property in RC5's vbRichClient.dll

Public WithEvents fPopUp As cfPopUp 'this PopupForm-HelperClass is hosted here in the vbWidgets.dll-Project
Attribute fPopUp.VB_VarHelpID = -1
Private WithEvents mInitiatorWidget As cWidgetBase
Attribute mInitiatorWidget.VB_VarHelpID = -1
Private WithEvents SC As cSubClass
Attribute SC.VB_VarHelpID = -1

Private mSubMenu As cwMenu, mActiveRootWidget As Object, mActiveMenuItem As cwMenuItem
Private mMouseX As Single, mMouseY As Single
Private mDataSource As cMenuItem, mRowHeight As Double, mMenuWidth As Double, mMenuHeight As Double
Private mDestroyedByKeyPress As Boolean

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
      W.BackColor = W.ShadeColor(W.BackColor, 1.02)
      W.SetClientAreaOffsets 1, 1, 1, 1
      W.CanGetFocus = False
      W.FontName = Cairo.Theme.GetIconFontName
      W.FontSize = 9
      W.Default = True
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

Public Property Get InitiatorWidget() As cWidgetBase
  If Not mInitiatorWidget Is Nothing Then Set InitiatorWidget = mInitiatorWidget
End Property

Public Property Get MenuKeyPath() As String
Dim M As cwMenu
  If mInitiatorWidget Is Nothing Then Exit Property
  If Not mInitiatorWidget.object Is Nothing Then
    If TypeOf mInitiatorWidget.object Is cwMenu Then
      Set M = mInitiatorWidget.object
      MenuKeyPath = M.MenuKeyPath & ">" & mDataSource.Key
    Else
      MenuKeyPath = mDataSource.Key
    End If
  End If
End Property

Public Property Get MenuItemKeyPath() As String
  If Not mActiveMenuItem Is Nothing And Not mActiveMenuItem.IsActive Then Exit Property
  If mDataSource Is Nothing Then Exit Property
  MenuItemKeyPath = MenuKeyPath
  If Not mActiveMenuItem Is Nothing Then MenuItemKeyPath = MenuItemKeyPath & ">" & mActiveMenuItem.Widget.Key
End Property

Public Sub InitAndShow(InitiatorWidget As cWidgetBase, Optional DataSource As cMenuItem, Optional ByVal DirectionBehaviour As enmMenuPopUpPosition, Optional ByVal WithFocus As Boolean)
Dim MenuItem As cwMenuItem, PopupPosX As Single, PopupPosY As Single, MousePosX As Long, MousePosY As Long
  
  If Not fPopUp Is Nothing Then DestroyPopup
  
  Set mInitiatorWidget = InitiatorWidget
  If mInitiatorWidget Is Nothing Then Exit Sub
 
  If TypeOf mInitiatorWidget.object Is cwMenuItem Then
    Set MenuItem = mInitiatorWidget.object
    Set mInitiatorWidget = mInitiatorWidget.Parent
  End If
  If mInitiatorWidget Is Nothing Then Exit Sub
  If mInitiatorWidget.Root Is Nothing Then Exit Sub
  If mInitiatorWidget.Root.MouseKeyDown Then mInitiatorWidget.MouseUp 1, 0, 0, 0
  
  If Not DataSource Is Nothing Then Set mDataSource = DataSource
  
  On Error Resume Next

  ConstructMenuFromDataSource mDataSource
  mMenuWidth = mMenuWidth * mInitiatorWidget.Zoom
  mMenuHeight = mMenuHeight * mInitiatorWidget.Zoom
  
  MousePosX = -1: MousePosY = -1
 
  If TypeOf mInitiatorWidget.object Is cwMenu Then 'it was an internal call
    MenuItem.Widget.CalculateScreenPopupOffsets mMenuWidth, mMenuHeight, PopupPosX, PopupPosY, True, MousePosX, MousePosY
  Else 'the initial usercall
    Set mActiveRootWidget = mInitiatorWidget.Root.ActiveWidget
    If Not mActiveRootWidget Is Nothing Then
      If TypeOf mActiveRootWidget.object Is cwMenuBar Then Set mActiveRootWidget = Nothing
    End If
    Select Case DirectionBehaviour
      Case PopupBehaviourDropDownLeftAligned
        Dim x As Single, y As Single
        mInitiatorWidget.CalculateScreenOffsets x, y
        MousePosX = x + mMenuWidth
        MousePosY = y + mInitiatorWidget.Height * mInitiatorWidget.Zoom
        mInitiatorWidget.CalculateScreenPopupOffsets mMenuWidth, mMenuHeight, PopupPosX, PopupPosY, False, MousePosX, MousePosY
        If MousePosY <> PopupPosY Then PopupPosY = PopupPosY - mInitiatorWidget.Height * mInitiatorWidget.Zoom
      Case PopupBehaviourDropDownRightAligned
        mInitiatorWidget.CalculateScreenPopupOffsets mMenuWidth, mMenuHeight, PopupPosX, PopupPosY, False, MousePosX, MousePosY
      Case PopupBehaviourMousePosExpandRight
        Cairo.GetMouseCursorPos MousePosX, MousePosY
        MousePosY = MousePosY + 1
        mInitiatorWidget.CalculateScreenPopupOffsets mMenuWidth, mMenuHeight, PopupPosX, PopupPosY, True, MousePosX, MousePosY
      Case PopupBehaviourMousePosExpandLeft
        Cairo.GetMouseCursorPos MousePosX, MousePosY
        mInitiatorWidget.CalculateScreenPopupOffsets mMenuWidth, mMenuHeight, PopupPosX, PopupPosY, False, MousePosX, MousePosY
    End Select
  End If
  
  Set fPopUp = New cfPopUp
  fPopUp.Load Me, PopupPosX, PopupPosY, mMenuWidth, mMenuHeight, mInitiatorWidget.Zoom
  
  fPopUp.Show
  
  If Not TypeOf mInitiatorWidget.object Is cwMenu Then
    mInitiatorWidget.Root.BlockKeyEvents True, fPopUp.Form
    Set SC = New_c.SubClass
    If App.LogMode Then SC.Hook mInitiatorWidget.Root.DialogFormHwnd
  End If
  If WithFocus And Widgets.Count > 0 Then Widgets(1).Widget.SetFocus
If Err Then Err.Clear
End Sub

Private Sub ShowSubMenu(ByVal MenuItem As cwMenuItem, Optional ByVal WithFocus As Boolean)
  If Not mSubMenu Is Nothing Then
    If mSubMenu.MenuKeyPath = MenuItemKeyPath Then Exit Sub
    mSubMenu.DestroyPopup
  End If
  Set mSubMenu = New cwMenu
      mSubMenu.Widget.FontSize = MenuItem.Widget.FontSize
      mSubMenu.InitAndShow MenuItem.Widget, MenuItem.SubMenuDS, , WithFocus
End Sub

Private Sub ConstructMenuFromDataSource(DataSource As cMenuItem)
Const IcoOffsX& = 30
Dim DSItem As cMenuItem, MenuItem As cwMenuItem, MCC As cCairoContext, RM As cwMenu
Dim i As Long, dx As Double, dy As Double, Caption As String
 
  Set RM = RootMenu
  
  Set mDataSource = DataSource
  Set MCC = W.MeasureContext
  mRowHeight = MCC.GetFontHeight + 7
  
  mMenuWidth = 0
  For i = 0 To mDataSource.SubItemCount - 1
    Set DSItem = mDataSource.SubItemByIndex(i)
    Caption = DSItem.Caption
    dx = MCC.GetTextExtents(Caption)
    If mMenuWidth < dx Then mMenuWidth = dx
  Next i
  mMenuWidth = IcoOffsX + mMenuWidth + 45
  
  mMenuHeight = 2
  For i = 0 To mDataSource.SubItemCount - 1
    Set DSItem = mDataSource.SubItemByIndex(i)
 
    Caption = DSItem.Caption
    If Caption = "-" Then
      dy = 8
    Else
      If Not RM Is Nothing Then RM.RaiseReplaceCaption DSItem, Caption
      dy = mRowHeight
    End If
    
    Set MenuItem = Widgets.Add(New cwMenuItem, DSItem.Key, 0, mMenuHeight, mMenuWidth - 2, dy)
        MenuItem.Widget.ImageKey = DSItem.IconKey
        MenuItem.Widget.Enabled = DSItem.Enabled
        MenuItem.Widget.FontName = W.FontName
        MenuItem.Widget.FontSize = W.FontSize
        MenuItem.Widget.BackColor = W.BackColor
        MenuItem.Caption = Caption
        MenuItem.IsCheckable = DSItem.IsCheckable
        MenuItem.IsOption = DSItem.IconKey = OptionMenuItemIconKey  'Here is the workaround to have
                                                                    'option-like cwMenuItems from cMenuItems
                                                                    'with "magic" iconkey indicator
        If DSItem.IsCheckable Then MenuItem.Checked = DSItem.Checked
        If Caption = "-" Then MenuItem.Widget.Enabled = False
        If DSItem.SubItemCount Then Set MenuItem.SubMenuDS = DSItem
    mMenuHeight = mMenuHeight + dy
  Next i
  mMenuHeight = mMenuHeight + 4
End Sub

Public Sub RaiseReplaceCaption(CurItem As cMenuItem, Caption As String)
  RaiseEvent ReplaceCaption(CurItem, Caption)
  W.RaiseBubblingEvent Me, "ReplaceCaption", CurItem, Caption
End Sub
Public Sub RaiseClick(CurMenuItemPath As String)
Dim MItem As cMenuItem
  On Error Resume Next
    Set MItem = mDataSource.SubItemByPathKey(CurMenuItemPath)
    If MItem.IsCheckable Then MItem.Checked = Not MItem.Checked
    DestroyPopup
  If Err Then Err.Clear
  RaiseEvent Click(CurMenuItemPath)
  W.RaiseBubblingEvent Me, "Click", CurMenuItemPath
End Sub

Private Sub SC_WindowProc(Result As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Select Case Msg 'check only for NonClient - WM_NCLBUTTONDOWN,WM_NCRBUTTONDOWN,WM_NCMBUTTONDOWN
    Case &HA1, &HA4, &HA7
        DestroyPopup
  End Select
  If Not SC Is Nothing Then Result = SC.CallWindowProc(Msg, wParam, lParam)
End Sub

Public Sub DestroyPopup()
Dim M As cwMenu, WasRootMenu As Boolean
  On Error Resume Next

    If RootMenu Is Me Then WasRootMenu = True
  
    If Not mInitiatorWidget Is Nothing Then
      If Not mInitiatorWidget.object Is Nothing Then
        If TypeOf mInitiatorWidget.object Is cwMenu Then
          Set M = mInitiatorWidget.object
          Set M.SubMenu = Nothing
        Else
          Widgets.RemoveAll
        End If
      End If
    End If
    If Not mSubMenu Is Nothing Then
      mSubMenu.DestroyPopup
      Set mSubMenu = Nothing
    End If

    Set mActiveMenuItem = Nothing
    Set mActiveRootWidget = Nothing
    Set mInitiatorWidget = Nothing
    
    If WasRootMenu Then
      RaiseEvent MenuDestroyed(mDestroyedByKeyPress)
      W.RaiseBubblingEvent Me, "MenuDestroyed", mDestroyedByKeyPress
    End If
    mDestroyedByKeyPress = False
    
    If Not fPopUp Is Nothing Then fPopUp.Unload
    Set fPopUp = Nothing
  If Err Then Err.Clear
End Sub
 
Private Property Get InMouseOverStack() As Boolean
Dim M As cwMenu, MouseM As cwMenu
  Set MouseM = MenuUnderCursor
  If MouseM Is Nothing Then Exit Property
  
  Set M = RootMenu
  Do Until M Is MouseM Or M Is Nothing
    If Me Is M Then InMouseOverStack = True
    Set M = M.SubMenu
  Loop
End Property

Public Property Get SubMenu() As cwMenu
  Set SubMenu = mSubMenu
End Property
Public Property Set SubMenu(NewValue As cwMenu)
  Set mSubMenu = NewValue
End Property

Public Property Get RootMenu() As cwMenu
Dim M As cwMenu
  If mInitiatorWidget Is Nothing Then Exit Property
  If Not mInitiatorWidget.object Is Nothing Then
    If TypeOf mInitiatorWidget.object Is cwMenu Then
      Set M = mInitiatorWidget.object
      Set RootMenu = M.RootMenu
    Else
      Set RootMenu = Me
    End If
  End If
End Property

Public Property Get MenuUnderCursor() As cwMenu
  If RootMenu Is Nothing Then Exit Property
  Set MenuUnderCursor = GetMenuUnderCursor(RootMenu)
End Property
Friend Function GetMenuUnderCursor(ByVal Menu As cwMenu) As cwMenu
Dim WR As cWidgetRoot
  Set WR = Menu.Widget.Root
  If WR Is Nothing Then Exit Function
  If WR.hWnd = WR.GetWindowUnderCursor Then
    Set GetMenuUnderCursor = Menu
  ElseIf Not Menu.SubMenu Is Nothing Then
    Set GetMenuUnderCursor = GetMenuUnderCursor(Menu.SubMenu)
  End If
End Function

'------ Event-checks, which ensure fPopUp-Deactivation under certain conditions -----
Private Sub fPopUp_AppDeactivate()
  DestroyPopup
End Sub
Private Sub fPopUp_MouseClickOutside()
  If Not InMouseOverStack Then DestroyPopup
End Sub
Private Sub fPopUp_CheckForAdditionalCloseConditions()
Dim Par As cWidgetForm
  
  If mInitiatorWidget Is Nothing Then Exit Sub
  If mInitiatorWidget.object Is Nothing Then Exit Sub
  If TypeOf mInitiatorWidget.object Is cwMenu Then Exit Sub
  If mInitiatorWidget.Root Is Nothing Then Exit Sub
  
  On Error Resume Next
  
    If Not mActiveRootWidget Is Nothing And Not mActiveRootWidget Is mInitiatorWidget.Root.ActiveWidget Then
      DestroyPopup
      Exit Sub
    End If
    
  If Err Then Err.Clear: Exit Sub
  
    If mInitiatorWidget.Root.WidgetForm Is Nothing Then Exit Sub
    Set Par = mInitiatorWidget.Root.WidgetForm.Parent
    If Not Par Is Nothing Then
      If Par.WindowState = vbMinimized Then DestroyPopup
    Else
      If mInitiatorWidget.Root.WidgetForm.WindowState = vbMinimized Then DestroyPopup
    End If
    
  If Err Then Err.Clear
End Sub
Private Sub mInitiatorWidget_ContainerResize()
  If mInitiatorWidget Is Nothing Then Exit Sub
  If mInitiatorWidget.object Is Nothing Then Exit Sub
  If Not TypeOf mInitiatorWidget.object Is cwMenu Then DestroyPopup
End Sub

'---------------- End of fPopUp-Deactivation-Checks ------------------

Private Sub fPopUp_InitialMouseUp(ByVal Button As Integer)
Dim M As cwMenu
  Set M = MenuUnderCursor
  If Not M Is Nothing Then M.W_MouseUp Button, 0, mMouseX, mMouseY
End Sub
 
Private Sub W_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
Dim SW As cWidgetBase
  If TypeOf Sender Is cwMenuItem Then Set SW = Sender.Widget Else Exit Sub
  
  Select Case EventName
    Case "W_MouseUp"
      Set mActiveMenuItem = Sender
      If P3 < 0 Or P4 < 0 Or P3 > SW.Width Or P4 > SW.Height Then Exit Sub
      W_MouseUp CInt(P1), CInt(P2), SW.Left + P3, SW.Top + P4
    Case "W_MouseMove"
      Set mActiveMenuItem = Sender
      W_MouseMove CInt(P1), CInt(P2), SW.Left + P3, SW.Top + P4
    Case "W_MouseDown"
      Set mActiveMenuItem = Sender
      W_MouseDown CInt(P1), CInt(P2), SW.Left + P3, SW.Top + P4
      
    Case "W_GotFocus"
      Set mActiveMenuItem = Sender
      If Not mActiveMenuItem Is Sender Then
        If Not mSubMenu Is Nothing Then mSubMenu.DestroyPopup
      End If
      
    Case "ShowSubMenu"
      Set mActiveMenuItem = Sender
      ShowSubMenu P1
 
  End Select
'  Debug.Print "W_BubblingEvent", MenuItemKeyPath
End Sub

Private Sub W_AccessKeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    W_KeyDown KeyAscii, 0
    Exit Sub
  End If
  AccessKeyPress KeyAscii
End Sub
Public Sub AccessKeyPress(KeyAscii As Integer)
  If Not mSubMenu Is Nothing Then mSubMenu.AccessKeyPress KeyAscii: Exit Sub
  
  Dim MItem As cwMenuItem
  For Each MItem In Widgets
    MItem.Widget.AccessKeyPress KeyAscii
  Next MItem
End Sub

Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyDown KeyCode, Shift
End Sub
Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
  If Not mSubMenu Is Nothing Then mSubMenu.KeyDown KeyCode, Shift: Exit Sub
  
  Dim i As Long, j As Long, MItem As cwMenuItem, MItemIdx As Long, RM As cwMenu
 
  For Each MItem In Widgets
    If MItem.Widget.Focused Then Exit For
  Next MItem
  If Not MItem Is Nothing Then MItemIdx = Widgets.GetOneBasedChildIndexByKey(MItem.Widget.Key)
  
  Select Case KeyCode
    Case vbKeyEscape: mDestroyedByKeyPress = True: DestroyPopup
    Case vbKeyLeft
        If RootMenu Is Me Then
          RaiseEvent MenuBarEntryShift(True)
          W.RaiseBubblingEvent Me, "MenuBarEntryShift", True
        Else
          DestroyPopup
        End If
    Case vbKeyRight
      If Not MItem Is Nothing Then
        If MItem.Widget.Enabled And Not MItem.SubMenuDS Is Nothing Then
          ShowSubMenu MItem, True
          Exit Sub
        End If
      End If
      If RootMenu Is Me Then
        RaiseEvent MenuBarEntryShift(False)
        W.RaiseBubblingEvent Me, "MenuBarEntryShift", False
      End If
    Case vbKeyUp
      For i = 1 To Widgets.Count
         j = MItemIdx - i
         If j < 1 Then j = Widgets.Count
        If Widgets(j).Widget.Enabled Then Widgets(j).Widget.SetFocus: Exit For
      Next i
    Case vbKeyDown
      For i = 1 To Widgets.Count
         j = MItemIdx + i
         If j > Widgets.Count Then j = 1
        If Widgets(j).Widget.Enabled Then Widgets(j).Widget.SetFocus: Exit For
      Next i
      
    Case vbKeyReturn
      
      If Not MItem Is Nothing Then
        If MItem.Widget.Enabled Then
          If Not RootMenu Is Nothing Then
            mDestroyedByKeyPress = True
            If MItem.IsOption Then UnCheckAllMenuOptions MItem
            RootMenu.RaiseClick MenuItemKeyPath
          End If
        End If
      End If
      If Not RootMenu Is Nothing Then mDestroyedByKeyPress = True: RootMenu.DestroyPopup
  End Select
End Sub

Private Sub W_KeyPress(KeyAscii As Integer)
  KeyPress KeyAscii
End Sub
Public Sub KeyPress(KeyAscii As Integer)
  If Not mSubMenu Is Nothing Then mSubMenu.KeyPress KeyAscii: Exit Sub
  Dim MItem As cwMenuItem
  For Each MItem In Widgets
    MItem.Widget.AccessKeyPress KeyAscii
  Next MItem
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  UpdateLastMouseInfo x, y
End Sub
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  UpdateLastMouseInfo x, y
End Sub
Public Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button < 0 Or x <= 0 Or y <= 0 Then Exit Sub Else UpdateLastMouseInfo x, y
  
  If mSubMenu Is Nothing And Not mActiveMenuItem Is Nothing Then
      If Not mActiveMenuItem.SubMenuDS Is Nothing And mActiveMenuItem.Widget.Enabled Then
        ShowSubMenu mActiveMenuItem
      ElseIf mActiveMenuItem.Widget.Enabled Then
        If Not RootMenu Is Nothing Then
          If mActiveMenuItem.IsOption Then UnCheckAllMenuOptions mActiveMenuItem    'First we clear all the options
          RootMenu.RaiseClick MenuItemKeyPath                                       'And in here we set the current one
        End If
      End If
  End If
End Sub

Friend Sub UpdateLastMouseInfo(ByVal x As Single, ByVal y As Single)
  If Not mSubMenu Is Nothing Then
    If mSubMenu.MenuKeyPath <> MenuItemKeyPath Then
      mSubMenu.DestroyPopup
      Set mSubMenu = Nothing
    End If
  End If
  mMouseX = x: mMouseY = y
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  CC.SetSourceColor W.BackColor
  CC.Paint
  CC.SetLineWidth 1, True
  W.Alpha = 0.7
  Cairo.Theme.DrawTo CC, W, thmTypeBorder, 0, 0, 0, dx_Aligned, dy_Aligned
  W.Alpha = 1
End Sub

Private Sub UnCheckAllMenuOptions(MItem As cwMenuItem)
'Unchecks all the menu options items that are in the same level of hierarchy of MItem
Dim MItemX As cwMenuItem, i As Long, MIdx As Long
  If MItem.Widget.Parent Is Nothing Then Exit Sub
  If MItem.Widget.Parent.ChildCount = 0 Then Exit Sub
  
  MIdx = MItem.Widget.Parent.Widgets.GetOneBasedChildIndexByKey(MItem.Widget.Key)
  For i = MIdx - 1 To 1 Step -1 'first we go upwards
    If TypeOf MItem.Widget.Parent.Widgets(i) Is cwMenuItem Then
      Set MItemX = MItem.Widget.Parent.Widgets(i)
      If MItemX.Caption = "-" Then Exit For
      If MItemX.IsOption Then
        mDataSource.SubItemByKey(MItemX.Widget.Key).Checked = False
      End If
    End If
  Next
  For i = MIdx To MItem.Widget.Parent.Widgets.Count  'now we loop downwards
    If TypeOf MItem.Widget.Parent.Widgets(i) Is cwMenuItem Then
      Set MItemX = MItem.Widget.Parent.Widgets(i)
      If MItemX.Caption = "-" Then Exit For
      If MItemX.IsOption Then
        mDataSource.SubItemByKey(MItemX.Widget.Key).Checked = False
      End If
    End If
  Next
End Sub

Private Sub Class_Terminate()
  Set SC = Nothing
  If Not fPopUp Is Nothing Then fPopUp.Unload
End Sub

