VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long

'all these Events here are mainly re-delegations of the cWebKit-Events
Event SetPageTitle(PageTitle As String)
Event LoadCommited(URL As String)
Event LoadFinished(ByVal BackCount As Long, ByVal ForwardCount As Long)
Event NewWindowRequest(NewWebKitInstance As cWebKit, ByVal OpenAsTab As Boolean)
Event DecideAboutResourceLoadRequest(RequestURL As String, Cancel As Boolean)
Event DecidePolicyForMimeType(MimeTypeLeftPart As String, MimeTypeRightPart As String, RequestURL As String, NavPolicy As vbRichClient5.NavigationPolicyDecision)
Event DecidePolicyForNavigation(NavInfo As cNavigationInfo, RequestURL As String, NavPolicy As vbRichClient5.NavigationPolicyDecision)
Event DecidePolicyForNewWindow(NavInfo As cNavigationInfo, RequestURL As String, NavPolicy As vbRichClient5.NavigationPolicyDecision)
Event ShowAuthenticationDialog(ByVal URL As String, UserName As String, PassWord As String)
Event ShowJSAlertDlg(Message As String)
Event ShowJSConfirmDlg(Message As String, Cancel As Boolean)
Event ShowJSInputDlg(Message As String, DefaultText As String, ReturnText As String)
Event JSEventCallBack(EventName As String, JSONEventObj As cCollection, JSONResult As cCollection)
Event FindFirst()
Event FindNext()
Event FindPrevious()
Event PrintRequest()

Public WithEvents WebKit As cWebKit 'this is the WebKit-Binding-class
Attribute WebKit.VB_VarHelpID = -1
Private mViewInitialized As Boolean

'**** cairo-widget implementation-conventions (also use a cw-Prefix for your 'cwMyWidget'-ClassNames) ****
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
  W.LockRefresh = True 'this normally not done, but this Widget does no drawing in itself, so we inform the Engine, that it can be ignored in the cairo-RenderLoop
End Sub
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property
'**** end of cairo-widget implementation-conventions ****

Private Sub Class_Terminate()
  Set WebKit = Nothing
End Sub

Public Sub Navigate2(URL As String, CachePolicy As NavigationCachePolicy, ByVal TimeOutIntervalSec As Double)
  WebKit.Navigate2 URL, CachePolicy, TimeOutIntervalSec
End Sub

Public Sub DelegateFocus()
  If Not WebKit Is Nothing Then WebKit.DelegateFocus
End Sub
Public Sub SynchronizeCoords()
Dim xScreen As Single, yScreen As Single, PT(0 To 1) As Long

  If WebKit Is Nothing Or W.Root Is Nothing Then Exit Sub
  W.CalculateScreenOffsets xScreen, yScreen
  PT(0) = xScreen: PT(1) = yScreen
  ScreenToClient W.Root.hWnd, PT(0)

'  W.Root.Widget.LockRefresh = True
    
    WebKit.Move PT(0), PT(1), W.ScaleWidthPxl, W.ScaleHeightPxl
    DoEvents
'    Debug.Print "cwBrowserMove", IsWindowEnabled(W.hWnd), IsWindowVisible(W.hWnd)
'  W.Root.Widget.LockRefresh = False
End Sub

Private Sub W_AddedToHierarchy()
  InitWebKit
End Sub

Private Sub InitWebKit()
  If mViewInitialized Then Exit Sub
  If W.Root Is Nothing Then Exit Sub

  mViewInitialized = True
  
  If WebKit Is Nothing Then Set WebKit = New_c.WebKit
  WebKit.InitializeView W.Root.hWnd, W.Left, W.Top, W.Width, W.Height
 
'Debug.Print "initwebkit", W.hWnd
'  EnableWindow W.hWnd, 1
'  EnableWindow GetParent(W.hWnd), 1
'  WebKit.InitializeView W.hWnd, 0, 0, W.ScaleWidthPxl, W.ScaleHeightPxl
End Sub

'the following three Event-Handlers ensure, that the Widget- and the BrowserWindow-Focus is "synced"
Private Sub W_GotFocus()
  If WebKit Is Nothing Then Exit Sub
  If Not WebKit.Focused Then WebKit.DelegateFocus
End Sub
 
Private Sub W_LostFocus()
  If WebKit Is Nothing Then Exit Sub
  If WebKit.Focused Then W.Root.Widget.SetFocus
End Sub
 
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
SynchronizeCoords
End Sub

Private Sub WebKit_DecideAboutResourceLoadRequest(RequestURL As String, Cancel As Boolean)
  RaiseEvent DecideAboutResourceLoadRequest(RequestURL, Cancel)
  
'  If Cancel Then Debug.Print "cancelled Resource-Request: --> ", RequestURL
End Sub

Private Sub WebKit_DecidePolicyForMimeType(MimeTypeLeftPart As String, MimeTypeRightPart As String, RequestURL As String, NavPolicy As vbRichClient5.NavigationPolicyDecision)
  RaiseEvent DecidePolicyForMimeType(MimeTypeLeftPart, MimeTypeRightPart, RequestURL, NavPolicy)
End Sub

Private Sub WebKit_DecidePolicyForNavigation(NavInfo As cNavigationInfo, RequestURL As String, NavPolicy As vbRichClient5.NavigationPolicyDecision)
  RaiseEvent DecidePolicyForNavigation(NavInfo, RequestURL, NavPolicy)
End Sub

Private Sub WebKit_DecidePolicyForNewWindow(NavInfo As cNavigationInfo, RequestURL As String, NavPolicy As vbRichClient5.NavigationPolicyDecision)
  RaiseEvent DecidePolicyForNewWindow(NavInfo, RequestURL, NavPolicy)
End Sub

Private Sub WebKit_DecideAboutDownload(ByVal URL As String, Cancel As Boolean)
  Debug.Print "DecideAboutDownload"
End Sub
Private Sub WebKit_DownloadStart(DownloadObj As cDownload)
  Debug.Print "DownloadStart"
End Sub
Private Sub WebKit_DownloadProgress(DownloadObj As cDownload, ByVal Percent As Single, ByVal StatusCode As Long, StatusDescription As String, StatusValue As String)
  Debug.Print "DownloadProgress", Percent
End Sub
Private Sub WebKit_DownloadComplete(DownloadObj As cDownload, ByVal ErrNum As Long, ErrString As String)
  Debug.Print "DownloadComplete with: ",
  If ErrNum = 0 Then
    Debug.Print "DataLength: " & DownloadObj.GetContentLen
  Else
    Debug.Print ErrString
  End If
End Sub

Private Sub WebKit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, ByVal xWithScrollOffs As Long, ByVal yWithScrollOffs As Long)
  W.MouseDown Button, Shift, x, y 'redelegate mousedown to the Widget
End Sub

Private Sub WebKit_PrintRequest()
  RaiseEvent PrintRequest
End Sub

Private Sub WebKit_ShowAuthenticationDialog(ByVal URL As String, UserName As String, PassWord As String)
  RaiseEvent ShowAuthenticationDialog(URL, UserName, PassWord)
End Sub

Private Sub WebKit_ShowJSAlertDlg(Message As String)
  RaiseEvent ShowJSAlertDlg(Message)
End Sub

Private Sub WebKit_ShowJSConfirmDlg(Message As String, Cancel As Boolean)
  RaiseEvent ShowJSConfirmDlg(Message, Cancel)
End Sub

Private Sub WebKit_ShowJSInputDlg(Message As String, DefaultText As String, ReturnText As String)
  RaiseEvent ShowJSInputDlg(Message, DefaultText, ReturnText)
End Sub

Private Sub WebKit_UserDragOverDst(Data As cDataObject, ByVal AllowedEffects As vbRichClient5.WidgetDropEffects, Effect As vbRichClient5.WidgetDropEffects, Button As Integer, Shift As Integer, x As Single, y As Single)
Effect = 3
Debug.Print "UserDropOverHandling", Button, Shift, x, y, Effect
End Sub
Private Sub WebKit_UserDragDropDst(Data As cDataObject, ByVal AllowedEffects As vbRichClient5.WidgetDropEffects, Effect As vbRichClient5.WidgetDropEffects, Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print "UserDropHandling", Button, Shift, x, y, Effect
End Sub

Private Sub WebKit_WebViewDragOverDst(Data As cDataObject, ByVal AllowedEffects As vbRichClient5.WebKitDropEffects, Effect As vbRichClient5.WebKitDropEffects, Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print "WebKitDropOverHandling", Button, Shift, x, y, Effect
End Sub
Private Sub WebKit_WebViewDragDropDst(Data As cDataObject, ByVal AllowedEffects As vbRichClient5.WebKitDropEffects, Effect As vbRichClient5.WebKitDropEffects, Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print "WebKitDropHandling", Button, Shift, x, y, Effect
End Sub


Private Sub WebKit_FindFirst()
  RaiseEvent FindFirst
End Sub
Private Sub WebKit_FindNext()
  RaiseEvent FindNext
End Sub
Private Sub WebKit_FindPrevious()
  RaiseEvent FindPrevious
End Sub

Private Sub WebKit_GotFocus()
  If Not W.Focused Then W.SetFocus
End Sub

'same thing here - let's sync the Browser-Window, according to our hosting Widgets coordinates
Private Sub W_Resize()
  SynchronizeCoords
End Sub
Private Sub W_Moving()
  SynchronizeCoords
End Sub

'and also the same here...
Private Sub W_Show()
  If Not WebKit Is Nothing Then WebKit.Visible = W.Visible
End Sub
Private Sub W_Hide()
  If Not WebKit Is Nothing Then WebKit.Visible = False
End Sub


'this Event is called, when the internal "Browser-Tabbing" is reaching its bounds
'(so we have the chance, to hand-over the focus to our "normal", surrounding Widgets)
Private Sub WebKit_FocusWantsExit(ByVal MoveFocusForward As Boolean)
Dim i&
  If W.Parent Is Nothing Then Exit Sub
  If W.Parent.ChildCount <= 1 Then Exit Sub

  i = W.Parent.Widgets.GetOneBasedChildIndexByKey(W.Key)
  
  If MoveFocusForward Then
    If i = W.Parent.ChildCount Then i = 1 Else i = i + 1
  Else
    If i = 1 Then i = W.Parent.ChildCount Else i = i - 1
  End If

  W.Root.SetFocus W.Parent.Widgets(i).Widget 'hand-over the focus
End Sub


Private Sub WebKit_JSEventCallBack(EventName As String, JSONEventObj As cCollection, JSONResult As cCollection)
  RaiseEvent JSEventCallBack(EventName, JSONEventObj, JSONResult)
End Sub

Private Sub WebKit_LoadFinished(ByVal BackCount As Long, ByVal ForwardCount As Long)
  RaiseEvent LoadFinished(BackCount, ForwardCount)
End Sub

Private Sub WebKit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, ByVal xWithScrollOffs As Long, ByVal yWithScrollOffs As Long)
  If WebKit Is Nothing Then Exit Sub
  WebKit.Enabled = Not W.Root.DesignMode
End Sub
 
Private Sub WebKit_MouseMoveOverElement(ElmtInfo As cElementInfo, ByVal ModifierFlags As Long)
  If Len(ElmtInfo.Title) Then
    W.MouseEnter Nothing
    W.ToolTip = ElmtInfo.Title
  Else
    W.MouseLeave Nothing
  End If
End Sub

Private Sub WebKit_RegisteredEventCallBack(ByVal ElementKey As String, ByVal EventName As String)
'  Debug.Print ElementKey, EventName
End Sub

Private Sub WebKit_SetPageTitle(PageTitle As String)
  RaiseEvent SetPageTitle(PageTitle)
End Sub

Private Sub WebKit_LoadCommited(URL As String)
  RaiseEvent LoadCommited(URL)
End Sub

Private Sub WebKit_NewWindowRequest(NewWebKitInstance As cWebKit, ByVal OpenAsTab As Boolean)
  RaiseEvent NewWindowRequest(NewWebKitInstance, OpenAsTab)
End Sub

