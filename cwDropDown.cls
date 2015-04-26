VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event PopupEscapeKeyPressed(DestroyPopup As Boolean)
Event PopupReturnKeyPressed(DestroyPopup As Boolean)
Event PopupWidgetRequest(WidgetInstance As Object, WidthPxl As Single, HeightPxl As Single)
Event PopupDestroyed()

Private WithEvents PopUp As cfPopUp 'this PopupForm-HelperClass is hosted here in the vbWidgets.dll-Project
Attribute PopUp.VB_VarHelpID = -1
Public WithEvents Btn As cwButton 'we inherit visually from a preimplemented Widget (in this case a cwButton in CheckButton-Mode)
Attribute Btn.VB_VarHelpID = -1
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private mScreenX As Single, mScreenY As Single, mWindowUnderStartPos As Long
Private mPopupWidget As Object

Private Sub Class_Initialize()
  Set Btn = New cwButton
      Btn.ButtonStyle = CheckButton
      Btn.HAlignment = vbLeftJustify
      Btn.ImageSize = 16
  Set W = Btn.Widget '<- in our visual inheritance-scenario, we don't create a new Cairo.WidgetBase, but use the one from Btn.Widget instead
      W.ImplementsWheelMessages = True
      
  W.RuntimePropertiesCommaSeparated = "Text"
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Public Property Get Caption() As String
  Caption = Btn.Caption
End Property
Public Property Let Caption(ByVal NewValue As String)
  Btn.Caption = NewValue
End Property

Public Property Get PopupWidget() As Object
  Set PopupWidget = mPopupWidget
End Property

Public Property Get Text() As String
  Text = Btn.Caption
End Property

Public Sub DestroyPopup()
  Set PopUp = Nothing
  Set mPopupWidget = Nothing
  If Not Btn Is Nothing Then Btn.Value = 0
  RaiseEvent PopupDestroyed
  W.RaiseBubblingEvent Me, "PopupDestroyed"
End Sub
 
'------ Event-checks, which ensure Popup-Deactivation under certain conditions -----
Private Sub PopUp_AppDeactivate()
  Btn.Value = 0
End Sub
Private Sub PopUp_MouseClickOutside()
  If Not W.MouseOver Then Btn.Value = 0
End Sub
Private Sub PopUp_CheckForAdditionalCloseConditions()
  If mWindowUnderStartPos <> W.Root.WindowFromPoint(mScreenX, mScreenY) Then Btn.Value = 0
End Sub
Private Sub W_ContainerResize()
  Btn.Value = 0
End Sub
Private Sub W_LostFocus()
  Btn.Value = 0
End Sub
'---------------- End of Popup-Deactivation-Checks ------------------

're-delegation-messages into the currently "popped-up" Widget
Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DestroyPopup As Boolean
  If mPopupWidget Is Nothing Then Exit Sub
  
  If KeyCode = vbKeyEscape Then
    DestroyPopup = True '<- set the default in this case to "destroy"
    RaiseEvent PopupEscapeKeyPressed(DestroyPopup)
    W.RaiseBubblingEvent Me, "PopupEscapeKeyPressed", DestroyPopup
  ElseIf KeyCode = vbKeyReturn Then
    RaiseEvent PopupReturnKeyPressed(DestroyPopup)
    W.RaiseBubblingEvent Me, "PopupReturnKeyPressed", DestroyPopup
  End If
  
  If DestroyPopup Then
    Btn.Value = 0
  Else
    If Not mPopupWidget Is Nothing Then mPopupWidget.Widget.KeyDown KeyCode, Shift
  End If
End Sub
Private Sub W_KeyPress(KeyAscii As Integer)
  If mPopupWidget Is Nothing Then Exit Sub
  mPopupWidget.Widget.KeyPress KeyAscii
End Sub
Private Sub W_KeyUp(KeyCode As Integer, Shift As Integer)
  If mPopupWidget Is Nothing Then Exit Sub
  mPopupWidget.Widget.KeyUp KeyCode, Shift
End Sub
Private Sub W_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
  If mPopupWidget Is Nothing Then Exit Sub
  mPopupWidget.Widget.MouseEnter Nothing
  mPopupWidget.Widget.MouseWheel MouseKeys, Rotation, xAbs, yAbs
End Sub


Private Sub Btn_Click()
Dim PopUpOffsX As Single, PopUpOffsY As Single, WidthPxl As Single, HeightPxl As Single
  If Btn.Value = 1 Then 'Check-Button is checked (down)
  
    W.CalculateScreenDimensions WidthPxl, HeightPxl  'get the PixelSize of the Widget for the current Screen
    RaiseEvent PopupWidgetRequest(mPopupWidget, WidthPxl, HeightPxl)
    If mPopupWidget Is Nothing Then 'apparently nobody filled something into the Byref-Params, so we try again with the Bubbling-Event
      W.RaiseBubblingEvent Me, "PopupWidgetRequest", mPopupWidget, WidthPxl, HeightPxl
    End If
    If mPopupWidget Is Nothing Then Btn.Value = 0: Exit Sub 'still nobody provided something, so we exit here
    
    W.CalculateScreenOffsets mScreenX, mScreenY
    mWindowUnderStartPos = W.Root.WindowFromPoint(mScreenX, mScreenY)

    W.CalculateScreenPopupOffsets WidthPxl, HeightPxl, PopUpOffsX, PopUpOffsY

    Set PopUp = New cfPopUp
    PopUp.Load mPopupWidget, PopUpOffsX, PopUpOffsY, WidthPxl, HeightPxl, W.Zoom, True
    PopUp.Show
  Else
    DestroyPopup
  End If
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  W.MouseUp Button, Shift, x, y
End Sub
 
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
Dim ArrowSize As Single, LeftOffs As Single, RightOffs As Single
  ArrowSize = 8
  RightOffs = 2.5 * ArrowSize
  LeftOffs = dx_Aligned - RightOffs
  
  Cairo.Theme.DrawTo CC, W, thmTypeSeparatorLine, 0, LeftOffs, 3, 1, dy_Aligned - 4
  Cairo.Theme.DrawTo CC, W, thmTypeArrow, 0, LeftOffs + (RightOffs - ArrowSize) / 2, (dy_Aligned - ArrowSize) / 2, ArrowSize, ArrowSize, , thmDirectionDown
End Sub
