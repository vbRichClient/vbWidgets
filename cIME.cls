VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event HandleIMEPositioning(FocusedWidget As cWidgetBase, AllowIME As Boolean)
Event HandleIMEChar(FocusedWidget As cWidgetBase, ByVal IMEKeyCode As Integer, IMEWChar As String)

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type COMPOSITIONFORM
  dwStyle As Long
  ptCurrentPos As POINTAPI
  rcArea As RECT
End Type
 
Private Declare Function ImmAssociateContextEx Lib "imm32" (ByVal hWnd As Long, ByVal hIMC As Long, ByVal dwFlags As Long) As Long
Private Declare Function ImmGetContext Lib "imm32" (ByVal hWnd As Long) As Long
Private Declare Function ImmReleaseContext Lib "imm32" (ByVal hWnd As Long, ByVal hIMC As Long) As Long
Private Declare Function ImmSetOpenStatus Lib "imm32" (ByVal hIMC As Long, ByVal B As Long) As Long
Private Declare Function ImmGetOpenStatus Lib "imm32" (ByVal hIMC As Long) As Long
Private Declare Function ImmSetCompositionWindow Lib "imm32" (ByVal hIMC As Long, lpCompositionForm As COMPOSITIONFORM) As Long
Private Declare Function ImmGetCompositionStringW Lib "imm32" (ByVal hIMC As Long, ByVal Flags As Long, ByVal pStr As Long, ByVal SLen As Long) As Long
Private Declare Function ImmGetDefaultIMEWnd Lib "imm32" (ByVal hWnd As Long) As Long
 
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal Flags As Long) As Long

Private hWnd As Long, WithEvents SC As cSubClass, WithEvents tmrFoc As cTimer
Attribute SC.VB_VarHelpID = -1
Attribute tmrFoc.VB_VarHelpID = -1
 
Public Sub BindToForm(Form As cWidgetForm)
  hWnd = Form.hWnd
  Set SC = New_c.SubClass
      SC.Hook hWnd
  Set tmrFoc = New_c.Timer(30, True, "")
End Sub

Public Sub SwitchOpenStatus(ByVal bOpen As Boolean)
Dim hIMC As Long
    hIMC = ImmGetContext(hWnd): If hIMC = 0 Then Exit Sub
    If ImmGetOpenStatus(hIMC) <> IIf(bOpen, 1, 0) Then ImmSetOpenStatus hIMC, IIf(bOpen, 1, 0)
    ImmReleaseContext hWnd, hIMC
End Sub

Public Sub SetPosition(ByVal x As Long, ByVal y As Long)
Const CFS_POINT = 2, CFS_FORCE_POSITION = &H20
Dim hIMC As Long, CF As COMPOSITIONFORM
    hIMC = ImmGetContext(hWnd): If hIMC = 0 Then Exit Sub
        CF.dwStyle = CFS_FORCE_POSITION
        CF.ptCurrentPos.x = x
        CF.ptCurrentPos.y = y
        ImmSetCompositionWindow hIMC, CF
    ImmReleaseContext hWnd, hIMC
End Sub
 
Public Function GetCompositionString() As String
Const GCS_COMPSTR = 8
Dim hIMC As Long, SLen As Long
    hIMC = ImmGetContext(hWnd): If hIMC = 0 Then Exit Function
        SLen = ImmGetCompositionStringW(hIMC, GCS_COMPSTR, 0, 0) \ 2
        GetCompositionString = Space$(SLen)
        ImmGetCompositionStringW hIMC, GCS_COMPSTR, StrPtr(GetCompositionString), LenB(GetCompositionString)
    ImmReleaseContext hWnd, hIMC
End Function

Private Sub SC_WindowProc(Result As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
Const WM_IME_SETCONTEXT = 641, WM_IME_STARTCOMPOSITION = 269, WM_IME_CHAR = 646
On Error GoTo 1

   Select Case Msg
      Case WM_IME_SETCONTEXT
        SwitchOpenStatus wParam
        
      Case WM_IME_STARTCOMPOSITION
        HandleIMEPos
      
      Case WM_IME_CHAR
        Dim WFoc As cWidgetBase, KeyCode As Integer
        Set WFoc = FocusedWidget
        KeyCode = CInt("&H" & Hex(wParam And &HFFFF&))
        If Not WFoc Is Nothing Then If WFoc.Key = tmrFoc.Tag Then RaiseEvent HandleIMEChar(WFoc, KeyCode, ChrW(KeyCode))
        Exit Sub 'handled ourselves - so we skip the default message-handler at the end of this function
   End Select
  
1: Result = SC.CallWindowProc(Msg, wParam, lParam)
End Sub
 
Private Sub tmrFoc_Timer()
  HandleIMEPos
End Sub

Private Function FocusedWidget() As cWidgetBase
  If Cairo.WidgetForms.Exists(hWnd) Then Set FocusedWidget = Cairo.WidgetForms(hWnd).WidgetRoot.ActiveWidget
End Function

Private Sub HandleIMEPos()
Dim WFoc As cWidgetBase, AllowIME As Boolean
On Error GoTo 1

   Set WFoc = FocusedWidget
   If WFoc Is Nothing Then
     tmrFoc.Tag = ""
   Else
     RaiseEvent HandleIMEPositioning(WFoc, AllowIME)
     If AllowIME Then tmrFoc.Tag = WFoc.Key
   End If
  
1: SwitchOpenStatus AllowIME
End Sub
  
Private Sub Class_Terminate()
  Set tmrFoc = Nothing
  Set SC = Nothing
  SwitchOpenStatus False
End Sub

