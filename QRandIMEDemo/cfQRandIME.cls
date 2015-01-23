VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfQRandIME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Form As cWidgetForm, WithEvents IME As cIME
Attribute Form.VB_VarHelpID = -1
Attribute IME.VB_VarHelpID = -1

Public QREnc As New cQREncode, QRDec As New cQRDecode 'the two (non-visible) QR-CodecClass-Vars
Public TB1 As cwTBoxWrap, TB2 As cwTBoxWrap 'the two TextBox-Wrapper-Classes
Public QR1 As cwQRSimple, QR2 As cwQRSimple 'the two QR-Widgets
 
Private Sub Class_Initialize()
  Set Form = Cairo.WidgetForms.Create(vbFixedDialog, "QR-Widgets and IME-Window-Positioning", , 800, 600)
      Form.IconImageKey = "QRico2"
      Form.WidgetRoot.ImageKey = "bgPatForm"
      Form.WidgetRoot.ImageKeyRenderBehaviour = ImgKeyRenderRepeat
      
  Set IME = New cIME 'create the vbWidgets.cIME-instance
      IME.BindToForm Form '...and bind our cWidgetForm-instance to it (IME will throw two Events at us then)
End Sub

Private Sub Form_Load() 'handle Widget-Creation and -Adding on this Form
  Form.Widgets.Add(New cwSeparatorLabel, "Sep1", 11, 8, Form.ScaleWidth - 22, 42).SetCaptionAndImageKey "EditBox-DemoArea", "Edit", &H11AA66
    Set TB1 = Form.Widgets.Add(New cwTBoxWrap, "TB1", 25, 60, 280, 38)
        TB1.TBox.CueBannerText = "Session-Login..."
        TB1.Widget.ImageKey = "session1"
    Set TB2 = Form.Widgets.Add(New cwTBoxWrap, "TB2", 325, 60, 280, 38)
        TB2.TBox.CueBannerText = "Place some Info here..."
        TB2.Widget.ImageKey = "info1"
      
  Form.Widgets.Add(New cwSeparatorLabel, "Sep2", 11, 155, Form.ScaleWidth - 22, 42).SetCaptionAndImageKey "QRCode-DemoArea", "Preview", &H1030EE
    Set QR1 = Form.Widgets.Add(New cwQRSimple, "QR1", 25, 240, 250, 220)
    Set QR2 = Form.Widgets.Add(New cwQRSimple, "QR2", 325, 280, 250, 220)
End Sub

Private Sub Form_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
  If EventName = "Change" And TypeOf Sender Is cwTextBox Then 'we handle the Change-Event of the QRWidget-Child-Textboxes here
    If Not (Sender Is QR1.TBox Or Sender Is QR2.TBox) Then Exit Sub
    
    'resolve to the (TextBox-Hosting) cwQRSimple-Widget in question
    Dim QR As cwQRSimple: Set QR = IIf(Sender Is QR1.TBox, QR1, QR2)
    
    'Encode the current Text of our QR-Widget - and place the returned Pixel-Surface in QR.QRSrf
    Set QR.QRSrf = QREnc.QREncode(New_c.Crypt.VBStringToUTF8(QR.Text))
 
    'to verify, we perform a true Decoding of the QR-Text from the Pixels of the just created QR-Widgets QR-Surface
    QRDec.DecodeFromSurface QR.QRSrf
    'and reflect this decoded Unicode-StringResult in the Caption of the QR-Widget (so, ideally QR.Caption should match QR.Text)
    If QRDec.QRDataLen(0) Then QR.Caption = New_c.Crypt.UTF8ToVBString(QRDec.QRData(0)) Else QR.Caption = ""
  End If
  
  'the QR-Widgets (cwQRSimple) are moveable - and in case they have an active IME-Window, we will move that too
  If EventName = "W_Moving" And TypeOf Sender Is cwQRSimple Then IME_HandleIMEPositioning Sender.TBox.Widget, True
End Sub

Private Sub IME_HandleIMEPositioning(FocusedWidget As cWidgetBase, AllowIME As Boolean)
  If TypeOf FocusedWidget.Object Is cwTextBox Then
    AllowIME = True '<- here we allow IME-Windows only for cwTextBox-Widgets (but we could also allow IME on other Widget-Types)
    IME.SetPosition FocusedWidget.AbsLeftPxl + 3, FocusedWidget.AbsTopPxl + FocusedWidget.ScaleHeightPxl + 4
  End If
End Sub

Private Sub IME_HandleIMEChar(FocusedWidget As cWidgetBase, ByVal IMEKeyCode As Integer, IMEWChar As String)
  FocusedWidget.KeyPress IMEKeyCode 'simply delegate the incoming IMEKeyCode into the Widget in question
  'the above is the more generic delegation-method into any Widget (which are all derived from cWidgetBase)
  
  '*alternatively* (for cwTextBoxes, which is the only Widget-Type we allow IME for in this Demo here)
  'we could also use:
'  Dim TB As cwTextBox
'  Set TB = FocusedWidget.Object
'      TB.SelText = IMEWChar
End Sub
