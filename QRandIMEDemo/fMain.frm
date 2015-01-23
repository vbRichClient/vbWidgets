VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Normal VB-MainForm"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1995
      Left            =   2700
      ScaleHeight     =   1935
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1995
      Left            =   180
      ScaleHeight     =   1935
      ScaleWidth      =   2175
      TabIndex        =   1
      Top             =   1200
      Width           =   2235
   End
   Begin VB.CommandButton cmdShowRC5IMEForm 
      Caption         =   "Show RC5-IME-Form modally"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   2835
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private VBFormAlreadyUnloaded As Boolean

Private Sub cmdShowRC5IMEForm_Click()
  With New cfQRandIME ' instantiate the RC5-FormHosting-Class
  
    .Form.Show , Me 'this will create and show the RC5-Form with the VB-Form as the underlying Parent
    
    'now we enter the W-capable RC5-message-pump, which will loop "in  place" till the RC5-Form gets closed again
    Cairo.WidgetForms.EnterMessageLoop True, False
 
    'the RC5-Form was closed, so let's read-out the Public Vars of its hosting cf-Class
    If Not VBFormAlreadyUnloaded Then '<- ... read the comment in Form_Unload, on why we need to check this flag
      Set Picture1.Picture = .QR1.QRSrf.Picture
      Set Picture2.Picture = .QR2.QRSrf.Picture
    End If
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer) 'this can happen whilst the RC5-ChildForm is showing, ...
  VBFormAlreadyUnloaded = True  'so we set a Flag, to not implicitely load this VB-ParentForm again, when filling the Result-PicBoxes
End Sub

Private Sub Form_Terminate() 'the usual RC5-cleanup call (when the last VB-Form was going out of scope)
  If Forms.Count = 0 Then New_c.CleanupRichClientDll
End Sub

