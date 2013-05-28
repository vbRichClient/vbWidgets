VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwCueBannerProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------------------+
'   The purpose of a cue banner is to provide a hint to the user as |
'   to the purpose of an edit control.                              |
'-------------------------------------------------------------------+
'   Properties:                                                     |
'       Caption     text of the hint                                |
'       TextBox     instance of the cwTextBox-class                 |
'-------------------------------------------------------------------+

Private p_Caption                       As String

Private WithEvents p_TextBox            As cwTextBox
Attribute p_TextBox.VB_VarHelpID = -1
Private WithEvents p_TextBoxWidget      As cWidgetBase
Attribute p_TextBoxWidget.VB_VarHelpID = -1

Private lblCue                          As cwLabel
Attribute lblCue.VB_VarHelpID = -1
Private WithEvents lblCueWidget         As cWidgetBase
Attribute lblCueWidget.VB_VarHelpID = -1

Private Sub Class_Terminate()

    Set p_TextBox = Nothing
    Set lblCue = Nothing

End Sub


Public Property Get Caption() As String
Attribute Caption.VB_Description = "text of the hint"

    Caption = p_Caption

End Property

Public Property Let Caption(ByVal Caption As String)

    p_Caption = Caption
    configureContent

End Property

Public Property Get TextBox() As cwTextBox
Attribute TextBox.VB_Description = "instance of the cwTextBox-class"

    Set TextBox = p_TextBox

End Property

Public Property Set TextBox(TextBox As cwTextBox)
    Dim h                               As Long
    Dim w                               As Long

    Set p_TextBox = TextBox
    Set p_TextBoxWidget = TextBox.Widget

    w = p_TextBox.Widget.ScaleWidth
    h = p_TextBox.Widget.ScaleHeight

    Set lblCue = p_TextBox.Widgets.Add(New cwLabel, "lblCue", 0, 0, w, h)
    Set lblCueWidget = lblCue.Widget
    With lblCue
        .Caption = p_Caption
        .VAlign = 0
        .InnerSpace = 3
        .WordWrap = True
        .Widget.Visible = True
        .Widget.BackColor = -1
        .Widget.ForeColor = RGB(100, 100, 100)
        .Widget.FontName = "Arial"
        .Widget.FontSize = p_TextBoxWidget.FontSize
        .Widget.FontItalic = True
        .Widget.Enabled = True
    End With

End Property







Private Sub lblCueWidget_Click()
    configureContent            ' lblCue.Widget.Visible = False
    p_TextBox.Widget.SetFocus
End Sub

Private Sub lblCueWidget_GotFocus()
    configureContent            ' lblCue.Widget.Visible = False
    p_TextBox.Widget.SetFocus
End Sub

Private Sub p_TextBox_Change()
    lblCue.Widget.Visible = False
    p_TextBox.Widget.SetFocus
End Sub

Private Sub configureContent()
    If Not p_TextBox Is Nothing Then
        If Len(p_Caption) > 0 Then
            If Len(p_TextBox.Text) = 0 Then
                lblCue.Widget.Visible = True
            Else
                lblCue.Widget.Visible = False
            End If
        Else
            lblCue.Widget.Visible = False
        End If
    End If
End Sub

Private Sub p_TextBoxWidget_LostFocus()
    configureContent
End Sub
