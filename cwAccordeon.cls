VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwAccordeon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event SelectionChanged(ActiveEntry As cwAccordeonEntry)

Private Alpha As Single, dx As Single, dy As Single
Private mCaption As String, mCurrentEntryIndex As Long
Private mCaptionHeight As Single, mAccEntryHeight As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
 
Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
  W.SetClientAreaOffsets 1, 1, 1, 1
  W.FontBold = True
  W.FontItalic = True
  W.ForwardFocus = True

  mCurrentEntryIndex = -1
  mCaptionHeight = 28
  mAccEntryHeight = 25
End Sub
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
  mCaption = NewValue
  W.Refresh
End Property

Public Sub AddEntry(Caption As String, IconResourceKey As String, AssociatedWidget As Object)
Dim Key As String, NewAccordeonEntry As New cwAccordeonEntry
  NewAccordeonEntry.Caption = Caption
  NewAccordeonEntry.Widget.ImageKey = IconResourceKey
  
  Key = Widgets.Count \ 2 & "_AccEntry"
  Widgets.Add NewAccordeonEntry, Key, 0, 0, W.Width, mAccEntryHeight
  Widgets.Add AssociatedWidget, Key & "Assoc", 0, 0, W.Width, mAccEntryHeight, False
  
  AdjustEntries
End Sub

Private Sub AdjustEntries()
Dim i As Long, y As Single, yy As Single, WEntry As cWidgetBase
  If Widgets.Count = 0 Then Exit Sub
  W.LockRefresh = True
    y = mCaptionHeight
    For i = 0 To mCurrentEntryIndex
      'the cwAccordeon-Entry
      Set WEntry = Widgets(2 * i + 1).Widget
      WEntry.Move -1, y, W.Width, mAccEntryHeight
      y = y + mAccEntryHeight
      
      'the associated Widget
      Set WEntry = Widgets(2 * i + 2).Widget
      If i < mCurrentEntryIndex Then WEntry.Visible = False
    Next i
  
    yy = W.ScaleHeight - (Widgets.Count \ 2 - mCurrentEntryIndex - 1) * mAccEntryHeight
    If Not WEntry Is Nothing Then
      WEntry.Move -1, y - 1, W.Width, yy - y + 2
      WEntry.Visible = True
'      WEntry.SetFocus
    End If
    
    y = yy
    For i = mCurrentEntryIndex + 1 To Widgets.Count \ 2 - 1
      'the cwAccordeon-Entry
      Set WEntry = Widgets(2 * i + 1).Widget
      WEntry.Move 0, y + 1, W.Width, mAccEntryHeight
      y = y + mAccEntryHeight
      
      'the associated Widget
      Set WEntry = Widgets(2 * i + 2).Widget
      WEntry.Visible = False
    Next i
  W.LockRefresh = False
End Sub

Public Property Get EntriesCount() As Long
  EntriesCount = Widgets.Count \ 2
End Property

Public Function EntryByIndex(ByVal EntryIndexZeroBased As Long) As cwAccordeonEntry
  If EntryIndexZeroBased < 0 Or EntryIndexZeroBased >= EntriesCount Then Exit Function
  Set EntryByIndex = Widgets(EntryIndexZeroBased * 2 + 1)
End Function

Public Property Get CurrentEntry() As cwAccordeonEntry
  If EntriesCount = 0 Then Exit Property
  Set CurrentEntry = Widgets(CurrentEntryIndex * 2 + 1)
End Property

Public Property Get CurrentEntryIndex() As Long
  CurrentEntryIndex = mCurrentEntryIndex
End Property
Public Property Let CurrentEntryIndex(ByVal NewValue As Long)
  If NewValue < -1 Or NewValue >= EntriesCount Then Exit Property
  If mCurrentEntryIndex = NewValue Then Exit Property
  mCurrentEntryIndex = NewValue
  
  RaiseEvent SelectionChanged(CurrentEntry)
  AdjustEntries
End Property

Private Sub W_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
Static FocusChange As Boolean
  If Sender Is Nothing Then Exit Sub
  If TypeOf Sender Is cwAccordeonEntry Then
    Select Case EventName
      Case "W_LostFocus"
        If Not CurrentEntry Is Sender Then FocusChange = True
      Case "W_Click"
        CurrentEntryIndex = Split(Sender.Widget.Key, "_")(0)
      Case "W_KeyPress"
        If P1 = vbKeySpace Then CurrentEntryIndex = Split(Sender.Widget.Key, "_")(0)
    End Select
    
  ElseIf Not CurrentEntry Is Nothing Then
    If Sender Is CurrentEntry.AssociatedWidget Then
      If EventName = "W_KeyDown" Then
        If Not (P1 = vbKeyTab And P2 = vbShiftMask) Then Exit Sub
        If FocusChange Then FocusChange = False: Exit Sub
        If CurrentEntryIndex <= 0 Then W.SetFocus Else EntryByIndex(CurrentEntryIndex - 1).Widget.SetFocus
      End If
    End If
  End If
End Sub
 
Private Sub W_Resize()
  AdjustEntries
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim i As Long, Pat As cCairoPattern, dyy As Single, TextShadowColor As Long
  'this If-Block ensures the fill of the TopArea, in case no Entry was selected (EntryIndex = -1)
  If mCurrentEntryIndex < 0 And EntriesCount > 0 Then
    CC.SetLineWidth 0
    CC.SetSourceColor W.BackColor, Alpha, 1.15
    dyy = W.ScaleHeight - (EntriesCount - mCurrentEntryIndex - 1) * mAccEntryHeight + 1
    CC.Rectangle -1, mCaptionHeight, dx + 2, dyy - mCaptionHeight, True
    CC.Fill
    
    CC.SetLineWidth 1, True
    CC.DrawLine 0, dyy, dx, dyy, True
    CC.SetSourceColor W.BorderColor, Alpha
    CC.Stroke
  End If
  
  'the Caption-Background
  Set Pat = Cairo.CreateLinearPattern(0, 1, 0, mCaptionHeight + 1)
    CC.SetLineWidth 0
    Pat.AddColorStop 0, W.BackColor, Alpha, 1.3
    Pat.AddColorStop 0.03, W.BackColor, Alpha, 0.97
    Pat.AddColorStop 1, W.BackColor, Alpha, 1.08
    CC.Rectangle 1, 1, dx - 2, mCaptionHeight, True
  CC.Fill , Pat
  CC.SetLineWidth 1, True
    CC.DrawLine 0, mCaptionHeight, dx, mCaptionHeight, True
    CC.SetSourceColor W.BorderColor, Alpha * 0.9
  CC.Stroke
  
  'a blurred TextShadow for the Caption
  CC.Save
    W.SelectFontSettingsInto CC
    TextShadowColor = W.ShadeColor(W.SelectionColor, 0.7)
    CC.TranslateDrawings 1.4, 1.5
    CC.TranslateDrawings 0.5, 0.5
    CC.DrawText 1, -2, dx - 2, mCaptionHeight + 4, mCaption, False, vbLeftJustify, 3, True, dtNormal, Alpha, True

    CC.SetLineJoin CAIRO_LINE_JOIN_ROUND
    CC.SetLineCap CAIRO_LINE_CAP_ROUND
    For i = 0 To 1 'this loop re-renders always on the same path (ensured by the True-Param in the .Stroke)
      CC.SetLineWidth 5 - 3 * i  'but with decreasing linewidth
      CC.SetSourceColor TextShadowColor, 0.11 + i * 0.05    'and slightly increasing Opacity
      CC.Stroke True
    Next i
    CC.SetLineWidth 1
    CC.SetSourceColor TextShadowColor, 0.2
    CC.Stroke
  CC.Restore
 
  'now the Caption on top of its Shadow
  W.SelectFontSettingsInto CC, W.ShadeColor(W.FocusColor, 0.95)
  CC.DrawText 1, -2, dx - 2, mCaptionHeight + 4, mCaption, False, vbLeftJustify, 3, True, dtNormal, Alpha
  
  'themed default-Border around our Widget
  CC.SetLineWidth 1, True
  Cairo.Theme.DrawTo CC, W, thmTypeBorder, 0, 0, 0, dx, dy
End Sub
  


