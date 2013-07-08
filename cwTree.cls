VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Event OwnerDrawItem(ByVal Index As Long, CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Event Click()
Event DblClick()
Event MouseUpClick()
Event MouseMoveOnListItem(ByVal HoverIndex As Long, ByVal RelX As Single, ByVal RelY As Single)

Public WithEvents VList As cwVList 'we inherit visually from a preimplemented Widget
Attribute VList.VB_VarHelpID = -1
Public DoubleClickExpandsNodes As Boolean

Private mDataSourceKey As String, WithEvents mDS As cDataSource, WithEvents tmrDecouple As cTimer
Attribute mDataSourceKey.VB_VarHelpID = -1
Attribute mDS.VB_VarHelpID = -1
Attribute tmrDecouple.VB_VarHelpID = -1
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private mCaption As String
Private mLastHoverIndex As Long, mLastArrowSize As Single, mLastArrowXOffs As Single, mInArrowArea As Boolean

Private Sub Class_Initialize()
  Set VList = New cwVList 'the Widget-Instance, to inherit from
      VList.RowHeight = 21
      VList.ListCount = 0
      VList.AllowLeftRightKeys = False
  'the following is some kind of "visual inheritance", since we use the already existent 'W'
  'from the VList, instead of creating our own, new 'W'-instance per: Set W = Cairo.WidgetBase
  Set W = VList.Widget
  DoubleClickExpandsNodes = True 'set the default to Expand/Collapse also on DoubleClicks
End Sub
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

'*** Public Properties ****
Public Property Get DataSourceKey() As String
  DataSourceKey = mDataSourceKey
End Property
Public Property Let DataSourceKey(ByVal NewValue As String)
  mDataSourceKey = NewValue
  On Error Resume Next
    Set DataSource = Cairo.DataSources(mDataSourceKey)
  On Error GoTo 0
End Property

Public Property Get DataSource() As cDataSource
  Set DataSource = mDS
End Property
Public Property Set DataSource(DS As cDataSource)
  Set mDS = DS
      mDS.TreeRefresh

  VList.ListIndex = -1 'the DS is at BOF after that (no selection)
  VList.ListCount = mDS.TreeVisibleCount
End Property

Public Property Get ListCount() As Long
  ListCount = VList.ListCount
End Property
Public Property Let ListCount(ByVal NewValue As Long)
  VList.ListCount = NewValue
End Property

Public Property Get ListIndex() As Long
  ListIndex = VList.ListIndex
End Property
Public Property Let ListIndex(ByVal NewValue As Long)
  VList.ListIndex = NewValue
End Property

Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
Dim Btn As cwButton
  If mCaption = NewValue Then Exit Property
  mCaption = NewValue
  If Len(mCaption) Then
    If VList.HeaderHeight = 0 Then VList.HeaderHeight = 23
    If Not Widgets.Exists("btnExpand") Then
      MakeCaptionButton "btnExpand", 3, "6", "Expand all", vbGreen
      MakeCaptionButton "btnCollapse", 20, "5", "Collapse all", vbCyan
    End If
  Else
    VList.HeaderHeight = 0
    Widgets.Remove "btnExpand"
    Widgets.Remove "btnCollapse"
  End If
  W.Refresh
End Property

Private Sub MakeCaptionButton(Key As String, ByVal xOffs As Long, Caption As String, ToolTip As String, BackColor As Long)
Dim Btn As cwButton
  Set Btn = Widgets.Add(New cwButton, Key, xOffs, (VList.HeaderHeight - 17) \ 2, 17, 17)
      Btn.Widget.CanGetFocus = False
      Btn.Widget.ToolTip = ToolTip
      Btn.Widget.BackColor = BackColor
      Btn.Widget.BorderColor = W.ShadeColor(W.BorderColor, 4)
      Btn.Widget.FocusColor = W.ShadeColor(W.BorderColor, 3)
      Btn.Widget.HoverColor = W.ShadeColor(Btn.Widget.BackColor, 1.5)
      Btn.Widget.FontName = "WebDings"
      Btn.Caption = Caption
      Btn.BorderRadius = 8
      Btn.DownStateCaptionOffset = 0
End Sub

Public Sub DrawArrow(CC As cCairoContext, ByVal Index As Long, ByVal xOffs As Single, ByVal ArrowSize As Single, ByVal Expanded As Boolean, Optional ByVal Expandable As Boolean = True)
Dim State As enmThemeDrawingState
  If mLastHoverIndex = Index Then
    mLastArrowSize = ArrowSize
    mLastArrowXOffs = xOffs
    If mInArrowArea Then State = thmStateHovered
  End If
  If Expandable Then
    Cairo.Theme.DrawTo CC, W, thmTypeArrow, State, xOffs, (VList.RowHeight - ArrowSize) \ 2 + 1, ArrowSize, ArrowSize, 0, _
                       IIf(Expanded, thmDirectionSE, thmDirectionRight)
  End If
End Sub

'------------  all the different Event-Handlers  ------------------------
Private Sub mDS_TreeStateChanged(ByVal Node As vbRichClient5.cCollection, ByVal Reason As TreeChangeReason)
  Set tmrDecouple = New_c.Timer(100, True, "TreeStateChanged")
End Sub

Private Sub mDS_Move(ByVal NewRowIdxZeroBased As Long)
  If VList.ListIndex = NewRowIdxZeroBased Then Exit Sub
  VList.ListIndex = NewRowIdxZeroBased
End Sub

Private Sub VList_Click()
Dim DoRaise As Boolean
  If Not mDS Is Nothing Then
    If VList.ListIndex = -1 And mDS.AbsolutePosition > 0 Then
      mDS.MoveFirst
      mDS.MovePrevious 'ensure BOF-Position
      DoRaise = True
    ElseIf VList.ListIndex <> -1 And mDS.AbsolutePosition <> VList.ListIndex + 1 Then
      mDS.AbsolutePosition = VList.ListIndex + 1
      DoRaise = True
    End If
'    If mDS.Count > 0 Then
'      If VList.ListIndex = -1 And mDS.AbsolutePosition <> VList.ListIndex Then
'        mDS.MoveFirst
'        mDS.MovePrevious
'        DoRaise = True
'      ElseIf mDS.AbsolutePosition <> VList.ListIndex + 1 Then
'        mDS.AbsolutePosition = VList.ListIndex + 1
'        DoRaise = True
'      End If
'    End If
  End If
  If DoRaise Then
    RaiseEvent Click
    W.RaiseBubblingEvent Me, "Click"
  End If
End Sub
 
Private Sub VList_DblClick()
  If mInArrowArea Then Exit Sub
  
  RaiseEvent DblClick
  W.RaiseBubblingEvent Me, "DblClick"
  If DoubleClickExpandsNodes Then ExpandOrCollapse VList.ListIndex
End Sub

Private Sub VList_MouseMoveOnListItem(ByVal HoverIndex As Long, ByVal RelX As Single, ByVal RelY As Single)
Dim InArrowArea As Boolean
  mLastHoverIndex = HoverIndex
  InArrowArea = (RelX > mLastArrowXOffs - 4 And RelX < mLastArrowXOffs + mLastArrowSize + 4)
  If InArrowArea <> mInArrowArea Then
    mInArrowArea = InArrowArea
    W.Refresh
  Else
    If Not VList.ShowHoverBar Then W.Refresh
  End If
  RaiseEvent MouseMoveOnListItem(HoverIndex, RelX, RelY)
End Sub

Private Sub VList_MouseUpClick()
  RaiseEvent MouseUpClick
  W.RaiseBubblingEvent Me, "MouseUpClick"
End Sub

Private Sub VList_OwnerDrawHeader(CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
  If Len(mCaption) = 0 Then Exit Sub
  Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, 0, -1, -2, dx + 3, dy + 3, 1, thmDirectionDown
  Cairo.Theme.DrawTo CC, W, thmTypeShine, 0, 0, -1, dx, Int((dy + 2) \ 2) - 1, 0, thmDirectionDown
  
  CC.DrawLine 0, dy, dx, dy, True, 1, W.BorderColor, 0.3
  W.SelectFontSettingsInto CC
  CC.DrawText 40, 1, dx - 40, dy, mCaption, True, vbLeftJustify, 2, True
End Sub

Private Sub VList_OwnerDrawItem(ByVal Index As Long, CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
  RaiseEvent OwnerDrawItem(Index, CC, dx, dy, Alpha)
End Sub

Private Sub W_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
Dim W As cWidgetBase: Set W = Sender.Widget
  If mDS Is Nothing Then Exit Sub
  If W.Key = "btnExpand" And EventName = "Click" Then
    mDS.TreeNodeExpand mDS.Col, True
    VList.ListIndex = -1
  ElseIf W.Key = "btnCollapse" And EventName = "Click" Then
    mDS.TreeNodeCollapse mDS.Col, True
    VList.ListIndex = -1
  End If
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If mInArrowArea Then ExpandOrCollapse VList.HoverIndex
End Sub

Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Key, Value, Parent As cCollection
  If VList.ListIndex < 0 Then Exit Sub
  If Not mDS.TreeElementInfoByVisibleIndex(VList.ListIndex, Key, Value, , , Parent) Then Exit Sub
  If TypeOf Value Is cCollection Then 'we have a Node
    Select Case KeyCode
      Case vbKeyRight
        If mDS.TreeNodeIsExpanded(Value) Then
          If VList.ListIndex < mDS.Count - 1 Then VList.ListIndex = VList.ListIndex + 1
        Else
          mDS.TreeNodeExpand Value
          If Not mDS.TreeNodeIsExpanded(Value) Then 'that's when the expansion has failed, or was reset by the client-App in an Event
            If VList.ListIndex < mDS.Count - 1 Then VList.ListIndex = VList.ListIndex + 1
          End If
        End If
      Case vbKeyLeft
        If mDS.TreeNodeIsExpanded(Value) Then
          mDS.TreeNodeCollapse Value
        Else
          If VList.ListIndex > 0 Then VList.ListIndex = VList.ListIndex - 1
        End If
      Case vbKeyBack:  VList.ListIndex = mDS.TreeNodeGetVisibleIndex(Parent)
    End Select
  Else 'we have a Value-Item
    Select Case KeyCode
      Case vbKeyRight: If VList.ListIndex < mDS.Count - 1 Then VList.ListIndex = VList.ListIndex + 1
      Case vbKeyLeft:  If VList.ListIndex > 0 Then VList.ListIndex = VList.ListIndex - 1
      Case vbKeyBack:  VList.ListIndex = mDS.TreeNodeGetVisibleIndex(Parent)
    End Select
  End If
End Sub

Private Sub tmrDecouple_Timer()
Dim Node As cCollection
  Select Case tmrDecouple.Tag
    Case "TreeStateChanged": VList.ListCount = mDS.Count 'setting the VList.Count to the DataSource.Count ensures a Tree-Refresh
  End Select
  Set tmrDecouple = Nothing
End Sub

Private Sub ExpandOrCollapse(ByVal Index As Long)
Dim Key, Value
  If Not mDS.TreeElementInfoByVisibleIndex(Index, Key, Value) Then Exit Sub
  If Not TypeOf Value Is cCollection Then Exit Sub
'  If Value.Count = 0 Then Exit Sub
  If mDS.TreeNodeIsExpanded(Value) Then mDS.TreeNodeCollapse Value Else mDS.TreeNodeExpand Value
End Sub

