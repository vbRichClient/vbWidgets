VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwVList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ColumnSortState
  SortDesc = -1
  SortNone = 0
  SortAsc = 1
End Enum

Event OwnerDrawRowSelector(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Event OwnerDrawHeader(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Event OwnerDrawItem(ByVal Index As Long, CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Event Click()
Event HeaderClick(ByVal ColIdx As Long, State As ColumnSortState)
Event DblClick()
Event DeleteKeyPressed()
Event ScrollIndexChange()
Event HScrollChange()
Event MouseUpClick()
Event MouseMoveOnListItem(ByVal HoverIndex As Long, ByVal RelX As Single, ByVal RelY As Single)

Private Alpha As Single, dx As Single, dy As Single
Private mListCount As Long, mListIndex As Long, mHoverIndex As Long
Private mRowHeight As Long, mVisibleRows As Long, mScrollerSize As Single
Private mHeaderHeight As Long, mRowSelectorWidth As Long
Private mSelectionInProgress As Boolean, mMultiSelect As MultiSelectConstants
Private mSelX As Single, mSelY As Single, mSelFirstState As Boolean, mSelFirstIndex As Long
Private mSelBits() As Byte, mSelCtrl As Boolean
Private mAllowDrag As Boolean, mShowHoverBar As Boolean, mShowSelection As Boolean

Private mColSortStates() As ColumnSortState, mCurSortCol As Integer

Private mColDefaultWidth%, mColWidthStart%, mColWidths%(), mColMap%()
Private mCurCol%, mColumnIndex%, mColDeltaStart%, mColClickIdx% ', mPreviousListIndex&
Private mAllowColResize As Boolean, mAllowRowResize As Boolean, SelectMode As Boolean

Public WithEvents HScrollBar As cwHScrollBar
Attribute HScrollBar.VB_VarHelpID = -1
Public WithEvents VScrollBar As cwVScrollBar
Attribute VScrollBar.VB_VarHelpID = -1
Public AllowLeftRightKeys As Boolean

Private WithEvents TScroll As cTimer
Attribute TScroll.VB_VarHelpID = -1
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
 
Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
      W.BackColor = vbWhite
      W.SetClientAreaOffsets 1, 1, 1, 1
      W.ImplementsWheelMessages = True
      W.RuntimePropertiesCommaSeparated = "VisibleRows,ListCount,ScrollIndex,ListIndex,HoverIndex,Selected,ColumnIndex,ColumnCount,ColumnWidth,ColMapIndex,DrawWidth,DrawHeight"
  AllowLeftRightKeys = True
  mShowHoverBar = True
  mShowSelection = True
  mRowHeight = 17
  mListIndex = -1
  mHoverIndex = -1
  ReDim mColWidths(-1 To 0)
  ReDim mColMap(-1 To 0)
  mColDefaultWidth = 120
  mColWidths(0) = mColDefaultWidth
  
  Set VScrollBar = Widgets.Add(New cwVScrollBar, "VScrollBar", , , , , False)
      VScrollBar.Widget.InnerWidget = True
  Set HScrollBar = Widgets.Add(New cwHScrollBar, "HScrollBar", , , , , False)
      HScrollBar.Widget.InnerWidget = True
  Set TScroll = New_c.Timer
End Sub

'Default-Interface Members (according to our cwWidget-Convention)
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

'******** Public Interface Implementation ***********
Public Property Get VisibleRows() As Long
  VisibleRows = mVisibleRows
End Property

Public Property Get ListCount() As Long
  ListCount = mListCount
End Property
Public Property Let ListCount(ByVal NewValue As Long)
  If NewValue < 0 Then Exit Property
  
  If NewValue < mListCount Then mHoverIndex = -1
  mListCount = NewValue
  If mListIndex > mListCount - 1 Then mListIndex = mListCount - 1
  W_Resize
End Property

Public Property Get ScrollerSize() As Long
  ScrollerSize = mScrollerSize
End Property

Public Property Get ScrollIndex() As Long
  ScrollIndex = VScrollBar.Value
End Property
Public Property Let ScrollIndex(ByVal NewValue As Long)
  If Not VScrollBar.Widget.Visible Then Exit Property
  If NewValue < 0 Then NewValue = 0
  If NewValue > VScrollBar.Max Then NewValue = VScrollBar.Max
  VScrollBar.Value = NewValue
End Property

Public Property Get ListIndex() As Long
  ListIndex = mListIndex
End Property
Public Property Let ListIndex(ByVal NewValue As Long)
  If NewValue < -1 Then Exit Property
  If NewValue > mListCount - 1 Then NewValue = mListCount - 1
  If mListIndex = NewValue Then Exit Property
  mListIndex = NewValue
  If mListIndex = -1 Then W.Refresh Else EnsureVisibleSelection
  
  RaiseEvent Click
  W.RaiseBubblingEvent Me, "Click"
End Property

Public Property Get HoverIndex() As Long
  HoverIndex = mHoverIndex
End Property

Public Property Get MultiSelect() As MultiSelectConstants
  MultiSelect = mMultiSelect
End Property
Public Property Let MultiSelect(ByVal NewValue As MultiSelectConstants)
  mMultiSelect = NewValue
  If mMultiSelect > vbMultiSelectNone Then ReDim mSelBits(0 To mListCount \ 8)
End Property

Public Sub EnsureVisibleSelection()
Dim NewPos&
  If mListIndex < 0 Then Exit Sub
  If mListIndex > mListCount Then Exit Sub
  If mListIndex < VScrollBar.Value Then
    VScrollBar.Value = mListIndex
  End If
  If mListIndex > VScrollBar.Value + mVisibleRows - 1 Then
    NewPos = mListIndex - mVisibleRows + 1
    If NewPos < 0 Then NewPos = 0
    If NewPos > VScrollBar.Max Then NewPos = VScrollBar.Max
    VScrollBar.Value = NewPos
  End If
  W.Refresh
End Sub

Public Function GetListIndexFromMouseY(ByVal y As Long) As Long
Dim CurYvirt As Long, TotalYvirt As Long
  CurYvirt = (-2 + y + VScrollBar.Value * mRowHeight)
  TotalYvirt = (-2 + mListCount * mRowHeight)
  
  If CurYvirt > TotalYvirt And Not mSelectionInProgress Then
    GetListIndexFromMouseY = -1
  Else
    GetListIndexFromMouseY = CurYvirt \ mRowHeight
  End If
End Function
 
Public Function GetColumnIndexFromMouseX(ByVal x As Single) As Long
Dim i&, xx&
  xx = mRowSelectorWidth
  GetColumnIndexFromMouseX = -1
  For i = HScrollBar.Value To UBound(mColWidths)
    xx = xx + ColumnWidth(i)
    If x < xx And x > xx - ColumnWidth(i) Then
      GetColumnIndexFromMouseX = i
      Exit For
    ElseIf x < mRowSelectorWidth - 1 Then
      GetColumnIndexFromMouseX = -1
      Exit For
    End If
  Next i
End Function

Public Property Get AllowDrag() As Boolean
  AllowDrag = mAllowDrag
End Property
Public Property Let AllowDrag(ByVal NewValue As Boolean)
  mAllowDrag = NewValue
End Property

Public Property Get AllowColResize() As Boolean
  AllowColResize = mAllowColResize
End Property
Public Property Let AllowColResize(ByVal NewVal As Boolean)
  mAllowColResize = NewVal
End Property

Public Property Get AllowRowResize() As Boolean
  AllowRowResize = mAllowRowResize
End Property
Public Property Let AllowRowResize(ByVal NewVal As Boolean)
  mAllowRowResize = NewVal
End Property

Public Property Get ShowHoverBar() As Boolean
  ShowHoverBar = mShowHoverBar
End Property
Public Property Let ShowHoverBar(ByVal NewValue As Boolean)
  mShowHoverBar = NewValue
End Property

Public Property Get ShowSelection() As Boolean
  ShowSelection = mShowSelection
End Property
Public Property Let ShowSelection(ByVal NewValue As Boolean)
  If mShowSelection = NewValue Then Exit Property
  mShowSelection = NewValue
  W.Refresh
End Property

Public Property Get Selected(ByVal Index As Long) As Boolean
Dim IdxTop&, IdxBottom&
  If Index < 0 Or Index >= mListCount Then Exit Property
  
  If mHoverIndex > mSelFirstIndex Then
    IdxTop = mSelFirstIndex: IdxBottom = mHoverIndex
  Else
    IdxTop = mHoverIndex: IdxBottom = mSelFirstIndex
  End If
  
  If mSelCtrl And mSelectionInProgress And Index >= IdxTop And Index <= IdxBottom Then
    Selected = Not mSelFirstState
  Else
    If mMultiSelect = vbMultiSelectNone Then
      Selected = (Index = mListIndex)
    Else
      If UBound(mSelBits) < Index \ 8 Then ReDim mSelBits(0 To Index \ 8)
      Selected = mSelBits(Index \ 8) And (2 ^ (Index Mod 8))
    End If
  End If
End Property
Public Property Let Selected(ByVal Index As Long, ByVal NewValue As Boolean)
  If Index < 0 Or Index >= mListCount Then Exit Property
  If mMultiSelect = vbMultiSelectNone Then
    If NewValue Then ListIndex = Index Else ListIndex = -1
  Else
    If UBound(mSelBits) < Index \ 8 Then ReDim mSelBits(0 To Index \ 8)

    If NewValue Then
      mSelBits(Index \ 8) = mSelBits(Index \ 8) Or (2 ^ (Index Mod 8))
    Else
      mSelBits(Index \ 8) = mSelBits(Index \ 8) And Not (2 ^ (Index Mod 8))
    End If
  End If
End Property

Public Sub ClearSelections(Optional ByVal NewListIndex As Long = -1)
  If mMultiSelect > vbMultiSelectNone Then ReDim mSelBits(0 To mListCount \ 8)
  ListIndex = NewListIndex
End Sub

Public Property Get RowHeight() As Long
  RowHeight = mRowHeight
End Property
Public Property Let RowHeight(ByVal NewValue As Long)
  If NewValue < 1 Then NewValue = 1
  If NewValue > 300 Then NewValue = 300
  mRowHeight = NewValue
  W_Resize
End Property
 
Public Property Get HeaderHeight() As Long
  HeaderHeight = mHeaderHeight
End Property
Public Property Let HeaderHeight(ByVal NewValue As Long)
  If NewValue < 0 Then NewValue = 0
  If NewValue > 100 Then NewValue = 100
  mHeaderHeight = NewValue
  W_Resize
End Property
  
Public Property Get RowSelectorWidth() As Long
  RowSelectorWidth = mRowSelectorWidth
End Property
Public Property Let RowSelectorWidth(ByVal NewValue As Long)
  If NewValue < 0 Then NewValue = 0
  If NewValue > 100 Then NewValue = 100
  mRowSelectorWidth = NewValue
  W_Resize
End Property

Public Property Get ColumnIndex() As Long
  ColumnIndex = mColumnIndex
End Property
Public Property Let ColumnIndex(ByVal NewValue As Long)
  If NewValue < -1 Then Exit Property
  If NewValue > ColumnCount - 1 Then NewValue = ColumnCount - 1
  If mColumnIndex = NewValue Then Exit Property
  mColumnIndex = NewValue
  
  If mColumnIndex = -1 Then W.Refresh Else EnsureVisibleSelection
End Property

Public Property Get ColumnCount() As Long
  ColumnCount = UBound(mColMap) + 1
End Property
Public Property Let ColumnCount(ByVal NewValue As Long)
Dim i&, OldUB&
  
  If NewValue < 1 Then NewValue = 1
  If NewValue > 32768 Then NewValue = 32768
'  If NewValue < 1 Or NewValue > 32768 Then HScrollBar.Value = 0: mListIndex = -1: Exit Property

  OldUB = UBound(mColWidths)
  mCurSortCol = -1
  ReDim Preserve mColSortStates(-1 To NewValue - 1)
  ReDim Preserve mColWidths(-1 To NewValue - 1)
  ReDim Preserve mColMap(-1 To NewValue - 1)
 
  For i = OldUB + 1 To UBound(mColWidths)
    mColWidths(i) = mColDefaultWidth
    mColMap(i) = i
  Next i
  
  AdjustHScrollBar
  AdjustDimensions
End Property

Public Property Get ColumnDefaultWidth() As Integer
  ColumnDefaultWidth = mColDefaultWidth
End Property
Public Property Let ColumnDefaultWidth(ByVal NewValue As Integer)
  mColDefaultWidth = NewValue
End Property

Public Property Get ColumnWidth(ByVal Idx As Long) As Integer
  If Idx < 0 Or Idx > UBound(mColWidths) Then Exit Property
  If UBound(mColWidths) = 0 Then
    ColumnWidth = DrawWidth
  Else
    ColumnWidth = mColWidths(mColMap(Idx))
  End If
End Property
Public Property Let ColumnWidth(ByVal Idx As Long, ByVal NewValue As Integer)
  If Idx < 0 Or Idx > UBound(mColWidths) Then Exit Property
  If NewValue < 0 Or NewValue > 4096 Then Exit Property
  mColWidths(mColMap(Idx)) = NewValue
  AdjustHScrollBar
End Property

Public Sub ResetSortStates()
  ReDim mColSortStates(-1 To ColumnCount - 1)
  mCurSortCol = -1
  W.Refresh
End Sub

Public Property Get ColMapIndex(ByVal Idx As Long) As Long
  If Idx < 0 Or Idx > UBound(mColMap) Then Exit Property
  ColMapIndex = mColMap(Idx)
End Property

Public Function VisibleCols() As Long
Dim i&, cx&, dx&
  dx = DrawWidth
  For i = HScrollBar.Value To UBound(mColWidths)
    cx = cx + ColumnWidth(i)
    VisibleCols = VisibleCols + 1
    If cx > dx Then Exit Function
  Next i
End Function

Public Sub MoveColumnToNewIndex(ByVal CurColIdx As Long, ByVal NewColIdx As Long)
Dim i As Long, UB As Long, IdxValue As Integer
  UB = UBound(mColMap)
  If CurColIdx < 0 Or CurColIdx > UB Then Exit Sub
  If NewColIdx < 0 Then NewColIdx = 0 Else If NewColIdx > UB Then NewColIdx = UB
  If CurColIdx = NewColIdx Then Exit Sub
  
  IdxValue = mColMap(CurColIdx) 'buffer the current Value at the current Index
  
  For i = CurColIdx To UB - 1: mColMap(i) = mColMap(i + 1): Next 'remove slot from List
  For i = UB - 1 To NewColIdx Step -1: mColMap(i + 1) = mColMap(i): Next 're-add free slot to List
  
  mColMap(NewColIdx) = IdxValue 'write the buffered value back
End Sub

Public Property Get DrawWidth() As Single
  DrawWidth = W.ScaleWidth - mRowSelectorWidth
  If VScrollBar.Widget.Visible Then DrawWidth = DrawWidth - ScrollerSize
End Property

Public Property Get DrawHeight() As Single
  DrawHeight = W.ScaleHeight - mHeaderHeight
  If HScrollBar.Widget.Visible Then DrawHeight = DrawHeight - ScrollerSize
End Property

Private Sub HScrollBar_Change()
  RaiseEvent HScrollChange
  W.Refresh
End Sub
Private Sub HScrollBar_Scroll()
  HScrollBar_Change
End Sub

'********* Event-Implementation of our hosted Classes (cwVScrollbar and cTimer)
Private Sub VScrollBar_Change()
  RaiseEvent ScrollIndexChange
  W.RaiseBubblingEvent Me, "ScrollIndexChange"
  W.Refresh
End Sub
Private Sub VScrollBar_Scroll()
  W.Refresh
End Sub
Private Sub VScrollBar_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
  ScrollIndex = ScrollIndex + LineOffset
End Sub

Private Sub TScroll_Timer()
Dim DoRefresh As Boolean
  If VScrollBar.Widget.Visible Then
    If TScroll.Tag = "Up" Then
      VScrollBar.Widget.LockRefresh = True
        If VScrollBar.Value > 0 Then VScrollBar.Value = VScrollBar.Value - 1
        mHoverIndex = VScrollBar.Value
        mListIndex = mHoverIndex
      VScrollBar.Widget.LockRefresh = False
      DoRefresh = True
    ElseIf TScroll.Tag = "Down" Then
      VScrollBar.Widget.LockRefresh = True
        If VScrollBar.Value < VScrollBar.Max Then VScrollBar.Value = VScrollBar.Value + 1
        mHoverIndex = VScrollBar.Value + mVisibleRows - 1
        If mHoverIndex > mListCount - 1 Then mHoverIndex = mListCount - 1
        mListIndex = mHoverIndex
      VScrollBar.Widget.LockRefresh = False
      DoRefresh = True
    End If
  End If
  
  If HScrollBar.Widget.Visible Then
    If TScroll.Tag = "Left" Then
      HScrollBar.Widget.LockRefresh = True
        If HScrollBar.Value > 0 Then HScrollBar.Value = HScrollBar.Value - 1
      HScrollBar.Widget.LockRefresh = False
      DoRefresh = True
    ElseIf TScroll.Tag = "Right" Then
      HScrollBar.Widget.LockRefresh = True
        If HScrollBar.Value < HScrollBar.Max Then HScrollBar.Value = HScrollBar.Value + 1
      HScrollBar.Widget.LockRefresh = False
      DoRefresh = True
    End If
  End If
  If DoRefresh Then W.Refresh
End Sub
 
Private Sub W_AddedToHierarchy()
  mSelectionInProgress = False
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As cWidgetBase)
  W.Refresh
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As cWidgetBase)
  mHoverIndex = -1
  W.Refresh
End Sub

'*************** cWidgetBase-Event-Implementation ************
Private Sub W_Resize()
  AdjustHScrollBar
  AdjustDimensions
  If Not W.IsHidden Then W.Refresh
End Sub
Public Sub AdjustDimensions()
Dim Rows As Long, Zoom As Single
  If dx = 0 Then dx = W.Width
  If dy = 0 Then dy = W.Height
  Zoom = W.Zoom
  mScrollerSize = Int(Cairo.Theme.GetScrollerSize * Zoom) / Zoom
   
  mVisibleRows = (dy - 3 - mHeaderHeight - IIf(HScrollBar.Widget.Visible, mScrollerSize, 0)) \ mRowHeight
  If mVisibleRows = 0 And mListCount > 0 Then mVisibleRows = 1
  If mVisibleRows > 1 Then VScrollBar.LargeChange = mVisibleRows - 1
  If mVisibleRows - 1 > mListCount Then Rows = mListCount Else Rows = mVisibleRows - 1
  If mListCount - Rows + 1 >= 0 Then VScrollBar.Max = mListCount - Rows - 1
  If VScrollBar.Max > 0 Then
    VScrollBar.Widget.Move dx - mScrollerSize - 2, 0, mScrollerSize, dy - 2 - IIf(HScrollBar.Widget.Visible, mScrollerSize, 0)
    VScrollBar.Widget.Visible = True
  Else 'make invisible
    VScrollBar.Widget.Visible = False
  End If
  
  HScrollBar.BottomRightEdgeWidth = IIf(VScrollBar.Widget.Visible, mScrollerSize, 0)
  HScrollBar.Widget.Move 0, dy - mScrollerSize - 2, dx - 2, mScrollerSize
  
'  If mListIndex < 0 Then VScrollBar.Value = 0
End Sub
  
Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
Dim OldSelIdx&, i&
  If mSelectionInProgress Then Exit Sub
  If Not AllowLeftRightKeys And (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight) Then Exit Sub
  If Shift = 0 Then mHoverIndex = -1
  OldSelIdx = mListIndex
  Select Case KeyCode
    Case vbKeyA
      If Shift And vbCtrlMask <> 0 And mMultiSelect > vbMultiSelectNone Then
        For i = 0 To mListCount - 1
          mListIndex = 0: Selected(i) = True
        Next i
        W.Refresh
        Exit Sub
      End If
    Case vbKeyLeft, vbKeyUp: ChangeListIndex mListIndex - 1
    Case vbKeyRight, vbKeyDown: ChangeListIndex mListIndex + 1
    Case vbKeyHome: ChangeListIndex 0
    Case vbKeyEnd: ChangeListIndex mListCount - 1
    Case vbKeyPageUp: ChangeListIndex mListIndex - IIf(VScrollBar.LargeChange < mListCount, VScrollBar.LargeChange, mListCount)
    Case vbKeyPageDown: ChangeListIndex mListIndex + IIf(VScrollBar.LargeChange < mListCount, VScrollBar.LargeChange, mListCount)
  End Select
  If OldSelIdx <> mListIndex Then
    RaiseEvent Click
    W.RaiseBubblingEvent Me, "Click"
  End If
End Sub
Private Sub ChangeListIndex(ByVal NewValue As Long) 'small Helper for the Event above
  If mListCount > 0 And NewValue < 0 Then NewValue = 0
  If NewValue < -1 Then NewValue = -1
  If NewValue > mListCount - 1 Then NewValue = mListCount - 1
  If mListIndex = NewValue Then Exit Sub
  mListIndex = NewValue
  If mListIndex = -1 Then W.Refresh Else EnsureVisibleSelection
End Sub

Private Sub W_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    RaiseEvent DeleteKeyPressed
    W.RaiseBubblingEvent Me, "DeleteKeyPressed"
  End If
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim NewSel As Long
  x = x - mRowSelectorWidth
  y = y - mHeaderHeight
  
  If y < 0 Then CalcColIdx x, True: Exit Sub
  
  mSelX = 0: mSelY = 0
  mSelCtrl = False
  NewSel = GetListIndexFromMouseY(y)
  
  If NewSel >= mListCount Then Exit Sub Else mSelectionInProgress = True
  mSelFirstIndex = NewSel
  If mAllowDrag Then mSelX = x: mSelY = y
  
  If mMultiSelect > vbMultiSelectNone Then
    mHoverIndex = NewSel
    mSelFirstState = Selected(NewSel)
  End If
  
  Select Case MultiSelect
    Case vbMultiSelectSimple
      If mSelFirstState = False Or mAllowDrag = False Then 'proceed as usual
        Selected(NewSel) = Not mSelFirstState
      End If
      W.Refresh
 
    Case vbMultiSelectExtended
      If Shift And vbCtrlMask Then
        mSelCtrl = True
        If mAllowDrag Then mSelX = 0: mSelY = 0
        W.Refresh
      ElseIf Shift And vbShiftMask Then
        mSelCtrl = True
        mSelFirstIndex = mListIndex
        mSelFirstState = False
        If mAllowDrag Then mSelX = 0: mSelY = 0
        W.Refresh
      ElseIf Not (mAllowDrag And mSelFirstState) Then
        mSelX = 0: mSelY = 0
        ReDim mSelBits(0 To mListCount \ 8)
        mSelCtrl = True
        mSelFirstState = False
      End If
  End Select
  
  If NewSel > -1 And NewSel < mListCount Then
    mHoverIndex = NewSel
    mListIndex = NewSel
    If y >= mRowHeight * mVisibleRows - 1 Then
      If mAllowDrag And mMultiSelect <> vbMultiSelectExtended Then mSelectionInProgress = False
      ListIndex = mHoverIndex
      EnsureVisibleSelection
    Else
      W.Refresh
    End If
  End If
End Sub

Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim NewSel&, i&, DoRefresh As Boolean
  x = x - mRowSelectorWidth
  y = y - mHeaderHeight
  
  NewSel = GetListIndexFromMouseY(y)
  
  If mSelX > 0 And mSelY > 0 Then 'check for dragging
    If Abs(x - mSelX) > 4 Or Abs(y - mSelY) > 4 Then
      mSelX = 0: mSelY = 0
      
      If mSelFirstIndex > -1 And mSelFirstIndex < mListCount Then mListIndex = mSelFirstIndex
      mHoverIndex = -1
      W.Refresh
      W.StartDrag
      mSelectionInProgress = False
    End If
  End If
  
  TScroll.Tag = ""
  
  If Not mSelectionInProgress Then
    If NewSel >= -1 And NewSel < mListCount Then
      If mHoverIndex <> NewSel Then
        If x < dx - 1 Then
          mHoverIndex = NewSel
        End If
        If (mMultiSelect = vbMultiSelectNone) And mShowHoverBar Then DoRefresh = True
      End If
    End If
  Else
    If NewSel > -1 And NewSel < mListCount Then
      If mHoverIndex <> NewSel Then
        mHoverIndex = NewSel: mListIndex = NewSel
        DoRefresh = True
      End If
    End If

    If y <= 1 Then
      i = 300 + y * 10: If i < 10 Then i = 10
      TScroll.Tag = "Up": TScroll.Interval = i: If Not TScroll.Enabled Then TScroll.Enabled = True
    ElseIf y - mRowHeight * mVisibleRows > 0 Then
      i = 300 - (y - mRowHeight * mVisibleRows) * 10: If i < 10 Then i = 10
      TScroll.Tag = "Down": TScroll.Interval = i: If Not TScroll.Enabled Then TScroll.Enabled = True
    End If
    If HScrollBar.Widget.Visible Then
      If x <= 1 Then
        i = 300 + x * 10: If i < 10 Then i = 10
        TScroll.Tag = "Left": TScroll.Interval = i: If Not TScroll.Enabled Then TScroll.Enabled = True
      ElseIf x - DrawWidth > 0 Then
        i = 300 - (x - DrawWidth) * 10: If i < 10 Then i = 10
        TScroll.Tag = "Right": TScroll.Interval = i: If Not TScroll.Enabled Then TScroll.Enabled = True
      End If
    End If
  End If
  RaiseEvent MouseMoveOnListItem(mHoverIndex, x, y - (mHoverIndex - VScrollBar.Value) * mRowHeight - 1)

  If DoRefresh And Len(TScroll.Tag) = 0 Then W.Refresh
End Sub

Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim i&, NewSel&
  x = x - mRowSelectorWidth
  y = y - mHeaderHeight
  
  If y < 0 And mListIndex < 0 Then TScroll.Enabled = False: Exit Sub
  
  NewSel = GetListIndexFromMouseY(y)
  
  If mSelX > 0 And mSelY > 0 And NewSel = mSelFirstIndex And mSelFirstState Then
    If MultiSelect = vbMultiSelectSimple Then
      Selected(mSelFirstIndex) = False
    ElseIf MultiSelect = vbMultiSelectExtended And Not mSelCtrl Then
      ReDim mSelBits(0 To mListCount \ 8)
      Selected(mSelFirstIndex) = True
    End If
  End If
  
  If mSelCtrl Then
    If mSelFirstIndex <= NewSel Then
      For i = mSelFirstIndex To NewSel
        If i >= 0 And i < mListCount Then Selected(i) = Not mSelFirstState
      Next i
    Else
      For i = NewSel To mSelFirstIndex
        If i >= 0 And i < mListCount Then Selected(i) = Not mSelFirstState
      Next i
    End If
  End If
  mSelX = 0: mSelY = 0
  TScroll.Enabled = False
  If mSelectionInProgress Then
    mSelectionInProgress = False
    
    If NewSel < 0 And mListCount > 0 Then
      ChangeListIndex 0
    Else
      If y < mRowHeight * mVisibleRows Then
        ChangeListIndex NewSel
      End If
    End If
  End If
  
  RaiseEvent Click
  W.RaiseBubblingEvent Me, "Click"
  RaiseEvent MouseUpClick
  W.RaiseBubblingEvent Me, "MouseUpClick"
End Sub

Private Sub W_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
 
  If MouseKeys And &H1000 Then  'horicontal Wheel-messages (WM_MOUSEHWHEEL-indicator)
    If HScrollBar.Value + LineOffset < 0 Or HScrollBar.Value + LineOffset > HScrollBar.Max Or HScrollBar.Widget.Visible = False Then Exit Sub
    HScrollBar.Widget.LockRefresh = True
      HScrollBar.Value = HScrollBar.Value + LineOffset
    HScrollBar.Widget.LockRefresh = False
    W.Refresh
 
  Else 'normal, vertical MouseWheel-Message
    If ScrollIndex + LineOffset < 0 Or (ScrollIndex + LineOffset) > VScrollBar.Max Then Exit Sub
    mHoverIndex = mHoverIndex + LineOffset
    ScrollIndex = ScrollIndex + LineOffset
  End If
End Sub

Private Sub W_DblClick()
  RaiseEvent DblClick
  W.RaiseBubblingEvent Me, "DblClick"
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  AdjustDimensions
  Alpha = W.AlphaInherited
  
  If Not W.IsHidden Then Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim i As Long, j As Long, ArrowSize As Long, Width As Single, Pat As cCairoPattern

  If W.BackColor <> -1 Then 'draw a BackGround first
    CC.SetLineWidth 0
    CC.Rectangle 0, 0, dx, dy, True
      CC.SetSourceColor W.BackColor, Alpha
    CC.Fill
  End If
 
  Width = dx - 4 - IIf(VScrollBar.Widget.Visible, mScrollerSize, 0)
  
  CC.Save
 
    If mHeaderHeight Then
      CC.TranslateDrawings 1, 1
      CC.ClipExplicit 0, 0, Width + 2, mHeaderHeight
        RaiseEvent OwnerDrawHeader(CC, Width + 2, mHeaderHeight - 1, Alpha)
        W.RaiseBubblingEvent Me, "OwnerDrawHeader", CC, Width + 2, mHeaderHeight - 1, Alpha
      CC.ResetClip
      CC.TranslateDrawings 1, mHeaderHeight
    Else
      CC.TranslateDrawings 2, 1
    End If
    
    If mRowSelectorWidth Then
      CC.ClipExplicit 0, 0, mRowSelectorWidth - 1, dy - mHeaderHeight - 1
        RaiseEvent OwnerDrawRowSelector(CC, mRowSelectorWidth - 1, dy - mHeaderHeight - 1, Alpha)
        W.RaiseBubblingEvent Me, "OwnerDrawRowSelector", CC, mRowSelectorWidth - 1, dy - mHeaderHeight - 1, Alpha
      CC.ResetClip
      CC.TranslateDrawings mRowSelectorWidth, 0
      Width = Width - RowSelectorWidth
    End If
 
    W.SelectFontSettingsInto CC
    For i = 0 To mVisibleRows
      j = i + ScrollIndex

      If j < mListCount Then
        If mRowSelectorWidth > 0 And Selected(j) And j = mListIndex Then
          If mRowSelectorWidth > mRowHeight Then ArrowSize = mRowHeight * 0.6 Else ArrowSize = mRowSelectorWidth * 0.6
          Cairo.Theme.DrawTo CC, W, thmTypeArrow, 0, -(mRowSelectorWidth + ArrowSize) \ 2 + 1, (mRowHeight - ArrowSize) \ 2 + 1, ArrowSize, ArrowSize, 0, thmDirectionRight
        End If
        
        CC.ClipExplicit 0, 0, Width, mRowHeight + 1
        
          CC.SetLineWidth 1, True
          
          If mShowHoverBar Then
            If MultiSelect = vbMultiSelectNone And j = HoverIndex And W.MouseOver Then
              Set Pat = Cairo.CreateLinearPattern(0, 0, 0, mRowHeight)
                  Pat.AddColorStop 0, W.HoverColor, Alpha * 0.3, 1.1
                  Pat.AddColorStop 1, W.HoverColor, Alpha * 0.3, 1
              CC.RoundedRect 0, 1, Width, mRowHeight - 1, 3, True
                CC.Fill True, Pat
                CC.SetSourceColor W.HoverColor, Alpha * 0.8, 1.07
              CC.Stroke
            End If
          End If
          
          If Selected(j) And mShowSelection Then
            Cairo.Theme.DrawTo CC, W, thmTypeListSelection, 0, 0, 0, Width, mRowHeight + 1, 3
          End If
          
          If mMultiSelect <> vbMultiSelectNone And j = mListIndex And W.Focused Then
            Cairo.Theme.DrawTo CC, W, thmTypeDottedRectangle, 0, 1, 1, Width - 2, mRowHeight - 1, 2
          End If
           
          RaiseEvent OwnerDrawItem(j, CC, Width, mRowHeight + 1, Alpha)
          W.RaiseBubblingEvent Me, "OwnerDrawItem", j, CC, Width, mRowHeight + 1, Alpha
        CC.ResetClip
        CC.TranslateDrawings 0, mRowHeight
      End If
    Next i
  CC.Restore
  
  CC.SetLineWidth 1, True
  Cairo.Theme.DrawTo CC, W, thmTypeBorder, 0, 0, 0, dx, dy
End Sub
  
Private Sub CalcColIdx(ByVal x As Single, Optional ByVal HeaderClick As Boolean)
Dim i&, xx&
  xx = mRowSelectorWidth
  mColClickIdx = -1
  For i = HScrollBar.Value To UBound(mColWidths)
    xx = xx + ColumnWidth(i)
    If HeaderClick Then
      If (x >= (xx - 2) And x <= (xx + 2)) And AllowColResize Then
        mColDeltaStart = x
        mColWidthStart = ColumnWidth(i)
        mColClickIdx = i
        Exit For
      End If
    End If
    If x < xx And x > xx - ColumnWidth(i) Then
      mColClickIdx = i + 1
      Exit For
    ElseIf x < mRowSelectorWidth - 1 Then
      mColClickIdx = -2
      Exit For
    End If
  Next i
End Sub

Private Sub AdjustHScrollBar()
Dim i As Long, xx As Long, DW As Long
  For i = 0 To UBound(mColWidths)
    xx = xx + ColumnWidth(i)
  Next i
  
  DW = DrawWidth
  If xx < DW Or UBound(mColWidths) = 0 Then
    HScrollBar.Value = 0
    HScrollBar.Widget.Visible = False
  Else
    xx = 0
    For i = UBound(mColWidths) To 0 Step -1
      xx = xx + ColumnWidth(i)
      If xx > DW Then
        HScrollBar.Max = IIf(i + 1 < UBound(mColWidths), i + 1, UBound(mColWidths))
        Exit For
      End If
    Next i
    HScrollBar.Widget.Visible = True
  End If
End Sub

Private Sub HandleHeaderMouseMoveStuff(ByVal x As Single, ByVal y As Single)
Dim yy As Long, dx As Long, dy As Long
  dx = W.ScaleWidth
  dy = W.ScaleHeight

  If (x > 0 And x < mRowSelectorWidth) And mColDeltaStart = 0 Then
    If mHeaderHeight Then
      yy = ((y + 2 - mHeaderHeight) \ mRowHeight) * mRowHeight
      If yy < 0 Then yy = 0
    Else
      yy = ((y - mHeaderHeight) \ mRowHeight) * mRowHeight + mRowHeight
    End If
    If ((y - mHeaderHeight + 1) >= yy And (y - mHeaderHeight - 1) <= yy) And (y - mHeaderHeight) < dx And mAllowRowResize Then
      W.MousePointer = IDC_SIZEWE
    Else
      If W.MouseOver Then W.MousePointer = IDC_ARROW
    End If
  ElseIf mColDeltaStart > 0 Then
    If mColWidthStart + (x - mColDeltaStart) > (dx - mScrollerSize - 1) Then
      ColumnWidth(mColClickIdx) = dx - ScrollerSize - 1
    ElseIf mColWidthStart + (x - mColDeltaStart) < 13 Then
      ColumnWidth(mColClickIdx) = 13
    Else
      ColumnWidth(mColClickIdx) = mColWidthStart + (x - mColDeltaStart)
    End If
    AdjustDimensions
    W.Refresh
  End If
End Sub

Private Sub HandleHeaderMouseUpStuff(ByVal x As Single, ByVal y As Single)
Dim PreviousColumnIdx%, OldColIdx%, LastState As ColumnSortState
  If mColClickIdx <> 0 Then
    If SelectMode Then
      mColClickIdx = 0
      PreviousColumnIdx = mColumnIndex
      mColumnIndex = GetColumnIndexFromMouseX(x)
      
    Else 'header-click
      If mColDeltaStart = 0 And (mColClickIdx > 0 Or mColClickIdx = -2) Then
        If mColClickIdx = -2 Then mCurCol = -1 Else mCurCol = mColClickIdx - 1
        OldColIdx = mColClickIdx
        CalcColIdx x
        If OldColIdx = mColClickIdx And y >= 0 And y <= mHeaderHeight Then
          LastState = mColSortStates(mCurCol)
          RaiseEvent HeaderClick(mCurCol, mColSortStates(mCurCol))
          If LastState <> mColSortStates(mCurCol) Then
            mCurSortCol = mCurCol
            W.Refresh
          End If
        End If
        
      End If
      mColClickIdx = 0
      If mColDeltaStart > 0 Then mColDeltaStart = 0: AdjustDimensions
      Exit Sub
    End If
  End If
  If mColDeltaStart > 0 Then mColDeltaStart = 0: AdjustDimensions ': Exit Sub
End Sub
