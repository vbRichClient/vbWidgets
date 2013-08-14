VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Click()
Event Validate(ByVal NewRowIdx As Long, ByVal NewColIdx As Long, ByVal OldRowIdx As Long, ByVal OldColIdx As Long, ByVal EditBox As cwTextBox, Cancel As Boolean)
Event AddNewAttempt(Cancel As Boolean)
Event DeleteAttempt(Cancel As Boolean)
Event HeaderClick(ByVal ColIdx As Long, State As ColumnSortState)
Event DataSourcePositionChanged()

Private WithEvents VList As cwVList 'we inherit visually from a preimplemented Widget
Attribute VList.VB_VarHelpID = -1
Private WithEvents W As cWidgetBase, WithEvents mDS As cDataSource, mDataSourceKey As String
Attribute W.VB_VarHelpID = -1
Attribute mDS.VB_VarHelpID = -1

Private mAllowEdit As Boolean, mAllowAddNew As Boolean, mAllowDelete As Boolean, mlngAlternateRowColor As Long

Private Sub Class_Initialize()
  Set VList = New cwVList
      VList.RowHeight = 20
      VList.ListCount = 100
      VList.HeaderHeight = 24
      VList.RowSelectorWidth = 24
      VList.RowHeight = 19
      VList.ShowHoverBar = False
 
  Set W = VList.Widget
      W.RuntimePropertiesCommaSeparated = "ColumnIndex,ColumnCount,ColumnWidth,RowIndex"
'      W.BackColor = -1
'      W.Alpha = 0.7
  mlngAlternateRowColor = -1
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

'*** Public Properties ****
Public Property Get HeaderHeight() As Long
  HeaderHeight = VList.HeaderHeight
End Property
Public Property Let HeaderHeight(ByVal NewValue As Long)
  VList.HeaderHeight = NewValue
End Property

Public Property Get MultiSelect() As MultiSelectConstants
  MultiSelect = VList.MultiSelect
End Property
Public Property Let MultiSelect(ByVal NewValue As MultiSelectConstants)
  VList.MultiSelect = NewValue
End Property

Public Property Get ColumnIndex() As Long
  ColumnIndex = VList.ColumnIndex
End Property
Public Property Let ColumnIndex(ByVal NewValue As Long)
  VList.ColumnIndex = NewValue
End Property

Public Property Get ColumnCount() As Long
  ColumnCount = VList.ColumnCount
End Property
Public Property Let ColumnCount(ByVal NewValue As Long)
  VList.ColumnCount = NewValue
End Property

Public Property Get ColumnDefaultWidth() As Integer
  ColumnDefaultWidth = VList.ColumnDefaultWidth
End Property
Public Property Let ColumnDefaultWidth(ByVal NewValue As Integer)
  VList.ColumnDefaultWidth = NewValue
End Property

Public Property Get ColumnWidth(ByVal Idx As Long) As Integer
  ColumnWidth = VList.ColumnWidth(Idx)
End Property
Public Property Let ColumnWidth(ByVal Idx As Long, ByVal NewValue As Integer)
  VList.ColumnWidth(Idx) = NewValue
End Property

Public Function VisibleCols() As Long
  VisibleCols = VList.VisibleCols
End Function

Public Sub MoveColumnToNewIndex(ByVal CurColIdx As Long, ByVal NewColIdx As Long)
  VList.MoveColumnToNewIndex CurColIdx, NewColIdx
End Sub

Public Property Get RowIndex() As Long
  RowIndex = VList.ListIndex
End Property

Public Property Get AllowColResize() As Boolean
  AllowColResize = VList.AllowColResize
End Property
Public Property Let AllowColResize(ByVal NewVal As Boolean)
  VList.AllowColResize = NewVal
End Property

Public Property Get AllowRowResize() As Boolean
  AllowRowResize = VList.AllowRowResize
End Property
Public Property Let AllowRowResize(ByVal NewVal As Boolean)
  VList.AllowRowResize = NewVal
End Property

Public Property Get AllowEdit() As Boolean
  AllowEdit = mAllowEdit
End Property
Public Property Let AllowEdit(ByVal NewVal As Boolean)
  mAllowEdit = NewVal
End Property

 Public Property Get AllowAddNew() As Boolean
  AllowAddNew = mAllowAddNew
End Property
Public Property Let AllowAddNew(ByVal NewVal As Boolean)
  If mAllowAddNew = NewVal Then Exit Property
  mAllowAddNew = NewVal
  W.Refresh
End Property

Public Property Get AllowDelete() As Boolean
  AllowDelete = mAllowDelete
End Property
Public Property Let AllowDelete(ByVal NewVal As Boolean)
  If mAllowDelete = NewVal Then Exit Property
  mAllowDelete = NewVal
End Property

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
Dim i&, Cols&, Rows&
  If Not DS Is Nothing Then
    Rows = DS.RecordCount
    Cols = DS.FieldCount
  End If
 
  W.LockRefresh = True
    Set mDS = DS

    ResetSortStates 'clear the last Column-Sort-infos
 
    VList.ListCount = Rows
    VList.ColumnCount = Cols
 
    VList.ListIndex = -1 'the Rs is at BOF after that (no selection)
  W.LockRefresh = False
End Property

Public Sub ResetSortStates()
  VList.ResetSortStates
End Sub

Private Sub mDS_Move(ByVal NewRowIdxZeroBased As Long)
  If VList.ListIndex = NewRowIdxZeroBased Then Exit Sub
  VList.ListIndex = NewRowIdxZeroBased
  RaiseEvent DataSourcePositionChanged
  W.RaiseBubblingEvent Me, "DataSourcePositionChanged"
End Sub

Private Sub mDS_NewDataContentArrived()
  VList.ListCount = mDS.RecordCount
End Sub

Private Sub VList_Click()
Dim DoRaise As Boolean
  If Not mDS Is Nothing Then
    If mDS.RecordCount > 0 Then
      If VList.ListIndex = -1 And mDS.AbsolutePosition > 0 Then
        mDS.MoveFirst
        mDS.MovePrevious 'ensure BOF-Position
        DoRaise = True
      ElseIf VList.ListIndex <> -1 And mDS.AbsolutePosition <> VList.ListIndex + 1 Then
        mDS.AbsolutePosition = VList.ListIndex + 1
        DoRaise = True
      End If
    End If
  End If
  If DoRaise Then
    RaiseEvent Click
    W.RaiseBubblingEvent Me, "Click"
    RaiseEvent DataSourcePositionChanged
    W.RaiseBubblingEvent Me, "DataSourcePositionChanged"
  End If
End Sub

Private Sub VList_HeaderClick(ByVal ColIdx As Long, State As ColumnSortState)
  RaiseEvent HeaderClick(ColIdx, State)
End Sub

Private Sub VList_OwnerDrawHeader(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Dim i As Long, StartIdx As Long, xx As Double, ww As Long, State As enmThemeDrawingState
  Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, -1, -2, dx + 3, dy + 3, 1, thmDirectionDown
  Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 0, -1, dx, Int((dy + 2) \ 2) - 1, 0, thmDirectionDown
  CC.DrawLine 0, dy, dx, dy, True, 1, W.BorderColor, 0.3
 
  If mDS Is Nothing Then Exit Sub
  
  If Not VList.HScrollBar Is Nothing Then
    If VList.HScrollBar.Widget.Visible Then StartIdx = VList.HScrollBar.Value
  End If
  
  xx = VList.RowSelectorWidth
  For i = StartIdx To StartIdx + VisibleCols - 1
    ww = VList.ColumnWidth(i)
    If ww > 0 Then
      If i = StartIdx + VList.VisibleCols - 1 Then ww = dx - xx
      Cairo.Theme.DrawTo CC, W, thmTypeSeparatorLine, 0, xx, 0, 1, dy
      CC.DrawText xx, 1, ww, dy, mDS.FieldName(VList.ColMapIndex(i)), True, , 4, True
      xx = xx + ww
    End If
  Next i
End Sub

Private Sub VList_OwnerDrawRowSelector(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Dim i As Long, yy As Double, State As enmThemeDrawingState
  State = thmStatePressed
  Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, -1, -1, dx + 0.5 * dx, dy + 2, 1, thmDirectionLeft
  Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 0, 0, dx + dx, dy, 1, thmDirectionLeft
  For i = 0 To VList.VisibleRows - 1
    yy = yy + VList.RowHeight
    Cairo.Theme.DrawTo CC, W, thmTypeSeparatorLine, 0, -dx, yy, dx + dx, 1
  Next i
End Sub

Private Sub VList_OwnerDrawItem(ByVal Index As Long, CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Dim i As Long, StartIdx As Long, V, S As String, xx As Double, ww As Long
  If mDS Is Nothing Then Exit Sub
 
  If VList.HScrollBar.Widget.Visible Then StartIdx = VList.HScrollBar.Value
 
  xx = 0
  For i = StartIdx To StartIdx + VList.VisibleCols - 1
    ww = VList.ColumnWidth(i)
    If ww > 0 Then
      If i = StartIdx + VList.VisibleCols - 1 Then ww = dx - xx
      V = mDS.ValueMatrix(Index, VList.ColMapIndex(i))
      Select Case VarType(V)
        Case vbByte Or vbArray: S = "<BLOB>"
        Case Else: S = V
      End Select
      If AlternateRowColor <> -1 Then
        'here we ensure a different Background-Line-Coloring for Odd-Indexes.
        If Index Mod 2 Then
          CC.Rectangle 0, 0, dx, dy, True
            CC.SetSourceColor AlternateRowColor
          CC.Fill
        End If
      End If
 
      CC.DrawText xx, 0, ww, dy, S, True, vbLeftJustify, 2, True
      xx = xx + ww
    End If
  Next i
End Sub
 
Public Property Get AlternateRowColor() As Long
  AlternateRowColor = mlngAlternateRowColor
End Property

Public Property Let AlternateRowColor(ByVal lngAlternateRowColor As Long)
  If lngAlternateRowColor = mlngAlternateRowColor Then Exit Property
  mlngAlternateRowColor = lngAlternateRowColor
  W.Refresh
End Property
