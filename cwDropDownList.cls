VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwDropDownList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Click()
 
'DataSource-Implementation-Vars
Private WithEvents mDS As cDataSource, mDataSourceKey As String, mDataField As String
Attribute mDS.VB_VarHelpID = -1

Public WithEvents DropDown As cwDropDown
Attribute DropDown.VB_VarHelpID = -1
Public WithEvents VList As cwVList
Attribute VList.VB_VarHelpID = -1

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set DropDown = New cwDropDown
  Set VList = New cwVList
  Set W = DropDown.Widget
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

'***************** Start of typical-DataSource-related Procedures ***************
Public Function SetDataSource(CollectionOrRecordset As Object, Key As String, Optional DataField As String) As cDataSource
If Len(DataField) Then mDataField = DataField
Set SetDataSource = New_c.DataSource
    SetDataSource.Init CollectionOrRecordset, Key, Cairo.DataSourceDispatcher
    Set DataSource = SetDataSource
End Function

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
  DropDown.Caption = ""
  VList.ListCount = 0
  Set mDS = DS
  If mDS Is Nothing Then Exit Property
  mDataSourceKey = mDS.Key
  VList.ListCount = mDS.RecordCount
  If mDS.RecordCount Then mDS.MoveFirst: mDS.MovePrevious
End Property
 
Public Property Get DataField() As String
  DataField = mDataField
End Property
Public Property Let DataField(ByVal NewValue As String)
  mDataField = NewValue
End Property
Private Property Get FieldIndex() As Long
Dim i As Long
  If mDS Is Nothing Then Exit Property
  For i = 0 To mDS.FieldCount - 1
    If StrComp(mDS.FieldName(i), mDataField, vbTextCompare) = 0 Then FieldIndex = i: Exit For
  Next i
End Property
 
Private Sub mDS_Move(ByVal NewRowIdxZeroBased As Long)
  If VList.ListIndex <> NewRowIdxZeroBased Then VList.ListIndex = NewRowIdxZeroBased
End Sub
Private Sub mDS_NewDataContentArrived()
  VList.ListCount = mDS.RecordCount
End Sub
'***************** End of typical-DataSource-related Procedures ***************

Public Property Get Text() As String
  Text = DropDown.Text
End Property

Public Property Get ListIndex() As Long
  ListIndex = VList.ListIndex
End Property
Public Property Let ListIndex(ByVal NewValue As Long)
  VList.ListIndex = NewValue
End Property

Public Property Get ListCount() As Long
  If mDS Is Nothing Then Exit Sub
  ListCount = mDS.Count
End Property

Private Sub DropDown_PopupWidgetRequest(WidgetInstance As Object, WidthPxl As Single, HeightPxl As Single)
  Set WidgetInstance = VList
  W.Tag = VList.ListIndex
  HeightPxl = 200 * W.Zoom
End Sub
Private Sub DropDown_PopupEscapeKeyPressed(DestroyPopup As Boolean)
  VList.ListIndex = W.Tag
End Sub
Private Sub DropDown_PopupReturnKeyPressed(DestroyPopup As Boolean)
  DropDown.DestroyPopup
End Sub

Private Sub VList_Click()
Dim DoRaise As Boolean
  If mDS Is Nothing Then Exit Sub

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
  If DoRaise Then
    RaiseEvent Click
    W.RaiseBubblingEvent Me, "Click"
  End If
  If mDS.AbsolutePosition > 0 Then DropDown.Caption = mDS.FieldValue(FieldIndex)
End Sub
Private Sub VList_MouseUpClick()
  DropDown.DestroyPopup
End Sub
Private Sub VList_OwnerDrawItem(ByVal Index As Long, CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
  If mDS Is Nothing Then Exit Sub
  If Index >= mDS.RecordCount Then Exit Sub
  CC.DrawText 2, 0, dx - 2, dy, CStr(mDS.ValueMatrix(Index, FieldIndex)), True, vbLeftJustify, 3, True
End Sub
 
Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
  If DropDown.Btn.Value Then Exit Sub 'early exit in case the DropDown is shown (to avoid double-KeyHandling here)
  Select Case KeyCode
    Case vbKeyLeft, vbKeyUp:    If ListIndex > 0 Then ListIndex = ListIndex - 1
    Case vbKeyRight, vbKeyDown: If ListIndex < ListCount - 1 Then ListIndex = ListIndex + 1
  End Select
End Sub
