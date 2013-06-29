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
 
Private WithEvents mDS As cDataSource, mDataField As String, mFieldIndex As Long
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

Public Property Get DataSource()
  Set DataSource = mDS
End Property
Public Property Set DataSource(DS)
  Set mDS = DS
  If mDS.DataSource Is Nothing Then VList.ListCount = 0: Exit Property
  VList.ListCount = mDS.RecordCount
End Property

Public Property Get DataField() As String
  DataField = mDataField
End Property
Public Property Let DataField(ByVal NewValue As String)
  If mDataField = NewValue Then Exit Property
  mDataField = NewValue
  If mDS Is Nothing Then Exit Property
  Dim i As Long
  For i = 0 To mDS.FieldCount - 1
    If StrComp(mDS.FieldName(i), mDataField) = 0 Then mFieldIndex = i: Exit For
  Next i
End Property

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
 
Private Sub mDS_Move(ByVal NewRowIdxZeroBased As Long)
  If VList.ListIndex = NewRowIdxZeroBased Then Exit Sub
  VList.ListIndex = NewRowIdxZeroBased
End Sub

Private Sub mDS_NewDataContentArrived()
  VList.ListCount = mDS.RecordCount
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
  DropDown.Caption = mDS.FieldValue(mFieldIndex)
End Sub
Private Sub VList_MouseUpClick()
  DropDown.DestroyPopup
End Sub
Private Sub VList_OwnerDrawItem(ByVal Index As Long, CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
  If mDS Is Nothing Then Exit Sub
  If Index >= mDS.RecordCount Then Exit Sub
  CC.DrawText 2, 0, dx - 2, dy, CStr(mDS.ValueMatrix(Index, mFieldIndex)), True, vbLeftJustify, 3, True
End Sub
 
Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
  If DropDown.Btn.Value Then Exit Sub 'early exit in case the DropDown is shown (to avoid double-KeyHandling here)
  Select Case KeyCode
    Case vbKeyLeft, vbKeyUp:    If ListIndex > 0 Then ListIndex = ListIndex - 1
    Case vbKeyRight, vbKeyDown: If ListIndex < ListCount - 1 Then ListIndex = ListIndex + 1
  End Select
End Sub
