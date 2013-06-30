VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event DataIndexChange()

'DataSource-Implementation-Vars
Private WithEvents mDS As cDataSource, mDataSourceKey As String, mDataField As String
Attribute mDS.VB_VarHelpID = -1

Private WithEvents W As cWidgetBase, WithEvents VList As cwVList
Attribute W.VB_VarHelpID = -1
Attribute VList.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set VList = New cwVList
  Set W = VList.Widget '<- this does not create a new WidgetBase-Instance, but "inherits" the Vlist-WidgetBase
      W.FocusColor = -1
      W.BorderColor = -1
      W.BackColor = -1
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
  VList.ListCount = 0
  Set mDS = DS
  If mDS Is Nothing Then Exit Property
  mDataSourceKey = mDS.Key
  VList.ListCount = mDS.RecordCount
  If mDS.RecordCount Then mDS.MoveFirst
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
  VList.ListIndex = IIf(VList.ListCount, 0, -1)
End Sub
'***************** End of typical-DataSource-related Procedures ***************

Public Property Get DataIndex() As Long
  DataIndex = VList.ScrollIndex
End Property
 
Public Property Get DataValue()
  DataValue = GetDataValue(DataIndex)
End Property

Private Function GetDataValue(Index As Long)
  If mDS Is Nothing Then GetDataValue = "UpDown": Exit Function
  If Index < 0 Or Index >= mDS.RecordCount Then Exit Function
  GetDataValue = mDS.ValueMatrix(Index, FieldIndex)
End Function

Private Sub W_Resize()
  VList.RowHeight = W.ScaleHeight - 1
End Sub
 
Private Sub VList_ScrollIndexChange()
  VList.ListIndex = VList.ScrollIndex
  If Not mDS Is Nothing Then mDS.AbsolutePosition = VList.ScrollIndex + 1
End Sub

Private Sub VList_OwnerDrawItem(ByVal Index As Long, CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
  W.SelectFontSettingsInto CC
  CC.DrawText 1, 1, dx, dy, CStr(GetDataValue(Index)), True, vbCenter, 0, True
End Sub


