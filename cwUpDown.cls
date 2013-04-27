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

Private mDataSource As String, WithEvents mDS As cDataSource
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

Public Property Get DataSource() As String
  DataSource = mDataSource
End Property
Public Property Let DataSource(ByVal NewValue As String)
  mDataSource = NewValue
  On Error Resume Next 'catch the error, if not existing
    Set mDS = Cairo.DataSources(mDataSource)
    VList.ListCount = mDS.Count
  If Err Then Err.Clear
  VList.ListIndex = 0
End Property

Public Property Get DataIndex() As Long
  DataIndex = VList.ScrollIndex
End Property

Public Property Get DataValue()
  DataValue = GetDataValue(DataIndex)
End Property

Private Sub mDS_Move(ByVal NewRowIdxZeroBased As Long)
  VList.ListIndex = NewRowIdxZeroBased
End Sub

Private Sub mDS_NewDataContentArrived()
  VList.ListCount = mDS.RecordCount
  VList.ListIndex = 0
End Sub

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

Private Function GetDataValue(Index As Long)
  If mDS Is Nothing Then GetDataValue = "UpDown": Exit Function
  If Index < 0 Or Index >= mDS.RecordCount Then Exit Function
  GetDataValue = mDS.ValueMatrix(Index, 0)
End Function
 
