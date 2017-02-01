VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLayoutCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public MarginLeft, MarginTop, MarginRight, MarginBottom, MarginBetweenWidgets
 
Private mParent As cWidgetForm, mWidth!, mBandIndex&, mColIndex&, mWidgets As New Collection, mGroup As cwFrame

Friend Sub Init(Parent As cWidgetForm, ByVal Width As Single, ByVal BandIndex As Long, ByVal ColIndex As Long)
  Set mParent = Parent
  mWidth = Width
  mBandIndex = BandIndex
  mColIndex = ColIndex
  MarginLeft = 5: MarginTop = 5: MarginRight = 5: MarginBottom = 5: MarginBetweenWidgets = 5
End Sub

Public Property Get Width() As Single
  Width = mWidth
End Property
 
Public Sub ResizeWith(xOffs As Single, yOffs As Single, ByVal Width As Single)
  Dim W As cWidgetBase, dx As Single, dy As Single, MaxLabelWidth As Single
  
  For Each W In mWidgets
    If TypeOf W.object Is cwLabeledTextBox Then
      If W.object.CaptionWidth > MaxLabelWidth Then MaxLabelWidth = W.object.CaptionWidth
    End If
  Next
  MaxLabelWidth = MaxLabelWidth + 5
  
  dx = Width - MarginLeft - MarginRight: If dx < 1 Then dx = 1
  yOffs = yOffs + MarginTop
  For Each W In mWidgets
    If TypeOf W.object Is cwGrid Then dy = W.Height Else dy = 24
    If TypeOf W.object Is cwLabeledTextBox Then W.object.LabelWidthPercent = MaxLabelWidth / dx
    If TypeOf W.object Is cwFrame Then dy = ResizeInGroup(W, MarginTop + 5, dx)
    If TypeOf W.object Is cwImage Then dy = W.object.CanvasHeight * dx / W.object.CanvasWidth
    W.Move xOffs + MarginLeft, yOffs, dx, dy
    yOffs = yOffs + dy + MarginBetweenWidgets
  Next
  yOffs = yOffs + MarginBottom
  xOffs = xOffs + Width
End Sub

Private Function ResizeInGroup(WG As cWidgetBase, ByVal yOffs As Single, ByVal Width As Single) As Single
Dim WO As Object, W As cWidgetBase, xOffs As Single, dx As Single, dy As Single, MaxLabelWidth As Single
  
  For Each WO In WG.Widgets
    If TypeOf WO Is cwLabeledTextBox Then
      If WO.CaptionWidth > MaxLabelWidth Then MaxLabelWidth = WO.CaptionWidth
    End If
  Next
  MaxLabelWidth = MaxLabelWidth + 5
  If WG.Tag = "V" Then
    For Each WO In WG.Widgets: Set W = WO.Widget
      If TypeOf WO Is cwGrid Then dy = W.Height Else dy = 24
      If TypeOf WO Is cwLabeledTextBox Then WO.LabelWidthPercent = MaxLabelWidth / (Width - MarginLeft - MarginRight)
      
      W.Move MarginLeft, yOffs, Width - MarginLeft - MarginRight, dy
      yOffs = yOffs + dy + MarginBetweenWidgets
    Next
  ElseIf WG.Widgets.Count Then
    dy = WG.Widgets(1).Widget.Height: If dy < 24 Then dy = 24
    dx = Int((Width - MarginLeft - MarginRight - (WG.Widgets.Count - 1) * MarginBetweenWidgets) / WG.Widgets.Count)
    xOffs = MarginLeft
    For Each WO In WG.Widgets: Set W = WO.Widget
      W.Move xOffs, yOffs, dx, dy: xOffs = xOffs + dx + MarginBetweenWidgets
    Next
    yOffs = yOffs + dy + MarginBetweenWidgets
  End If
  ResizeInGroup = yOffs + MarginBottom + 5
End Function

Public Function BeginGroupVertical(ByVal Caption As String) As cWidgetBase
  Set mGroup = mParent.Widgets.Add(New cwFrame, GetKey)
      mGroup.Caption = Caption
      mGroup.Widget.Tag = "V"
  mWidgets.Add mGroup.Widget
  Set BeginGroupVertical = mGroup.Widget
End Function
Public Function BeginGroupHorizontal(ByVal Caption As String) As cWidgetBase
  Set mGroup = mParent.Widgets.Add(New cwFrame, GetKey)
      mGroup.Caption = Caption
      mGroup.Widget.Tag = "H"
  mWidgets.Add mGroup.Widget
  Set BeginGroupHorizontal = mGroup.Widget
End Function
Public Sub EndGroup()
  Set mGroup = Nothing
End Sub

Public Function AddLabeledTextBox(ByVal Caption As String, Optional ByVal Alignment As AlignmentConstants = vbRightJustify) As cwLabeledTextBox
  Set AddLabeledTextBox = mParent.Widgets.Add(New cwLabeledTextBox, GetKey)
      AddLabeledTextBox.Caption = Caption
      AddLabeledTextBox.Alignment = Alignment
  mWidgets.Add AddLabeledTextBox.Widget
End Function

Public Function AddButton(ByVal Caption As String) As cwButton
  Set AddButton = mParent.Widgets.Add(New cwButton, GetKey)
      AddButton.Caption = Caption
  mWidgets.Add AddButton.Widget
End Function

Public Function AddCheckBox(ByVal Caption As String, Optional ByVal Value As Long) As cwButton
  Set AddCheckBox = mParent.Widgets.Add(New cwButton, GetKey)
      AddCheckBox.ButtonStyle = CheckBox
      AddCheckBox.Caption = Caption
      AddCheckBox.Value = Value
  mWidgets.Add AddCheckBox.Widget
End Function

Public Function AddOptionBox(ByVal Caption As String, Optional ByVal Value As Boolean) As cwButton
  Set AddOptionBox = IIf(mGroup Is Nothing, mParent, mGroup).Widgets.Add(New cwButton, GetKey)
      AddOptionBox.ButtonStyle = OptionBox
      AddOptionBox.Caption = Caption
      AddOptionBox.Value = Value
      AddOptionBox.OptionGroupKey = "VLayout_OptGrp_" & mBandIndex & "_" & mColIndex
  If mGroup Is Nothing Then mWidgets.Add AddOptionBox.Widget
End Function

Public Function AddDataGrid(ByVal Caption As String, ByVal DataSource As cDataSource, Optional ByVal DGHeightPxl As Single = 280) As cwGrid
  Set AddDataGrid = IIf(mGroup Is Nothing, mParent, mGroup).Widgets.Add(New cwGrid, GetKey, , , 1, DGHeightPxl)
  Set AddDataGrid.DataSource = DataSource
      AddDataGrid.Widget.Tag = Caption
  If DataSource.Col Is Nothing Then
    If DataSource.Rs.Fields(0).PrimaryKey = True Then AddDataGrid.ColumnWidth(0) = 0
  Else
    'AddDataGrid.ColumnWidth(1) = 0 'hide the Value-Column (only show Key)
  End If
  If DataSource.RecordCount Then DataSource.MoveFirst
  If mGroup Is Nothing Then mWidgets.Add AddDataGrid.Widget
End Function
 
Public Function AddImageBox(ByVal Caption As String, ByVal ImageKey As String, Optional ByVal Width& = 256, Optional ByVal Height& = 256) As cwImage
  Set AddImageBox = IIf(mGroup Is Nothing, mParent, mGroup).Widgets.Add(New cwImage, GetKey)
      AddImageBox.Widget.Tag = Caption
      AddImageBox.Widget.ImageKey = ImageKey
      AddImageBox.SetCanvasSize Width, Height
  If mGroup Is Nothing Then mWidgets.Add AddImageBox.Widget
End Function

Private Function GetKey() As String
  Static ID As Long
  GetKey = "Band " & mBandIndex & ", Col " & mColIndex & ", " & ID
  ID = ID + 1
End Function

