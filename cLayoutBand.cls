VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLayoutBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mParent As cWidgetForm, mCells As New Collection, mBandIndex As Long

Friend Sub Init(Parent As cWidgetForm, CellWidths(), ByVal BandIndex As Long)
Dim NewCell As cLayoutCell, i As Long
  Set mParent = Parent
  mBandIndex = BandIndex
  Set mCells = Nothing
  For i = 0 To UBound(CellWidths)
    Set NewCell = New cLayoutCell: NewCell.Init mParent, CellWidths(i), mBandIndex, mCells.Count
    mCells.Add NewCell
  Next
End Sub

Public Function CellCount() As Long
  CellCount = mCells.Count
End Function

Public Function Cell(ByVal IndexZeroBased As Long) As cLayoutCell
  Set Cell = mCells(IndexZeroBased + 1)
End Function

Public Sub ResizeWith(ByVal xOffs As Single, yOffs As Single, ByVal Width As Single)
Dim Cell As cLayoutCell, AbsWidth As Single, PercSum As Single, PercFac As Single, y As Single, yMax As Single
  
  For Each Cell In mCells
    If Cell.Width > 1 Then AbsWidth = AbsWidth + Cell.Width Else PercSum = PercSum + Cell.Width
  Next
  
  Width = Width - AbsWidth
  If Width > 0 And PercSum > 0 Then PercFac = Width / IIf(CellCount = 1, 1, PercSum) Else PercFac = 1
  For Each Cell In mCells
    y = yOffs
      Cell.ResizeWith xOffs, y, Cell.Width * IIf(Cell.Width > 1, 1, PercFac)
    If y > yMax Then yMax = y
  Next
  yOffs = yMax
End Sub
