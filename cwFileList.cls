VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Click()
Event MouseUpClick()

Public WithEvents VList As cwVList 'we inherit visually from a preimplemented Widget
Attribute VList.VB_VarHelpID = -1

Private mDirList As cDirList
Private mPattern As String, mSortMode As DirListSortMode
Private mHidden As Boolean, mSystem As Boolean, mLastKeyCode As Integer, mLastIndex As Integer

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
 
Private Sub Class_Initialize()
  Set mDirList = New_c.FSO.GetDirList(App.Path, dlSortByNameLogically, mPattern, mHidden, mSystem)
  
  Set VList = New cwVList 'the Widget-Instance, to inherit from
      VList.RowHeight = 19
      VList.ListCount = mDirList.FilesCount
  
  'the following is some kind of "visual inheritance", since we use the already existent 'W'
  'from the VList, instead of creating our own, new 'W' per: Set W = Cairo.WidgetBase
  Set W = VList.Widget
End Sub
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

'*** Public Properties ****
Public Property Get DirList() As cDirList
  Set DirList = mDirList
End Property

Public Property Get Path() As String
  Path = mDirList.Path
End Property
Public Property Let Path(ByVal NewValue As String)
  mLastKeyCode = 0: mLastIndex = 0
  If Right$(NewValue, 1) <> "\" Then NewValue = NewValue & "\"
  Set mDirList = New_c.FSO.GetDirList(NewValue, mSortMode, mPattern, mHidden, mSystem)
  VList.ListCount = mDirList.FilesCount
  W.Refresh
End Property

Public Property Get Pattern() As String
  Pattern = mPattern
End Property
Public Property Let Pattern(ByVal NewValue As String)
  If mPattern = NewValue Then Exit Property Else mPattern = NewValue: ReScan
End Property

Public Property Get SortMode() As DirListSortMode
  SortMode = mSortMode
End Property
Public Property Let SortMode(ByVal NewValue As DirListSortMode)
  If mSortMode = NewValue Then Exit Property Else mSortMode = NewValue: ReScan
End Property

Public Property Get Hidden() As Boolean
  Hidden = mHidden
End Property
Public Property Let Hidden(ByVal NewValue As Boolean)
  If mHidden = NewValue Then Exit Property Else mHidden = NewValue: ReScan
End Property

Public Property Get System() As Boolean
  System = mSystem
End Property
Public Property Let System(ByVal NewValue As Boolean)
  If mSystem = NewValue Then Exit Property Else mSystem = NewValue: ReScan
End Property

Public Sub ReScan()
  Path = mDirList.Path
End Sub

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

Public Property Get FileName() As String
  If VList.ListIndex < 0 Then Exit Property
  FileName = mDirList.FileName(VList.ListIndex)
End Property

Public Property Get PathAndFileName() As String
  If VList.ListIndex < 0 Then Exit Property
  PathAndFileName = mDirList.Path & mDirList.FileName(VList.ListIndex)
End Property

Private Sub VList_Click()
  RaiseEvent Click
  W.RaiseBubblingEvent Me, "Click"
End Sub

Private Sub VList_MouseUpClick()
  RaiseEvent MouseUpClick
  W.RaiseBubblingEvent Me, "MouseUpClick"
End Sub

Private Sub VList_OwnerDrawItem(ByVal Index As Long, CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Dim Srf As cCairoSurface, FntSize As Double, FntOffs As Single, IcoOffs As Single
Const IcoSize As Single = 16
  If Index >= mDirList.FilesCount Then Exit Sub
  
  FntSize = CC.GetFontHeight
  FntOffs = Int((dy - FntSize) / 2)
  IcoOffs = Int((dy - IcoSize) / 2)

  Set Srf = New_c.FSO.GetFileIconSurface(mDirList.Path & mDirList.FileName(Index), False, IconSmall, IconNormal)
  CC.RenderSurfaceContent Srf, 1 + IcoOffs, IcoOffs, IcoSize, IcoSize, CAIRO_FILTER_GAUSSIAN
  CC.TextOut 1 + IcoOffs + dy, FntOffs, mDirList.FileName(Index)
End Sub
 
Private Sub W_KeyUp(KeyCode As Integer, Shift As Integer)
Dim Index As Long, WChar As String
    If KeyCode <= 40 Then Exit Sub
    Index = FindIndex(KeyCode, IIf(mLastKeyCode = KeyCode, mLastIndex + 1, 0), WChar)
    If mLastKeyCode = KeyCode And WChar <> ChrW$(KeyCode) Then Index = FindIndex(KeyCode, 0, WChar)
  
    If Index < mDirList.FilesCount Then
      mLastKeyCode = KeyCode
      mLastIndex = Index
      VList.ClearSelections Index
      W.Refresh
    End If
End Sub

Private Function FindIndex(KeyCode As Integer, StartIdx As Long, WChar As String) As Long
  For FindIndex = StartIdx To mDirList.FilesCount - 1
    WChar = UCase$(Left$(mDirList.FileName(FindIndex), 1))
    If WChar >= ChrW$(KeyCode) Then Exit For
  Next
End Function
