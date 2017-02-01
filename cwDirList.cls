VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwDirList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Change()
Event Click()
Event DblClick()
Event MouseUpClick()

Public WithEvents Tree As cwTree 'we inherit visually from a preimplemented Widget
Attribute Tree.VB_VarHelpID = -1
 
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Public LanScanMode As ServerScanMode

Private mPath As String, WithEvents mDS As cDataSource
Attribute mDS.VB_VarHelpID = -1

Private Sub Class_Initialize()
  LanScanMode = NormalScan_FromSystemCache ' FastSubNetPings_HostNames ' the enum-Value Normal_WNetEnumResource is also possible, but usually takes (much) longer
  
  Set Tree = New cwTree 'the Widget-Instance, to inherit from
      Tree.DoubleClickExpandsNodes = False '<- since we use the DoubleClick here, to signalize a Path-Change (as in the classic VB-DirListBox)

  Set mDS = New_c.DataSource
      mDS.Init New_c.Collection(False, TextCompare, True), "DirTree", , True

  Path = App.Path '<- pre-init with the App.Path
       
  'the following is some kind of "visual inheritance", since we use the already existent 'W'
  'from the VList, instead of creating our own, new 'W' per: Set W = Cairo.WidgetBase
  Set W = Tree.Widget
End Sub
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property


''*** Public Widget-Properties ****

Public Property Get DataSource() As cDataSource
  Set DataSource = mDS
End Property

Public Property Get ListCount() As Long
  ListCount = Tree.ListCount
End Property
Public Property Let ListCount(ByVal NewValue As Long)
  Tree.ListCount = NewValue
End Property

Public Property Get ListIndex() As Long
  ListIndex = Tree.ListIndex
End Property
Public Property Let ListIndex(ByVal NewValue As Long)
  Tree.ListIndex = NewValue
End Property
 
Public Function ValidatePath(ByVal Path As String) As String
  If Left$(Path, 4) = "LAN\" Then
    ValidatePath = "\\" & Mid$(Path, 5) 'replace our "LAN"-Prefix with the UNC-Path-Prefix
  Else
    ValidatePath = Path
  End If
End Function
  
Public Sub ReScan()
  Path = mPath
End Sub
 
Public Property Get SelectedPath() As String
Dim Key, Value
  If mDS Is Nothing Then Exit Property
  If Not mDS.TreeElementInfoByVisibleIndex(ListIndex, Key, Value) Then Exit Property
 
  SelectedPath = ValidatePath(mDS.TreeNodeFullPath(Value))
End Sub

Public Property Get Path() As String
  Path = mPath & "\"
End Property
Public Property Let Path(ByVal NewValue As String)
Dim i As Long, Drive, Node As cCollection, IsUNCPath As Boolean, PathArr() As String
 
  If Len(NewValue) = 0 Then Exit Property
  If Not New_c.FSO.FolderExists(NewValue) Then Exit Property
  
  If mDS.Count = 0 Then
    mDS.TreeNodeAdd mDS.TreeRoot, "LAN"
 
    For Each Drive In New_c.FSO.GetDrives
      mDS.TreeRoot.Add New_c.Collection(False, TextCompare, True), Drive
    Next
  End If

  mPath = IIf(Right$(NewValue, 1) = "\", Left$(NewValue, Len(NewValue) - 1), NewValue) 'strip the trailing "\"

  IsUNCPath = (Left$(mPath, 2) = "\\")
  PathArr = Split(IIf(IsUNCPath, Mid$(mPath, 3), mPath), "\")
  If IsUNCPath Then
    Set Node = mDS.TreeRoot.ItemByIndex(0)
    If Not Node Is Nothing Then mDS.TreeNodeExpand Node
  Else
    Set Node = mDS.TreeRoot
  End If
 
  For i = 0 To UBound(PathArr)
    If Not Node.Exists(PathArr(i)) Then mDS.TreeNodeAdd Node, PathArr(i)
    Set Node = Node.Item(PathArr(i)) 'here we switch from ParentNode to ChidNode
    If i < UBound(PathArr) Then mDS.TreeNodeExpand Node '... and expand the new ChidNode (triggers mDS_TreeStateChanged, which ensures Filling in the gaps from the FileSystem)
  Next i
  
  Set Tree.DataSource = mDS
      Tree.ListIndex = mDS.TreeNodeGetVisibleIndex(Node)
'      Debug.Print "PathSet", Tree.ListIndex, mDS.AbsolutePosition, mDS.TreeNodeGetVisibleIndex(Node)
  RaiseEvent Change
  If Not W Is Nothing Then W.RaiseBubblingEvent Me, "Change"
End Property


'******************* Event-Sinks ******************

Private Sub mDS_TreeStateChanged(ByVal Node As vbRichClient5.cCollection, ByVal Reason As vbRichClient5.TreeChangeReason)
Dim i As Long, DirList As cDirList, Path As String, Share
'Static Shares As cSortedDictionary
  If Reason = NodeExpanded Then
 
    Path = ValidatePath(mDS.TreeNodeFullPath(Node))
 
    'check for LAN-stuff
    If Path = "LAN" Then
      ExpandLanNode
      Exit Sub
    End If
 
    If InStrRev(Path, "\") = 2 Then 'a child directly below the LAN-Node, so it's not (yet) an UNC-Path (only the ServerName or IP)
      Node.RemoveAll
'      If Shares Is Nothing Then Set Shares = New_c.SortedDictionary(TextCompare)
'      If Not Shares.Exists(Path) Then Shares.Add Path, New_c.SMBScan.GetShares(Path)
'      For Each Share In Shares(Path)
      For Each Share In New_c.SMBScan.GetShares(Path)
        mDS.TreeNodeAdd Node, Share, , True
      Next Share
      mDS.TreeRefresh
      Exit Sub
    End If
 
    'it's a valid Path, so we proceed here now
    On Error Resume Next
      Set DirList = New_c.FSO.GetDirList(Path)
    If Err Or DirList.SubDirsCount = 0 Then
      Err.Clear
      mDS.TreeNodeCollapse Node, , True
    Else
      Node.RemoveAll
      For i = 0 To DirList.SubDirsCount - 1
        mDS.TreeNodeAdd Node, DirList.SubDirName(i), , True
      Next i
      mDS.TreeRefresh
    End If
  End If
End Sub
 
Private Sub Tree_OwnerDrawItem(ByVal Index As Long, CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
'define your Base-Offsets for indentation in the consts below
Const BaseOffsX As Long = 17, Indent As Long = 16, ArrowSize As Long = 8, IconSize As Long = 16

'all the Vars defined below, will get filled in the call that follows (ByRef)
Dim Key, Value, Expanded As Boolean, Level As Long '<-
  If Not mDS.TreeElementInfoByVisibleIndex(Index, Key, Value, Expanded, Level) Then Exit Sub
  
'Ok ... drawing-time
Dim x As Long, Path As String, DirList As cDirList, Expandable As Boolean
  x = BaseOffsX + (Level + 1) * Indent
  
  If TypeOf Value Is cCollection Then 'we have a Node-Entry
 
    Path = ValidatePath(mDS.TreeNodeFullPath(Value)) 'for all deeper nested Levels, we call TreeNodeFullPath (and validate the result)
    
    If Left$(Path, 2) = "\\" Or Path = "LAN" Then
      Expandable = True
    End If
    On Error Resume Next
      Expandable = New_c.FSO.DirectoryHasSubDirs(Path, True)
    On Error GoTo 0

    Tree.DrawArrow CC, Index, x - IconSize - ArrowSize * 1.5, ArrowSize, Expanded, Expandable 'any Node gets a small Triangle drawn, using a Helper-Function in cwTree
    CC.RenderSurfaceContent New_c.FSO.GetFileIconSurface(Path, IIf(Level, True, False), IconSmall, IconNormal), x - IconSize - 1, (dy - IconSize) \ 2, IconSize, IconSize
 
  End If
 
  CC.DrawText x, 0, dx - x, dy, CStr(Key), True, vbLeftJustify, 3, True
End Sub

Private Sub Tree_DblClick()
  RaiseEvent DblClick
  W.RaiseBubblingEvent Me, "DblClick"
End Sub

Private Sub Tree_MouseUpClick()
  RaiseEvent MouseUpClick
  W.RaiseBubblingEvent Me, "MouseUpClick"
End Sub
 
Private Sub Tree_Click()
  RaiseEvent Click
  If Not W Is Nothing Then W.RaiseBubblingEvent Me, "Click"
End Sub
 
Private Sub ExpandLanNode()
Dim Node As cCollection, SMB As cSMBScan, Servers As Collection, Server
  Set Node = mDS.TreeRoot.ItemByIndex(0)

  Set SMB = New_c.SMBScan
  If Not W.Root Is Nothing Then W.Root.MousePointer = IDC_WAIT
    Set Servers = SMB.GetServers(LanScanMode)
  If Not W.Root Is Nothing Then W.Root.MousePointer = IDC_ARROW
  
  For Each Server In Servers
    If Not Node.Exists(Server) Then mDS.TreeNodeAdd Node, Server
  Next Server
  
  mDS.TreeNodeExpand Node
End Sub
