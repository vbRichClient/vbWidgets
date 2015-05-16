VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Click(Sender As cwToolBarItem)
Event ArrowClick(Sender As cwToolBarItem)

Public Enum DropDownArrowType
  ddNone = 0
  ddCrumbBar = 1
  ddDropDown = 2
End Enum

Public AllowGripperDrawing As Boolean

Private dx As Single, dy As Single, Alpha As Single
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  AllowGripperDrawing = True
  Set W = Cairo.WidgetBase
      W.ForwardFocus = True
'      W.CanGetFocus = False
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Public Function AddItem(Key As String, Optional ImageKey As String, Optional Caption As String, Optional ByVal ArrowType As DropDownArrowType, Optional ToolTip As String, Optional ByVal Enabled As Boolean = True, Optional ByVal IsCheckable As Boolean) As cwToolBarItem
Dim NewItem As cwToolBarItem, WLastItem As cWidgetBase
Dim xOffs As Single, Width As Single, TxtWidth As Double
  Set NewItem = New cwToolBarItem
      NewItem.Widget.ToolTip = ToolTip
      NewItem.Widget.ImageKey = ImageKey
      NewItem.Widget.ToolTipDelay = W.ToolTipDelay
      NewItem.Widget.FontName = W.FontName
      NewItem.Widget.FontSize = W.FontSize
      NewItem.Widget.Enabled = Enabled
      NewItem.Caption = Caption
      NewItem.ArrowType = ArrowType
      NewItem.IsCheckable = IsCheckable
      
  xOffs = 9
  If Widgets.Count Then
    Set WLastItem = Widgets(Widgets.Count).Widget
    xOffs = WLastItem.Left + WLastItem.Width
  End If
  If Len(ImageKey) Then Width = dy - 5
  If Len(Caption) Then
    If Caption = "-" Then
      Width = 6
    Else
      With W.MeasureContext
        Width = Width + .GetTextExtents(Caption) + 10
      End With
    End If
  End If
  If ArrowType Then Width = Width + 15
  Widgets.Add NewItem, Key, 1 + xOffs, 3, Width, dy - 5
  Set AddItem = NewItem
End Function

Private Sub W_AddedToHierarchy()
  dx = W.Width
  dy = W.Height
  W_ContainerResize
End Sub

Private Sub W_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
  If Sender Is Nothing Then Exit Sub
  If TypeOf Sender Is cwToolBarItem And Left(EventName, 2) <> "W_" Then 'reflect the two normal Events of a cwToolBarItem
    If EventName = "Click" Then RaiseEvent Click(Sender)
    If EventName = "ArrowClick" Then RaiseEvent ArrowClick(Sender)
  End If
End Sub

Private Sub W_ContainerResize()
  W.Move 0, 0, W.Parent.ScaleWidth, W.Height
  W.Refresh
End Sub
 
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned
  dy = dy_Aligned
  Alpha = W.AlphaInherited
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim Pat As cCairoPattern
  Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
      Pat.AddColorStop 0, W.BackColor, Alpha, 1.11
      Pat.AddColorStop 1, W.BackColor, Alpha, 0.93
  CC.Paint , Pat
  CC.DrawLine 0, dy - 1, dx, dy - 1, True, 1, vbBlack, 0.3 * Alpha
  
  If AllowGripperDrawing Then
    W.Alpha = 0.85 * Alpha
      Cairo.Theme.DrawTo CC, W, thmTypeSeparatorLine, 0, 5, 3, 1, dy - 4
      Cairo.Theme.DrawTo CC, W, thmTypeSeparatorLine, 0, 7, 3, 1, dy - 4
    W.Alpha = Alpha
  End If
End Sub

