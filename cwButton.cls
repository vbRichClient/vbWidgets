VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CmdButtonStyle
  CommandButton
  ToolbarButton
  CheckBox
  CheckButton
  OptionBox
  OptionButton
End Enum

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event AccessKeyPress(KeyAscii As Integer)
Event OwnerDrawDownState(ByVal CC As cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, CancelBuiltIn As Boolean)
Event OwnerDrawUpState(ByVal CC As cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, CancelBuiltIn As Boolean)
Event OwnerDrawCheckState(ByVal CC As cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, CancelBuiltIn As Boolean)
Event OwnerDrawOptionState(ByVal CC As cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, CancelBuiltIn As Boolean)

Private mBorderRadius As Double

Private dx As Single, dy As Single, Alpha As Single
Private BDown As Boolean, OutSide As Boolean

Private mCaption As String, mButtonStyle As CmdButtonStyle, mValue
Private WithEvents tmrDefaultStateRefresh As cTimer
Attribute tmrDefaultStateRefresh.VB_VarHelpID = -1
Private mHAlignment As AlignmentConstants
 
'DataSource-Implementation-Vars
Private WithEvents mDS As cDataSource, mDataSourceKey As String, mDataField As String
Attribute mDS.VB_VarHelpID = -1

Public ImageSize As Long, DownStateCaptionOffset As Long

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
 
  mBorderRadius = 3.25
  mHAlignment = vbCenter
  mValue = 0
  DownStateCaptionOffset = 1
End Sub
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Public Property Get HAlignment() As AlignmentConstants
  HAlignment = mHAlignment
End Property
Public Property Let HAlignment(ByVal NewValue As AlignmentConstants)
  If mHAlignment = NewValue Then Exit Property
  mHAlignment = NewValue
  W.Refresh
End Property

Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
Dim AccKey$, Pos&
  If mCaption = NewValue Then Exit Property
  mCaption = NewValue
  W.Refresh
  Pos = InStr(Replace(mCaption, "&&", "--"), "&")
  If Pos Then AccKey = Mid$(Replace(mCaption, "&&", "--"), Pos + 1, 1)

  If Len(AccKey) Then W.AccessKeys = AccKey
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
  Set mDS = DS
  If mDS Is Nothing Then Exit Property
  mDataSourceKey = mDS.Key
  If mDS.AbsolutePosition > 0 Then Me.Value = mDS.FieldValue(FieldIndex)
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
 
Private Sub mDS_NewDataContentArrived()
  If mDS.AbsolutePosition > 0 Then Me.Value = mDS.FieldValue(FieldIndex)
End Sub

Private Sub mDS_Move(ByVal NewRowIdxZeroBased As Long)
  If mDS.AbsolutePosition > 0 Then Me.Value = mDS.FieldValue(FieldIndex)
End Sub

Private Sub mDS_FieldChange(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long)
  If StrComp(mDataField, mDS.FieldName(ColIdxZeroBased), vbTextCompare) Then Exit Sub  'this is not our FieldName
  If mDS.AbsolutePosition > 0 Then Me.Value = mDS.FieldValue(FieldIndex)
End Sub

Private Sub mDS_PropValueChange(Key As Variant, Value As Variant, OldValue As Variant)
  If StrComp(mDataField, Key, vbTextCompare) Then Exit Sub   'this is not our Key
  If mDS.AbsolutePosition > 0 Then Me.Value = mDS.FieldValue(FieldIndex)
End Sub
'***************** End of typical-DataSource-related Procedures ***************


Public Property Get ButtonStyle() As CmdButtonStyle
  ButtonStyle = mButtonStyle
End Property
Public Property Let ButtonStyle(ByVal NewValue As CmdButtonStyle)
  If mButtonStyle = NewValue Then Exit Property
  mButtonStyle = NewValue
  Value = 0
  W.Refresh
End Property

Public Property Get BorderRadius() As Double
  BorderRadius = mBorderRadius
End Property
Public Property Let BorderRadius(ByVal NewValue As Double)
  If NewValue < 0 Then NewValue = 0
  If NewValue = mBorderRadius Then Exit Property
  mBorderRadius = NewValue
  W.Refresh
End Property

Public Property Get Value()
  Value = mValue
End Property
Public Property Let Value(ByVal NewValue)
  'just some preformatting, to ensure correct Types in our Variant-Value-Property here
  Select Case mButtonStyle
    Case CmdButtonStyle.CommandButton: Exit Property
    
    Case CmdButtonStyle.CheckBox, CmdButtonStyle.CheckButton
      If IsEmpty(NewValue) Or IsNull(NewValue) Then NewValue = 2
      If NewValue = 2 Then mValue = 2: W.Refresh: Exit Property
      If CBool(NewValue) Then NewValue = 1 Else NewValue = 0

    Case CmdButtonStyle.OptionBox, CmdButtonStyle.OptionButton
      If IsEmpty(NewValue) Or IsNull(NewValue) Then NewValue = False
      If CBool(NewValue) Then NewValue = True Else NewValue = False
      If NewValue = False And mValue <> 0 Then mValue = False: W.Refresh: Exit Property
  End Select
  
  If mValue = NewValue Then Exit Property

  CheckAndRaiseClick

  W.Refresh
End Property


Private Sub tmrDefaultStateRefresh_Timer()
  If Not W.Default Then Set tmrDefaultStateRefresh = Nothing: Exit Sub
  If Not tmrDefaultStateRefresh.Tag = DefaultStateCondition Then
    tmrDefaultStateRefresh.Tag = DefaultStateCondition
    W.Refresh
  End If
End Sub

Private Sub W_AccessKeyPress(KeyAscii As Integer)
  If W.Root.IsHidden Then Exit Sub
  RaiseEvent AccessKeyPress(KeyAscii)
 
  If ((KeyAscii = vbKeyReturn And W.Default) Or (KeyAscii = vbKeyEscape And W.Cancel)) And mButtonStyle = CommandButton Then
    If Not (DefaultStateCondition Or W.Cancel) Then Exit Sub
    W.Refresh
    CheckAndRaiseClick
  ElseIf InStr(1, W.AccessKeys, Chr$(KeyAscii), vbTextCompare) Then
    W.SetFocus
    CheckAndRaiseClick
    W.Refresh
  End If
End Sub

Private Function DefaultStateCondition() As Boolean
  If W.Default And tmrDefaultStateRefresh Is Nothing Then
    Set tmrDefaultStateRefresh = New_c.Timer(50, True)
    tmrDefaultStateRefresh.Tag = DefaultStateCondition
  End If
  If W.Root Is Nothing Then Exit Function
  If W.Root.ActiveWidget Is Nothing Then
    DefaultStateCondition = W.Default
  ElseIf InStr(1, TypeName(W.Root.ActiveWidget.object), "Button", vbTextCompare) = 0 Then
    DefaultStateCondition = W.Default
  End If
End Function
 
Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
  If KeyCode = vbKeySpace Then BDown = True: W.Refresh
  If KeyCode = vbKeyReturn And Shift = 0 Then
    If mButtonStyle = CmdButtonStyle.CommandButton Then CheckAndRaiseClick
  End If
  If KeyCode = vbKeyEscape And Shift = 0 And W.Cancel Then
    If mButtonStyle = CmdButtonStyle.CommandButton Then CheckAndRaiseClick
  End If
End Sub

Private Sub W_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub W_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
  If KeyCode = 32 Then
    CheckAndRaiseClick
    BDown = False
    W.Refresh
  End If
End Sub
 
Private Sub W_MouseEnter(ByVal MouseLeaveWidget As cWidgetBase)
  W.Refresh
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As cWidgetBase)
  W.Refresh
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 1 Then BDown = True: W.Refresh
End Sub
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim NewOutsideValue As Boolean
 
  NewOutsideValue = (x < 0 Or x > dx Or y < 0 Or y > dy)
  If BDown And (NewOutsideValue <> OutSide) Then
    OutSide = NewOutsideValue
    W.Refresh
  Else
    OutSide = NewOutsideValue
  End If
End Sub
Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If BDown And Not OutSide Then CheckAndRaiseClick
  BDown = False: OutSide = False: W.Refresh
End Sub

Private Sub CheckAndRaiseClick()
  Select Case mButtonStyle
    Case CmdButtonStyle.CommandButton 'no Value-interpretation, just raise then Event
    
    Case CmdButtonStyle.CheckBox, CmdButtonStyle.CheckButton
      If mValue = 2 Or mValue = 1 Then mValue = 0 Else mValue = 1
      
    Case CmdButtonStyle.OptionBox, CmdButtonStyle.OptionButton
      If mValue = 0 Then
        UnCheckOtherOptions
        mValue = True
      Else
        Exit Sub
      End If
      
  End Select
  
  If Not mDS Is Nothing And Len(mDataField) Then mDS.PropValue(mDataField) = mValue

  RaiseEvent Click
  W.RaiseBubblingEvent Me, "Click", mValue
End Sub

Private Sub UnCheckOtherOptions()
Dim V, Button As cwButton
  If W.Parent Is Nothing Then Exit Sub
  If W.Parent.ChildCount = 0 Then Exit Sub
  For Each V In W.Parent.Widgets
    If TypeOf V Is cwButton Then
      Set Button = V
      If Button.ButtonStyle >= OptionBox Then
        If Not Button Is Me And Button.Value Then
          Button.Value = False
        End If
      End If
    End If
  Next
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  Alpha = W.AlphaInherited
  dx = dx_Aligned
  dy = dy_Aligned
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim CancelBuiltIn As Boolean, Alpha As Single
  Alpha = Alpha
  
  W.SelectFontSettingsInto CC

  Select Case mButtonStyle
    Case CmdButtonStyle.CommandButton
      If BDown And Not OutSide Then
        RaiseEvent OwnerDrawDownState(CC, 0, 0, dx, dy, CancelBuiltIn)
        If Not CancelBuiltIn Then DrawDownState CC
      Else
        RaiseEvent OwnerDrawUpState(CC, 0, 0, dx, dy, CancelBuiltIn)
        If Not CancelBuiltIn Then DrawUpState CC
      End If

    Case CmdButtonStyle.CheckButton
      RaiseEvent OwnerDrawCheckState(CC, 0, 0, dx, dy, CancelBuiltIn)
      If Not CancelBuiltIn Then
        Select Case mValue
          Case 0: DrawUpState CC
          Case 1, 2: DrawDownState CC
        End Select
      End If
      
    Case CmdButtonStyle.OptionButton
      RaiseEvent OwnerDrawOptionState(CC, 0, 0, dx, dy, CancelBuiltIn)
      If Not CancelBuiltIn Then
        If mValue Then DrawDownState CC Else DrawUpState CC
      End If
      
    Case CmdButtonStyle.CheckBox
      RaiseEvent OwnerDrawCheckState(CC, 0, 0, dx, dy, CancelBuiltIn)
      If Not CancelBuiltIn Then DrawBoxState CC, thmTypeCheckBox
      
    Case CmdButtonStyle.OptionBox
      RaiseEvent OwnerDrawOptionState(CC, 0, 0, dx, dy, CancelBuiltIn)
      If Not CancelBuiltIn Then DrawBoxState CC, thmTypeOptionBox

  End Select
End Sub

Private Sub DrawBoxState(CC As cCairoContext, BoxType As enmThemeDrawingType)
Dim BoxSize As Single, BoxOffs As Single
Dim x As Single, y As Single, Tdy As Double, Tdx As Double
Dim TWidth As Single, TRowHeight As Double, TRows As Long, State As enmThemeDrawingState

  W.BackColor = -1
  
  BoxOffs = IIf(BoxType = thmTypeOptionBox, 0.5, 0)
  BoxSize = 13 - 2 * BoxOffs
  x = 3 + BoxOffs
  y = Int((dy - BoxSize) \ 2) + BoxOffs
  If y < BoxOffs Then y = BoxOffs
  
  If mValue Then State = thmStateChecked
  If mValue = 2 Then State = thmStateCheckUndefined
  Cairo.Theme.DrawTo CC, W, BoxType, State, x, y, BoxSize, BoxSize, BoxSize / 2

  'first we calculate the needed TextRows, depending on the text-rectangles width (dx - xoffs - 1)
  x = BoxSize + 11
  TRows = CC.DrawText(x, 0, dx - x - 1, 0, mCaption, False, vbLeftJustify, 0, True, dtHasAccelerators)
  Tdy = TRows * CC.GetFontHeight + 1
  CC.DrawText x, Int((dy - Tdy) / 2), dx - x - 1, Tdy, mCaption, False, vbLeftJustify, 0, True, dtHasAccelerators

  If W.Focused Then
    x = x - 3
    Tdx = dx - x
    If Tdy > dy Then Tdy = dy
    
    If TRows = 1 Then 'no text-break: draw focus-rect only as wide as the caption-text (with some offset)
      TWidth = CC.GetTextExtents(mCaption)
      If Len(W.AccessKeys) Then TWidth = TWidth - CC.GetTextExtents("&")
      If TWidth + 7 < Tdx Then Tdx = TWidth + 7
    End If
    Cairo.Theme.DrawTo CC, W, thmTypeDottedRectangle, 0, x, Int((dy - Tdy) / 2), Tdx, Tdy, mBorderRadius - 1
  End If
End Sub

Private Sub DrawUpState(CC As cCairoContext)
Dim State As enmThemeDrawingState
  State = IIf(W.MouseOver, thmStateHovered, 0)
  Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, 0, 0, dx, dy, mBorderRadius, thmDirectionDown
  
  CC.RoundedRect 2, 2, dx - 4, dy - 4, mBorderRadius - 2
  CC.Clip
    Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 1, 1, dx - 2, Int(dy / 2) - 0.2, mBorderRadius - 1, thmDirectionDown
  CC.ResetClip
  
  DrawIconAndCaption CC, 0
  
  CC.SetLineWidth 1, True
  Cairo.Theme.DrawTo CC, W, thmTypeBorder, 0, 0, 0, dx, dy, mBorderRadius

  If W.Enabled Then
    If W.Focused Then Cairo.Theme.DrawTo CC, W, thmTypeDottedRectangle, 0, 2, 2, dx - 4, dy - 4, mBorderRadius - 1
    DrawBorderStateOverlays CC
  End If
End Sub

Private Sub DrawDownState(CC As cCairoContext)
Dim State As enmThemeDrawingState
  State = IIf(W.MouseOver, thmStateHovered Or thmStatePressed, thmStatePressed)
  Cairo.Theme.DrawTo CC, W, thmTypeButtonFace, State, 0, 0, dx, dy, mBorderRadius, thmDirectionDown
  Cairo.Theme.DrawTo CC, W, thmTypeShine, State, 1, 1, dx - 2, dy - 2, mBorderRadius - 1, thmDirectionDown

  DrawIconAndCaption CC, DownStateCaptionOffset
  
  CC.SetLineWidth 1, True
  Cairo.Theme.DrawTo CC, W, thmTypeBorder, thmStatePressed, 0, 0, dx, dy, mBorderRadius
  
  If W.Enabled Then
    If W.Focused Then Cairo.Theme.DrawTo CC, W, thmTypeDottedRectangle, 0, 2, 2, dx - 4, dy - 4, mBorderRadius - 1
    DrawBorderStateOverlays CC
  End If
End Sub

Private Sub DrawIconAndCaption(CC As cCairoContext, Optional ByVal PxlOffs As Single)
Dim ImgSize As Single, ImgOffsX As Single, ImgOffsY As Single
Dim TOffsX As Single, TRows As Long, Tdx As Single, Tdy As Single
  
  If PxlOffs Then PxlOffs = CLng(PxlOffs * W.Root.Zoom) / W.Root.Zoom
  TOffsX = 5
  
  If Len(W.ImageKey) Then 'render the Image
    ImgSize = IIf(ImageSize, ImageSize, dy - TOffsX - 2)
    ImgOffsY = Int((dy - ImgSize) / 2)
    Select Case W.ImageKeyRenderBehaviour
      
      Case ImgKeyRenderLeft:        ImgOffsX = TOffsX + ImgOffsY - 3
      Case ImgKeyRenderCentered:    ImgOffsX = (dx - ImgSize) / 2
      Case ImgKeyRenderRight:       ImgOffsX = dx - ImgSize - TOffsX - ImgOffsY
      Case ImgKeyRenderTopCentered: ImgSize = IIf(ImageSize, ImageSize, CC.GetFontHeight * 1.1)
                                    ImgOffsY = 2
                                    ImgOffsX = (dx - ImgSize) / 2
    End Select
    CC.RenderSurfaceContent W.ImageKey, PxlOffs + ImgOffsX, PxlOffs + ImgOffsY, ImgSize, ImgSize
    ImgOffsY = IIf(W.ImageKeyRenderBehaviour = ImgKeyRenderTopCentered, ImgOffsY + ImgSize, 0)
  End If
  
  If Len(mCaption) Then
    Tdx = dx - 6
    If mHAlignment = vbCenter Then
      If W.ImageKeyRenderBehaviour = ImgKeyRenderLeft Then
        TOffsX = ImgOffsX + ImgSize
        Tdx = dx - TOffsX - 3
      Else
        TOffsX = 3
      End If
    ElseIf mHAlignment = vbLeftJustify And W.ImageKeyRenderBehaviour = ImgKeyRenderLeft Then
      TOffsX = ImgOffsX + ImgSize + 3
      Tdx = dx - TOffsX - 3
    End If
    TRows = CC.DrawText(PxlOffs + TOffsX, 0, Tdx, 0, mCaption, False, mHAlignment, 0, True, dtHasAccelerators)
    Tdy = TRows * CC.GetFontHeight + 1.1
    
    If W.ImageKeyRenderBehaviour = ImgKeyRenderTopCentered Then
      CC.DrawText PxlOffs + TOffsX, PxlOffs + ImgOffsY + (((dy - ImgOffsY) - Tdy) / 2), _
                  Tdx, Tdy, mCaption, False, mHAlignment, 0, True, dtHasAccelerators
    Else
      CC.DrawText PxlOffs + TOffsX, PxlOffs + ImgOffsY + ((dy - Tdy) / 2) + 0.5, _
                  Tdx, Tdy, mCaption, False, mHAlignment, 0, True, dtHasAccelerators
    End If
  End If
End Sub

Private Sub DrawBorderStateOverlays(CC As cCairoContext)
Dim BorderOverlayColor As Long
  BorderOverlayColor = W.FocusColor
  
  If W.Focused Or DefaultStateCondition Then
    If Not (BDown And Not OutSide) Then CC.SetLineWidth 2.3
    CC.RoundedRect 0, 0, dx, dy, mBorderRadius, True

    CC.SetSourceColor BorderOverlayColor, Alpha * 0.5
    CC.Stroke
  End If
  If W.MouseOver Then
    CC.SetLineWidth 1
    CC.RoundedRect 0, 0, dx, dy, mBorderRadius, True

    CC.SetSourceColor BorderOverlayColor, Alpha * 0.5, 0.78
    CC.Stroke
  End If
End Sub


