VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum eMoveCursorType
   movecursor_Char   ' Offset cursor by characters
   movecursor_Word   ' Offset cursor by words
   movecursor_Row    ' Offset cursor by lines
End Enum

Event Click()
Event Scroll()
Event Change()
Event OwnerDrawBackGround(CC As cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, CancelBuiltIn As Boolean)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Validate(Cancel As Boolean)
Event MaxLengthViolation() ' An insert operation was cancelled because it would violate the MaxLength property value
Event BeforePaste(TextToPaste As String, Cancel As Boolean)

Private dx As Single, dy As Single, Alpha As Single, Zoom As Single, dxx As Single, dyy As Single
Private mText$, mPassWordChar$, mMultiLine As Boolean, mMaxRowWidth As Single, mVisibleRows&, mCueBannerText As String
Private RowsCharOffset() As Long, RowsCharCount() As Long, mRowCount As Long
Private mSelStart&, mSelLength&, SelBegin&
Private CaretPos&, CurRow&, CurPos&, CaretInvertCounter&
Private ShiftDown As Boolean, BDown As Boolean
Private mBorder As Boolean, mScrollBars As ScrollBarConstants
Private mRowHeight&, mLastTextHeight As Long
Private mInnerSpace&, mAlignment As AlignmentConstants, mVCenter As Boolean
Private mScrollerSize!, VSValue&, VSLargeChange&, VSMax&
Private mLocked As Boolean
Private mMaxLength As Long

Private WithEvents T As cTimer, WithEvents tCaret As cTimer
Attribute T.VB_VarHelpID = -1
Attribute tCaret.VB_VarHelpID = -1
Private WithEvents VScrollBar As cwVScrollBar, WithEvents HScrollBar As cwHScrollBar
Attribute VScrollBar.VB_VarHelpID = -1
Attribute HScrollBar.VB_VarHelpID = -1

Private InsideDraw As Boolean
Private mUndoDepth As Byte, HScrollX&, HScrollMax&
Private mAutoSelectAll As Boolean
Private mHideSelection As Boolean

Public UserdefinedHovering As Boolean

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property

Private Sub Class_Initialize()
  ReDim RowsCharOffset(0)
  ReDim RowsCharCount(0)
  
  mInnerSpace = 1
  mBorder = True
  Zoom = 1
  
  Set W = Cairo.WidgetBase
  W.BackColor = vbWhite
  W.ForeColor = vbBlack
  W.ImplementsWheelMessages = True
  W.MousePointer = IDC_IBEAM
  W.RuntimePropertiesCommaSeparated = "SelStart,SelLength,SelText,TopRow"
  mScrollerSize = Cairo.Theme.GetScrollerSize
End Sub

Public Property Get Text$()
  Text = mText
End Property
Public Property Let Text(ByVal NewValue$)
  If NewValue = mText Then Exit Property
  mText = NewValue
  mSelStart = 0
  mSelLength = 0
  CalcRows
  W.Refresh
  RaiseEvent Change
  W.RaiseBubblingEvent Me, "Change"
End Property

Public Property Get CueBannerText$()
  CueBannerText = mCueBannerText
End Property
Public Property Let CueBannerText(ByVal NewValue$)
  If NewValue = mCueBannerText Then Exit Property
  mCueBannerText = NewValue
  W.Refresh
End Property

Public Property Get Locked() As Boolean
  Locked = mLocked
End Property
Public Property Let Locked(ByVal NewValue As Boolean)
  mLocked = NewValue
End Property

Public Property Get UndoDepth() As Byte
  UndoDepth = mUndoDepth
End Property
Public Property Let UndoDepth(ByVal NewValue As Byte)
  mUndoDepth = NewValue
End Property

Public Property Get InnerSpace() As Long
  InnerSpace = mInnerSpace
End Property
Public Property Let InnerSpace(ByVal NewValue As Long)
  If NewValue < 0 Then NewValue = 0
  If NewValue > dx \ 3 Then NewValue = dx \ 3
  mInnerSpace = NewValue
  AdjustDimensions
End Property

Public Property Get Alignment() As AlignmentConstants
  Alignment = mAlignment
End Property
Public Property Let Alignment(ByVal NewValue As AlignmentConstants)
  If NewValue < 0 Then NewValue = 0
  If NewValue > 2 Then NewValue = 2
  mAlignment = NewValue
  W.Refresh
End Property

Public Property Get VCenter() As Boolean
  VCenter = mVCenter
End Property
Public Property Let VCenter(ByVal NewValue As Boolean)
  mVCenter = NewValue
  W.Refresh
End Property

Public Property Get PassWordChar$()
  PassWordChar = mPassWordChar
End Property
Public Property Let PassWordChar(ByVal NewValue$)
  mPassWordChar = Left$(NewValue, 1)
  AdjustDimensions
End Property

Public Property Get MaxLength() As Long
   MaxLength = mMaxLength
End Property
Public Property Let MaxLength(ByVal p_MaxLengthInCharsOrZeroForUnlimited As Long)
   If p_MaxLengthInCharsOrZeroForUnlimited < 0 Then Err.Raise 5, , "Invalid maximum length. Must be >=0"
   
   mMaxLength = p_MaxLengthInCharsOrZeroForUnlimited
   
   If mMaxLength > 0 Then
      If Len(mText) > mMaxLength Then
         Me.Text = Left$(mText, mMaxLength)
         
         If Right$(mText, 1) = vbCr Then
            Me.Text = Left$(mText, Len(mText) - 1)
         End If
      End If
   End If
End Property

Public Property Get SelStart&()
  SelStart = mSelStart
End Property
Public Property Let SelStart(ByVal NewValue&)
  If NewValue < 0 Then NewValue = 0
  If NewValue > Len(mText) Then NewValue = Len(mText)
  mSelStart = NewValue: mSelLength = 0
  CaretPos = mSelStart: SelBegin = CaretPos
  CalcCoords CaretPos, CurRow, CurPos
  CaretInvertCounter = 0
  W.Refresh
End Property

Public Property Get SelLength&()
  SelLength = mSelLength
End Property
Public Property Let SelLength(ByVal NewValue As Long)
  If NewValue < 0 Then NewValue = 0
  If NewValue > Len(mText) Then NewValue = Len(mText)
  mSelLength = NewValue
  CaretPos = mSelStart + mSelLength
  If mSelLength = 0 Then SelBegin = mSelStart
  CalcCoords CaretPos, CurRow, CurPos
  CaretInvertCounter = 0
  EnsureVisible
End Property

Public Property Get SelText$()
  If mSelLength Then SelText = Mid$(mText, mSelStart + 1, mSelLength)
End Property
Public Property Let SelText(ByVal NewValue$)
  InsertText NewValue
  W.Refresh
End Property

Public Property Get AutoSelectAll() As Boolean
  AutoSelectAll = mAutoSelectAll
End Property

Public Property Let AutoSelectAll(ByVal AutoSelectAll As Boolean)
  mAutoSelectAll = AutoSelectAll
End Property

Public Property Get HideSelection() As Boolean
   HideSelection = mHideSelection
End Property
Public Property Let HideSelection(ByVal HideSelection As Boolean)
   mHideSelection = HideSelection
End Property

Public Property Get Border() As Boolean
  Border = mBorder
End Property
Public Property Let Border(ByVal NewValue As Boolean)
  mBorder = NewValue
  AdjustDimensions
End Property

Public Property Get MultiLine() As Boolean
  MultiLine = mMultiLine
End Property
Public Property Let MultiLine(ByVal NewValue As Boolean)
  If NewValue = mMultiLine Then Exit Property
  mMultiLine = NewValue
  If Not ((mScrollBars And vbHorizontal) <> 0 Or mMultiLine = False) Then HScrollX = 0

  AdjustDimensions
End Property

Public Property Get ScrollBars() As ScrollBarConstants
  ScrollBars = mScrollBars
End Property
Public Property Let ScrollBars(ByVal NewValue As ScrollBarConstants)
  If NewValue < 0 Or NewValue > 3 Then Exit Property
  If mScrollBars = NewValue Then Exit Property
  mScrollBars = NewValue
  If mScrollBars And vbHorizontal Then
    If HScrollBar Is Nothing Then 'add one
      HScrollX = 0
      Set HScrollBar = New cwHScrollBar
      HScrollBar.Widget.InnerWidget = True
      HScrollBar.Widget.TabStop = False
      W.Widgets.Add HScrollBar, "HScrollBar"
      HScrollBar.SmallChange = 2
    End If
  Else
    HScrollX = 0: HScrollMax = 0
    If Not HScrollBar Is Nothing Then 'free resources
      Set HScrollBar = Nothing
      W.Widgets.Remove "HScrollBar"
    End If
  End If
  
  AdjustDimensions
End Property

Public Sub SelectAll()
  mSelStart = 0
  SelLength = Len(mText)
  CaretPos = Len(mText)
  CalcCoords CaretPos, CurRow, CurPos
End Sub

Public Function RowCount&()
  RowCount = mRowCount
End Function

Public Property Get TopRow&()
  TopRow = VSValue
End Property
Public Property Let TopRow(ByVal NewValue&)
  If Not mMultiLine Then Exit Property
  VSValue = NewValue
  If Not VScrollBar Is Nothing Then
    VScrollBar.Max = VSMax
    VScrollBar.LargeChange = VSLargeChange
    VScrollBar.Value = VSValue
    If VScrollBar.Value <> VSValue Then
      VSValue = VScrollBar.Value
    End If
  End If
  CaretInvertCounter = 0
  W.Refresh
End Property

Public Property Get RowHeight&()
  If mRowHeight = 0 Then
    If mLastTextHeight <= 0 Then
      RowHeight = 14
    Else
      RowHeight = mLastTextHeight
    End If
  Else
    RowHeight = mRowHeight
  End If
End Property
Public Property Let RowHeight(ByVal NewValue&)
  If NewValue < 0 Then NewValue = 0 Else If NewValue > 50 Then NewValue = 50
  mRowHeight = NewValue
End Property

Public Function VisibleRows&()
  VisibleRows = mVisibleRows
End Function

Public Sub EnsureVisible()
Dim NewTop&, NewLeft&, pS&, Ext As Single, MCC As cCairoContext
Dim TWidth As Single, AlgnOffs As Single, CurHScrollValue&, OldFont&
 
  NewTop = VSValue
  If (CurRow - NewTop + 1 > mVisibleRows - 2) Or CurRow < NewTop Then
    If CurRow > NewTop + mVisibleRows - 2 Then
      NewTop = CurRow - mVisibleRows + 2
    ElseIf CurRow < NewTop Then
      NewTop = CurRow
    End If

    If NewTop < 0 Then NewTop = 0
    If (NewTop > 0 And Not mMultiLine) Or CurRow = 0 Then NewTop = 0
    If NewTop <> VSValue Then TopRow = NewTop
  End If
  
  If (mScrollBars And vbHorizontal) <> 0 Or mMultiLine = False Then
    pS = StrPtr(mText) + RowsCharOffset(CurRow)

    Set MCC = SetCurFont
    
      TWidth = MCC.GetTextExtentPtr(pS, Abs(RowsCharCount(CurRow)))
      If (dxx - 3) - mMaxRowWidth > 0 Then HScrollX = 0
      AlgnOffs = GetAlignOffset(TWidth)
      
      Ext = MCC.GetTextExtentPtr(pS, CurPos)  '- AlgnOffs
  
      CurHScrollValue = HScrollX - AlgnOffs
      If (Ext - CurHScrollValue + 1 > dxx - 2) Or Ext < CurHScrollValue - 1 Then
        If Ext > CurHScrollValue + dxx - 3 Then
          NewLeft = Ext - dxx + 3
        ElseIf Ext < CurHScrollValue Then
          NewLeft = Ext
        End If
        If NewLeft <> CurHScrollValue Then
          HScrollX = NewLeft + AlgnOffs
          If HScrollX > HScrollMax Then HScrollX = HScrollMax
          If HScrollX < 0 Then HScrollX = 0
          If CaretPos = 0 And HScrollX <= 2 Then HScrollX = 0
          If Not HScrollBar Is Nothing Then
            HScrollBar.Value = HScrollX
          End If
        End If
      End If
      
    If Not InsideDraw Then MCC.Restore
  End If
  CaretInvertCounter = 0

  W.Refresh
End Sub

Private Sub HScrollBar_Change()
  HScrollX = HScrollBar.Value
  HScrollMax = HScrollBar.Max
  RaiseEvent Scroll
  W.Refresh
End Sub
Private Sub HScrollBar_Scroll()
  HScrollBar_Change
End Sub

Private Sub tCaret_Timer() 'only running, when focused
  If W.IsHidden Then Exit Sub ' Caret is outside the active Drawing-Rect
  
  If Not W.Root.IsTopWindowActive Then    ' our Parent-Window is deactivated
      If CaretInvertCounter <> 12 Then
        CaretInvertCounter = 12
        If Not InsideDraw Then W.Refresh
      End If
      Exit Sub
  End If
  
  CaretInvertCounter = (CaretInvertCounter + 1) Mod 24
  If (CaretInvertCounter Mod 12) = 0 Then
    If Not InsideDraw Then W.Refresh
  End If
End Sub

Private Sub VScrollBar_Change()
  VSValue = VScrollBar.Value
  W.Refresh
  RaiseEvent Scroll
End Sub
Private Sub VScrollBar_Scroll()
  VScrollBar_Change
End Sub

Private Sub W_Click()
   RaiseEvent Click
End Sub

Private Sub W_DblClick()
Dim i&, S$, WLeft%, WRight%, W%, PLeft&, PRight&
 
  S = Mid$(mText, RowsCharOffset(CurRow) \ 2 + 1, Abs(RowsCharCount(CurRow)))
  
  If Len(S) = 0 Then Exit Sub
  If CurPos > 0 Then
    WLeft = AscW(Mid$(S, CurPos, 1))
  End If
  If CurPos < Len(S) Then
    WRight = AscW(Mid$(S, CurPos + 1, 1))
  End If
  If WRight > WLeft Then W = WRight Else W = WLeft 'get the maximum
  If W <= 32 Then Exit Sub 'do nothing

'  If CurPos = Len(S) Then CurPos = CurPos - 1: mSelStart = mSelStart - 1
  If (W < 48 Or (W = 58)) And Not (W > 32 And W < 39 And W <> 34) Then  'non alphanumeric chars to expand
    For i = CurPos - 1 To 0 Step -1
      W = AscW(Mid$(S, i + 1, 1))
      If W <= 32 Or (W >= 48 And Not ((W > 57 And W < 65) Or (W > 90 And W < 97) Or W > 122)) Then Exit For
    Next i
    PLeft = i + 1: If PLeft < 0 Then PLeft = 0
    For i = CurPos To Len(S) - 1
      W = AscW(Mid$(S, i + 1, 1))
      If W <= 32 Or (W >= 48 And Not ((W > 57 And W < 65) Or (W > 90 And W < 97) Or W > 122)) Then Exit For
    Next i
    PRight = i: If PRight > Len(S) Then PRight = Len(S)
  Else 'chars in the alphanumeric-range are expanded
    For i = CurPos - 1 To 0 Step -1
      W = AscW(Mid$(S, i + 1, 1))
      If (W < 48 Or ((W > 57 And W < 65) Or (W > 90 And W < 97) Or W > 122)) And Not (W > 32 And W < 39 And W <> 34) Then Exit For
    Next i
    PLeft = i + 1: If PLeft < 0 Then PLeft = 0
    For i = CurPos To Len(S) - 1
      W = AscW(Mid$(S, i + 1, 1))
      If (W < 48 Or ((W > 57 And W < 65) Or (W > 90 And W < 97) Or W > 122)) And Not (W > 32 And W < 39 And W <> 34) Then Exit For
    Next i
    PRight = i: If PRight > Len(S) Then PRight = Len(S)
  End If
  
  If PLeft < CurPos Then mSelStart = mSelStart - (CurPos - PLeft)
  SelBegin = mSelStart ' + mSelLength
  SelLength = PRight - PLeft
  CaretPos = mSelStart + mSelLength
  CalcCoords CaretPos, CurRow, CurPos
  EnsureVisible
End Sub

Private Sub W_FontChanged()
  AdjustDimensions
End Sub

Private Sub W_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown As Boolean
Dim NewRow As Long
Dim NewPos As Long

  RaiseEvent KeyDown(KeyCode, Shift)
  ShiftDown = Shift And vbShiftMask
  CtrlDown = Shift And vbCtrlMask
  On Error Resume Next
  
  Select Case KeyCode
    Case vbKeyReturn: If mMultiLine Then InsertText vbCrLf

    Case vbKeyDelete: If ShiftDown And mSelLength > 0 Then CopyText: DeleteText
                      If Not ShiftDown Then DeleteText
    Case vbKeyUp: MoveCursor -1, movecursor_Row
    Case vbKeyDown: MoveCursor 1, movecursor_Row
    Case vbKeyRight
      MoveCursor 1, IIf(CtrlDown, movecursor_Word, movecursor_Char)
    
    Case vbKeyLeft
      MoveCursor -1, IIf(CtrlDown, movecursor_Word, movecursor_Char)
    
    Case vbKeyPageUp
      If mMultiLine Then
        MoveCursor -mVisibleRows + 1, movecursor_Row
      End If
      
    Case vbKeyPageDown
      If mMultiLine Then
        MoveCursor mVisibleRows - 1, movecursor_Row
        
        If (CurRow < RowCount - 1) Then
          ' Cursor not at bottom of visible textbox area,
          ' so bump the current caret row to the top of the textbox
          TopRow = CurRow
        End If
      End If
      
    Case vbKeyHome
      If mMultiLine And (Not CtrlDown) Then
        ' Move to beginning of current line
        NewPos = InStrRev(mText, vbNewLine, CaretPos)
        If NewPos <= 0 Then
          NewPos = 0
        Else
          NewPos = NewPos + 1
        End If
      Else
        ' Move to beginning of text
        NewPos = 0
      End If
      
      MoveCursor NewPos - CaretPos
    
    Case vbKeyEnd
      If mMultiLine And (Not CtrlDown) Then
        ' Move to end of current line
        NewPos = InStr(CaretPos + 1, mText, vbNewLine)
        If NewPos <= 0 Then
          NewPos = Len(mText)
        Else
          NewPos = NewPos - 1
        End If
      Else
        ' Move to end of text
        NewPos = Len(mText)
      End If
      
      MoveCursor NewPos - CaretPos
    
    Case vbKeyInsert: If CtrlDown Then CopyText Else If ShiftDown Then PasteText
    Case vbKeyX: If CtrlDown And mSelLength > 0 Then CopyText: DeleteText
    Case vbKeyC: If CtrlDown Then CopyText
    Case vbKeyV: If CtrlDown Then PasteText
    Case vbKeyA
      If CtrlDown Then
        Me.SelectAll
      End If

    Case Else: Exit Sub
  End Select
  
  If Err Then Err.Clear
  EnsureVisible
End Sub
Private Sub W_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
  On Error Resume Next
    Select Case KeyAscii
      Case vbKeyBack: DeleteText True
      Case Else
        If KeyAscii >= vbKeySpace Then
          InsertText ChrW$(KeyAscii)
        End If
    End Select
  If Err Then Err.Clear
  EnsureVisible
End Sub
Private Sub W_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
  If KeyCode = vbKeyShift Then ShiftDown = False
End Sub

Private Sub InsertText(NewText$)
  Dim Buffer As String
    
  If mLocked Then Exit Sub
  If Len(NewText) = 0 And mSelLength = 0 Then Exit Sub
  
  If mMaxLength > 0 Then
   If Len(NewText) + Len(mText) - mSelLength > mMaxLength Then
     ' Can't insert text without violating MaxLength property value
     RaiseEvent MaxLengthViolation
     Exit Sub
   End If
  End If
   
  mText = Left$(mText, mSelStart) & NewText & Mid$(mText, mSelStart + mSelLength + 1)
  CaretPos = mSelStart + Len(NewText)
  CalcRows
  mSelLength = 0: mSelStart = CaretPos
  SelBegin = CaretPos
  RaiseEvent Change
  W.RaiseBubblingEvent Me, "Change"
End Sub

Private Sub CopyText()
  If mSelLength = 0 Or Len(mPassWordChar) > 0 Then Exit Sub

  Cairo.Clipboard.Clear
  Cairo.Clipboard.SetText SelText, vbCFText
End Sub

Private Sub PasteText()
  Dim Text As String
  Dim Cancel As Boolean
  
  Text = Cairo.Clipboard.GetText
  
  RaiseEvent BeforePaste(Text, Cancel)
  If Cancel Then Exit Sub
  
  InsertText Text
End Sub

Private Sub DeleteText(Optional ByVal ToTheLeft As Boolean)
  If mSelLength = 0 Then
    If ToTheLeft And mSelStart > 0 Then
      If Mid$(mText, mSelStart, 1) = vbLf Then
        mSelStart = mSelStart - 2: mSelLength = 2
      Else
        mSelStart = mSelStart - 1: mSelLength = 1
      End If
    ElseIf Not ToTheLeft And mSelStart < Len(mText) Then
      If Mid$(mText, mSelStart + 1, 1) = vbCr Then mSelLength = 2 Else mSelLength = 1
    End If
  End If
  InsertText ""
End Sub

Private Sub MoveCursor(ByVal Offs&, Optional ByVal MoveType As eMoveCursorType = movecursor_Char)
  Dim i As Long
  Dim LoopStart As Long
  Dim LoopEnd As Long
  Dim LoopStep As Long
  Dim InWordBreak As Boolean
  Dim MoveCount As Long
  Dim FoundNonWhitespace As Boolean
  
  If Offs = 0 Then Exit Sub   ' No movement, short-circuit
  
  Select Case MoveType
  Case movecursor_Row
    ' Move cursor to line
    CurRow = CurRow + Offs
    If CurRow < 0 Then CurRow = 0 Else If CurRow > mRowCount - 1 Then CurRow = mRowCount - 1
    If CurPos > Abs(RowsCharCount(CurRow)) Then CurPos = Abs(RowsCharCount(CurRow))
    CalcCaretPos
    
  Case movecursor_Char
    ' Move cursor to character
    CaretPos = CaretPos + Offs
    If RowsCharCount(CurRow) <= 0 And Abs(RowsCharCount(CurRow)) = CurPos And Offs > 0 Then
      CaretPos = CaretPos + 1 'linebreak-correction
    End If
    If CurRow > 0 And Offs < 0 Then 'look one row above for linebreaks
      If RowsCharCount(CurRow - 1) <= 0 And CurPos = 0 Then CaretPos = CaretPos - 1
    End If
    If CaretPos < 0 Then CaretPos = 0 Else If CaretPos > Len(mText) Then CaretPos = Len(mText)
    CalcCoords CaretPos, CurRow, CurPos
   
  Case movecursor_Word
    ' Move cursor to word break
    If Offs < 0 Then
      ' Moving backwards
      LoopStart = CaretPos
      LoopEnd = 1
      LoopStep = -1
      Offs = Offs - 1   ' HACK: For backward movement, we'll move back to the end of an extra word,
                        '       then move forward to find the beginning of the word
    Else
      ' Moving forwards
      LoopStart = CaretPos + 1
      LoopEnd = Len(mText)
      LoopStep = 1
    End If
    
    For i = LoopStart To LoopEnd Step LoopStep
      Select Case Asc(Mid$(mText, i, 1))
      Case 32, 9, 10, 13  ' Space, Tab, CR or LF, i.e. Wordbreak character
        If (Offs < 0 And FoundNonWhitespace) Or (Offs > 0) Then
          If Not InWordBreak Then
            ' Signal that we are in a stretch of one or more whitespace characters
            InWordBreak = True
          End If
        End If
        
      Case Else   ' Non-wordbreak character
        If InWordBreak Then
          ' Set cursor position here
          CaretPos = i - IIf(Offs < 0, 0, 1)
          
          MoveCount = MoveCount + LoopStep
          If MoveCount = Offs Then
            ' We've moved the requisite number of words, so short-circuit the loop
            FoundNonWhitespace = False
              
            Exit For
          End If
        End If
        
        FoundNonWhitespace = True
      End Select
    Next i
    
    If i = 0 And Offs < 0 Then
      ' We didn't find a word break before hitting the beginning of the text in backward search
      ' So move the caret to the beginning of the text
      CaretPos = 0
    
    ElseIf i > Len(mText) And Offs > 0 Then
      ' We didn't find a word break before hitting the end of the text in forward search
      ' So move the caret to the end of the text
      CaretPos = Len(mText)
    End If
    
    CalcCoords CaretPos, CurRow, CurPos
  
    If Offs < 0 Then
      ' HACK: for backward movement
      If CaretPos > 0 Then
         ' Not at beginning of text box, so move forward one word
         MoveCursor 1, movecursor_Word
      End If
   End If
  
  Case Else
    Err.Raise 5, , "Unknown cursor move type: " & MoveType
    
  End Select
  
  If ShiftDown Then
    If SelBegin < CaretPos Then mSelStart = SelBegin Else mSelStart = CaretPos
    mSelLength = Abs(SelBegin - CaretPos)
  Else
    mSelStart = CaretPos: mSelLength = 0: SelBegin = CaretPos
  End If
End Sub

Private Sub W_GotFocus()
  If mAutoSelectAll Then
    Me.SelectAll
  End If
  Set tCaret = W.Root.CreateTimer(40, True)
  CaretInvertCounter = 0
End Sub

Private Sub W_LostFocus()
  Set tCaret = Nothing
  CaretInvertCounter = 1 'somewhere above 0
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As cWidgetBase)
  If Not UserdefinedHovering Then W.Refresh
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As cWidgetBase)
  If Not UserdefinedHovering Then W.Refresh
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button <> 1 Then Exit Sub
  If Button = 1 Then BDown = True
  If Len(mText) = 0 Then Exit Sub
  CalcCaretPosFromMouseCoords x, y
  If ShiftDown Then
    If CaretPos < SelBegin Then mSelStart = CaretPos Else mSelStart = SelBegin
    mSelLength = Abs(SelBegin - CaretPos)
  Else
    mSelLength = 0: mSelStart = CaretPos: SelBegin = CaretPos
  End If
  EnsureVisible
End Sub

Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim i&, dy&

  If BDown Then
  
    dy = (mVisibleRows - 1) * RowHeight
'    If (dyy Mod RowHeight) Then dy = (mVisibleRows - 2) * RowHeight
    
    If (y < 0 Or y > dy) And mMultiLine Then    'we are outside
      If T Is Nothing Then Set T = W.Root.CreateTimer(40, True)
      T.Tag = x & "_" & y:
      T.Interval = 40
      
      If y < 0 Then
        i = 200 + y * 10
        If i < 10 Then i = 10
        T.Interval = i
      ElseIf (y - dy) > 0 Then
        i = 200 - (y - dy) * 10
        If i < 10 Then i = 10
        T.Interval = i
      End If

    Else 'handling without scroll-timer-support
      Set T = Nothing
      CalcCaretPosFromMouseCoords x, y
      mSelLength = Abs(SelBegin - CaretPos)
      If CaretPos < SelBegin Then mSelStart = CaretPos Else mSelStart = SelBegin

      EnsureVisible
    End If
  End If
End Sub

Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  BDown = False
  Set T = Nothing
End Sub

Private Sub W_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
  If MouseKeys = 0 Then Spin Rotation, LineOffset
End Sub
Private Sub VScrollBar_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
  If MouseKeys = 0 Then Spin Rotation, LineOffset
End Sub
Private Sub Spin(ByVal Rotation As Long, ByVal LineOffset As Long)
  If Rotation > 0 Then
    If VSValue > 0 Then TopRow = VSValue + LineOffset
  ElseIf Rotation < 0 Then
    If VSValue < VSMax Then TopRow = VSValue + LineOffset
  End If
End Sub

Private Sub CalcCaretPosFromMouseCoords(ByVal x As Single, ByVal y As Single)
Dim pS&, Ext As Single, Ext2 As Single, MCC As cCairoContext
Dim BorderOffs&, TWidth As Single, AlgnOffs As Single, VOffs As Single
 
  BorderOffs = mInnerSpace
  If mBorder Then BorderOffs = BorderOffs + 1
  
  If mRowCount = 0 Then mRowCount = 1
  If mVCenter Then
    VOffs = (dy - mInnerSpace - BorderOffs - mRowCount * RowHeight - IIf(HScrollBar Is Nothing, 0, mScrollerSize)) / 2
 
    If mBorder Then
      If VOffs < -1 Then VOffs = -1
    Else
      If VOffs < -2 Then VOffs = -2
    End If
  End If
  y = y - VOffs - mInnerSpace
  
  If y > 0 Then
    CurRow = (y - BorderOffs) \ RowHeight + VSValue
  Else
    CurRow = (y - BorderOffs - RowHeight) \ RowHeight + VSValue
  End If
  If CurRow < 0 Then CurRow = 0 Else If CurRow > mRowCount - 1 Then CurRow = mRowCount - 1
  If (mScrollBars And vbHorizontal) <> 0 Or mMultiLine = False Then
    x = x + HScrollX
  End If
 
  If x < BorderOffs Then x = BorderOffs
  
  pS = StrPtr(mText) + RowsCharOffset(CurRow)
  

  Set MCC = SetCurFont
    'calculate Alignment-Offset
    If Alignment Then
      TWidth = MCC.GetTextExtentPtr(pS, Abs(RowsCharCount(CurRow)))
      AlgnOffs = GetAlignOffset(TWidth)
    End If
    x = x - AlgnOffs - BorderOffs
    
    CurPos = MCC.GetCharsFittingExtentPtr(pS, Abs(RowsCharCount(CurRow)), x)

    If CurPos < Abs(RowsCharCount(CurRow)) Then 'let's check, if we are in the middle of two adjacent Chars
      Ext = MCC.GetTextExtentPtr(pS, CurPos)
      Ext2 = MCC.GetTextExtentPtr(pS, CurPos + 1)
      If x > (Ext + Int((Ext2 - Ext - 3 / W.Zoom) / 2)) Then CurPos = CurPos + 1
    End If
    
    CalcCaretPos
 
  If Not InsideDraw Then MCC.Restore
End Sub

Private Sub SetScrollerValues()
Dim Rows&, Pixels&
  If mVisibleRows - 1 > mRowCount Then Rows = mRowCount + 1 Else Rows = mVisibleRows
  If mRowCount - Rows + 1 >= 0 Then
    VSMax = mRowCount - Rows + 1
    If Not VScrollBar Is Nothing Then
      VScrollBar.Max = VSMax
      VScrollBar.Widget.Refresh
    End If
  End If
  
  If mMaxRowWidth + 3 - dxx > 0 Then
    HScrollMax = mMaxRowWidth + 3 - dxx
  Else
    HScrollMax = 0
  End If
  If HScrollX > HScrollMax Then HScrollX = HScrollMax
  
  If mScrollBars And vbHorizontal Then
    HScrollBar.Max = HScrollMax
    HScrollBar.Value = HScrollX
    HScrollBar.Widget.Refresh
  End If
End Sub

Private Sub W_Resize()
  dx = W.Width
  dy = W.Height
  AdjustDimensions
End Sub

Private Sub W_Validate(Cancel As Boolean)
  RaiseEvent Validate(Cancel)
End Sub

Private Sub T_Timer()
Dim x&, y&, dy&

  If W.IsHidden Then Exit Sub
  dy = (mVisibleRows - 1) * RowHeight
  If (dyy Mod RowHeight) Then dy = (mVisibleRows - 2) * RowHeight
    
  x = Split(T.Tag, "_")(0)
  y = Split(T.Tag, "_")(1)

  ShiftDown = True
  If y < 0 Then 'Up
    CalcCaretPosFromMouseCoords x, y
    mSelLength = Abs(SelBegin - CaretPos)
    If CaretPos < SelBegin Then mSelStart = CaretPos Else mSelStart = SelBegin
    MoveCursor 0, movecursor_Row
  ElseIf (y - dy) > 0 Then
    CalcCaretPosFromMouseCoords x, y
    mSelLength = Abs(SelBegin - CaretPos)
    If CaretPos < SelBegin Then mSelStart = CaretPos Else mSelStart = SelBegin
    MoveCursor 0, movecursor_Row
  End If
  ShiftDown = False
  
  EnsureVisible
End Sub

Private Sub AdjustDimensions()
Dim BorderOffs!, Hdy!, Vdx!

  Zoom = W.Zoom
  mScrollerSize = Int(Cairo.Theme.GetScrollerSize * Zoom) / Zoom

  If (mScrollBars And vbVertical) <> 0 And mMultiLine Then Vdx = mScrollerSize: dxx = dx - Vdx Else dxx = dx
  If mScrollBars And vbHorizontal Then Hdy = mScrollerSize: dyy = dy - Hdy Else dyy = dy
  If mBorder Then
    dxx = dxx - 3
    dyy = dyy - 2
    BorderOffs = CLng(1 * Zoom) / Zoom - 0.1
  Else
    dxx = dxx - 1
  End If
  
  dxx = dxx - 2 * mInnerSpace
  dyy = dyy - 2 * mInnerSpace
  If dxx < 2 Then dxx = 2
  If dyy < 2 Then dyy = 2
  
  mVisibleRows = (dyy \ RowHeight)
  If (dyy Mod RowHeight) Then mVisibleRows = mVisibleRows + 1

  VSLargeChange = mVisibleRows - 2

  CalcRows
  
  On Error Resume Next
    If (mScrollBars And vbVertical) <> 0 And mMultiLine Then
      If VScrollBar Is Nothing Then
        Set VScrollBar = New cwVScrollBar
            VScrollBar.Widget.InnerWidget = True
        VScrollBar.Widget.TabStop = False
        W.Widgets.Add VScrollBar, "VScrollBar", dx - mScrollerSize - BorderOffs, BorderOffs + 0.5, , dy - 2 * BorderOffs - Hdy
        VScrollBar.Max = VSMax
        VScrollBar.LargeChange = VSLargeChange
        VScrollBar.Value = VSValue
      Else
        VScrollBar.Widget.Move dx - mScrollerSize - BorderOffs, BorderOffs + 0.5, , dy - 2 * BorderOffs - Hdy
        VScrollBar.Max = VSMax
        VScrollBar.LargeChange = VSLargeChange
      End If
      VScrollBar.Widget.Refresh
    Else
      If Not VScrollBar Is Nothing Then
        W.Widgets.Remove "VScrollBar"
        Set VScrollBar = Nothing
      End If
    End If
    If mScrollBars And vbHorizontal Then
      If mScrollerSize < dyy Then Hdy = mScrollerSize Else Hdy = dyy
      
      HScrollBar.Widget.Move BorderOffs + 0.5, dy - Hdy - BorderOffs, dx - 2 * BorderOffs - Vdx
      HScrollBar.LargeChange = dxx
      HScrollBar.Widget.Refresh
      If Hdy = dyy And Not VScrollBar Is Nothing Then
        VScrollBar.Widget.Move dx - mScrollerSize - BorderOffs, BorderOffs + 0.5, , dy - 2 * BorderOffs - Hdy '+ 1
      End If
    End If
  If Err Then Err.Clear
  
'  SetScrollerValues
'Debug.Print "AdjustDimensions"
  EnsureVisible
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  Alpha = W.AlphaInherited
  If dx <> dx_Aligned Or dy <> dy_Aligned Or Zoom <> W.Zoom Then
    dx = dx_Aligned
    dy = dy_Aligned
    AdjustDimensions
  End If
  Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
Dim i&, j&, CharOffs&, Color&, MaxIdx&, x1 As Single, y1 As Single, x2 As Single, y2 As Single, pS&, Ext As Single, EndExt As Single, TWidth As Single
Dim SelStartRow&, SelStartPos&, SelEndRow&, SelEndPos&, BorderOffs&, VStart&
Dim VOffs&, AlgnOffs As Single, CancelBuiltIn As Boolean
Dim Pat As cCairoPattern
Static LastTextHeight&
  
  If Not InsideDraw Then InsideDraw = True Else Exit Sub
  
  SetCurFont CC

  CC.SetLineWidth 1
  CC.SetLineCap Cairo_LINE_CAP_SQUARE
  
  mLastTextHeight = CC.GetFontHeight()
 
  If LastTextHeight = 0 Then LastTextHeight = 14
  If LastTextHeight <> mLastTextHeight Then
    LastTextHeight = mLastTextHeight
    AdjustDimensions
  End If
  
  If mRowCount <= 0 Then mRowCount = 1
  
  BorderOffs = mInnerSpace
  If mBorder Then BorderOffs = BorderOffs + 1
 
  If W.BackColor <> -1 Then
    CancelBuiltIn = False: RaiseEvent OwnerDrawBackGround(CC, 0, 0, dx + BorderOffs, dy, CancelBuiltIn)
    If Not CancelBuiltIn Then
      CC.SetSourceColor W.BackColor, Alpha
        CC.Rectangle 0, 0, dx + BorderOffs, dy
      CC.Fill
    End If
  End If
  
  Color = IIf(W.Focused, W.SelectionColor, W.ShadeColor(W.DisabledColor, 1.3))

  If mSelLength Then
    CalcCoords mSelStart, SelStartRow, SelStartPos
    CalcCoords mSelStart + mSelLength, SelEndRow, SelEndPos
  End If

  MaxIdx = IIf(VSValue + mVisibleRows > mRowCount, mRowCount - 1, VSValue + mVisibleRows - 1)
  pS = StrPtr(mText)

  If mVCenter Then
    VOffs = (dy - BorderOffs - 1 - mInnerSpace - mRowCount * RowHeight - IIf(HScrollBar Is Nothing, 0, mScrollerSize)) \ 2

    If mBorder Then
      If VOffs < -1 Then VOffs = -1
    Else
      If VOffs < -2 Then VOffs = -2
    End If

  End If
  
  If VSValue > MaxIdx Then VStart = MaxIdx Else VStart = VSValue
  If VStart < 0 Then VStart = 0
  'current x-coords
  x1 = BorderOffs + 1
  
  For i = VStart To MaxIdx

    y1 = j * RowHeight + BorderOffs + 1 + VOffs
    y2 = (j + 1) * RowHeight + BorderOffs + VOffs   'current y-Coords
    CharOffs = 0
   
    'Full Row-length
    TWidth = CC.GetTextExtentPtr(pS + RowsCharOffset(i), Abs(RowsCharCount(i))) + 1
    'calculate Alignment-Offset
    AlgnOffs = GetAlignOffset(TWidth) - HScrollX
   
    'Selection-Handling
    If mSelLength Then
      If (Not mHideSelection) Or W.Focused Then
        ' We either have focus, or we don't have focus and mHideSelection = False
        ' So draw the selection highlight
        If i = SelStartRow Then 'special handling in the first selected row
          Ext = CC.GetTextExtentPtr(pS + RowsCharOffset(i), SelStartPos)
        ElseIf i > SelStartRow And i < SelEndRow Then 'full row is selected
          Ext = 0
        End If
        EndExt = TWidth
        If i = SelEndRow Then 'special handling in the last selected Row
          If SelEndRow > SelStartRow Then Ext = 0
          EndExt = CC.GetTextExtentPtr(pS + RowsCharOffset(i), SelEndPos)
        End If
      
        If i >= SelStartRow And i <= SelEndRow Then
          CC.SetSourceColor Color, Alpha * 0.75 'prepare for selections
          CC.Rectangle (Ext + x1 + AlgnOffs), y1, (EndExt - Ext + 0.5), y2 - y1 + 1, False
          CC.Fill
        End If
      End If
    End If
    
    'now the Textout
    If Len(mPassWordChar) Then
      CC.TextOutPtr x1 + AlgnOffs, y1, StrPtr(String(Abs(RowsCharCount(i)), mPassWordChar)), Abs(RowsCharCount(i))
    Else
      CC.TextOutPtr x1 + AlgnOffs, y1, pS + RowsCharOffset(i), Abs(RowsCharCount(i))
    End If
  
     'Draw Caret
    If i = CurRow And W.Focused Then
      If CaretInvertCounter = 0 Then
        
        Ext = CC.GetTextExtentPtr(pS + RowsCharOffset(i), CurPos)
        
        If HScrollBar Is Nothing Then
          If Ext > dxx + HScrollX Then Ext = dxx + HScrollX
        Else
          If Ext > dxx + HScrollX And HScrollBar.Max = AlgnOffs Then Ext = dxx + HScrollX
        End If
  
        CC.SetLineWidth 1, True '<- last param ensures a rounded, pixel-aligned LineWidth
        CC.SetLineCap Cairo_LINE_CAP_BUTT
        CC.SetSourceColor vbBlack, Alpha
          x2 = Int((Ext + x1 + AlgnOffs + 0.5) * Zoom) / Zoom 'pixel-aligned x-Offset
          CC.DrawLine x2, y1 - 0.5, x2, y1 + RowHeight - 0.5, True
        CC.Stroke
      End If
    End If
    
    j = j + 1
  Next i
  
  If Len(mText) = 0 And Len(mCueBannerText) > 0 Then 'Shows the CueBanner Text
    CC.SelectFont W.FontName, W.FontSize, W.ShadeColor(W.DisabledColor, IIf(W.Focused, 1.2, 1)), , True
    CC.TextOut x1 + AlgnOffs, y1, mCueBannerText
  End If
  
  If mBorder Then
    CC.SetLineWidth 1, True
    Cairo.Theme.DrawTo CC, W, thmTypeBorder, 0, 0, 0, dx, dy
  End If
  InsideDraw = False
End Sub

Private Function GetAlignOffset(TWidth As Single) As Single
  Select Case Alignment
    Case vbLeftJustify: GetAlignOffset = 0
    Case vbCenter: GetAlignOffset = (dxx - TWidth - 1 + HScrollMax) / 2
     Case vbRightJustify
      If mBorder Then
        GetAlignOffset = (dxx - TWidth - 1) + HScrollMax + mInnerSpace - 2
      Else
        GetAlignOffset = (dxx - TWidth - 1) + HScrollMax + mInnerSpace - 1
      End If
  End Select
End Function
 
Private Sub CalcRows()  'calculates the Rows, taken up by the text + ensures correct LineBreaks
Dim MCC As cCairoContext

  If Len(mText) = 0 Then 'a short-cut, to avoid the more expensive calculations below
    ReDim RowsCharOffset(0)
    ReDim RowsCharCount(0)
    mRowCount = 0
    CaretPos = 0
    CurRow = 0
    CurPos = 0
    mMaxRowWidth = 0
    SetScrollerValues
    Exit Sub
  End If
  
  Set MCC = SetCurFont
  
  MCC.CalcTextRowsInfo mText, dxx, mMultiLine, (mScrollBars And vbHorizontal), RowsCharCount, RowsCharOffset, mRowCount, mMaxRowWidth
 
  SetScrollerValues 'since our RowCount is new at this point, adjust the scrollers to reflect this
  CalcCoords CaretPos, CurRow, CurPos
 
  If Not InsideDraw Then MCC.Restore
End Sub

Private Sub CalcCoords(ByVal LinearPos&, Row&, RowPos&)
Dim i&, CPos&, CCount&, CrLfOffs&
  Row = -1
  For i = 0 To mRowCount - 1
    CCount = RowsCharCount(i)
    If CCount > 0 Then CrLfOffs = 0 Else CrLfOffs = 2: CCount = -CCount 'vbCrLf-detection
    If LinearPos >= CPos And LinearPos <= CPos + CCount Then
      Row = i: RowPos = LinearPos - CPos
      If RowsCharCount(i) > 0 And RowsCharCount(i) = RowPos And i < mRowCount - 1 Then
        Row = Row + 1: RowPos = 0
      End If
      Exit Sub
    End If

    CPos = CPos + CCount + CrLfOffs

    If LinearPos < CPos And i < mRowCount - 1 Then
      Row = i + 1: RowPos = 0: LinearPos = RowPos: Exit Sub
    End If
  Next i
  If Row = -1 Then
    Row = mRowCount - 1
    If Row < 0 Then Row = 0
    RowPos = Abs(RowsCharCount(Row))
  End If
End Sub

Private Sub CalcCaretPos()
Dim i&, CPos&, CCount&, CrLfOffs&
  For i = 0 To CurRow - 1
    CCount = RowsCharCount(i)
    If CCount > 0 Then CrLfOffs = 0 Else CrLfOffs = 2: CCount = -CCount 'vbCrLf-detection
    CPos = CPos + CCount + CrLfOffs
    If i >= mRowCount - 1 Then Exit For
  Next i
  If CurPos > Abs(RowsCharCount(i)) Then
    CaretPos = CPos + Abs(RowsCharCount(i))
  Else
    CaretPos = CPos + CurPos
  End If
  If CaretPos > Len(mText) Then CaretPos = Len(mText)
End Sub

Private Function SetCurFont(Optional MCC As cCairoContext) As cCairoContext
  If MCC Is Nothing Then Set MCC = W.MeasureContext

  If Not InsideDraw Then
    MCC.Save
    MCC.ScaleDrawings Zoom, Zoom
  End If
  
  If Len(mPassWordChar) Then
    MCC.SelectFont "Courier", W.FontSize, W.ForeColor, W.FontBold, W.FontItalic, W.FontUnderline
  Else
    W.SelectFontSettingsInto MCC
  End If
  Set SetCurFont = MCC
End Function

